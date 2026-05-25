"""
╔══════════════════════════════════════════════════════════════════╗
║         COPA BRACELL ABSOLUTO — APURAÇÃO AUTOMÁTICA VIA SQL      ║
╚══════════════════════════════════════════════════════════════════╝

Fluxo:
  SQL (banco de dados)
    → apuracao_sql.py  (este script)
      → controle_copa_bracell.xlsx  (atualiza células SIM/NÃO)
        → dados.json                (alimenta o álbum HTML)

Dependências:
  pip install openpyxl pyodbc sqlalchemy python-dotenv
  # Para SQL Server:  pip install pyodbc
  # Para MySQL:       pip install pymysql
  # Para PostgreSQL:  pip install psycopg2-binary

Uso:
  python apuracao_sql.py               # roda apuração completa
  python apuracao_sql.py --dry-run     # mostra resultado sem salvar
  python apuracao_sql.py --equipe franca  # só uma equipe
"""

import json
import os
import sys
import argparse
from datetime import datetime
from pathlib import Path

# ── Carrega .env se existir (para esconder senha) ─────────────────
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass


# ══════════════════════════════════════════════════════════════════
#  1. CONFIGURAÇÃO DO BANCO DE DADOS
#  Edite aqui com suas credenciais
# ══════════════════════════════════════════════════════════════════

DB_CONFIG = {
    # Tipo: "sqlserver" | "mysql" | "postgresql" | "sqlite"
    "tipo": "sqlserver",

    # SQL Server (pyodbc)
    "sqlserver": {
        "server":   os.getenv("DB_SERVER",   "SEU_SERVIDOR"),
        "database": os.getenv("DB_NAME",     "SEU_BANCO"),
        "username": os.getenv("DB_USER",     "seu_usuario"),
        "password": os.getenv("DB_PASSWORD", "sua_senha"),
        # Driver ODBC instalado na máquina (veja: odbcinst -q -d)
        "driver":   os.getenv("DB_DRIVER",   "ODBC Driver 17 for SQL Server"),
    }}

# ══════════════════════════════════════════════════════════════════
#  2. CAMINHOS DOS ARQUIVOS
# ══════════════════════════════════════════════════════════════════

PASTA_PROJETO  = Path(__file__).parent
ARQUIVO_EXCEL  = PASTA_PROJETO / "controle_copa_bracell.xlsx"
ARQUIVO_JSON   = PASTA_PROJETO / "dados.json"
ABA_CONTROLE   = "Controle"

LINHA_HEADER   = 2
LINHA_INICIO   = 3
COL_FIGURINHA_INICIO = 5   # E
COL_ID_EQUIPE  = 4


# ══════════════════════════════════════════════════════════════════
#  3. MAPEAMENTO: ID_EQUIPE → QUERIES SQL
#
#  Cada figurinha tem uma query que retorna 1 linha com 1 coluna:
#    - valor > 0  ou True  → meta ATINGIDA → "SIM"
#    - valor = 0  ou False → não atingida  → ""  (célula vazia)
#
#  O parâmetro :equipe_id (ou %(equipe_id)s para MySQL) será
#  substituído pelo ID da equipe (ex: "franca", "alemanha").
#
#  ADAPTE as queries para refletir sua estrutura de banco.
# ══════════════════════════════════════════════════════════════════

# IDs das equipes (devem bater com coluna D da planilha)
EQUIPES = [
    "alemanha", "belgica", "canada", "espanha", "franca",
    "holanda", "inglaterra", "japao", "mexico", "portugal", "uruguai",
]

# Cada entrada é (id_figurinha, query_sql)
# A query deve retornar um único valor numérico ou booleano
QUERIES_FIGURINHAS = [
    # ── 01 — 90% de Efetividade ──────────────────────────────────
    ("efetividade_90", """
        SELECT CASE WHEN AVG(CAST(efetividade AS FLOAT)) >= 0.90 THEN 1 ELSE 0 END
        FROM   apuracao_promotores
        WHERE  equipe_id = :equipe_id
          AND  periodo   = :periodo
    """),

    # ── 02 — 50% Promotores c/ Ponto Extra ───────────────────────
    ("pe_50pct_prom", """
        SELECT CASE
            WHEN CAST(SUM(CASE WHEN tem_ponto_extra = 1 THEN 1 ELSE 0 END) AS FLOAT)
                 / NULLIF(COUNT(*), 0) >= 0.50 THEN 1 ELSE 0 END
        FROM   apuracao_promotores
        WHERE  equipe_id = :equipe_id
          AND  periodo   = :periodo
    """),

    # ── 03 — 100% Promotores c/ Ponto Extra ──────────────────────
    ("pe_100pct_prom", """
        SELECT CASE
            WHEN CAST(SUM(CASE WHEN tem_ponto_extra = 1 THEN 1 ELSE 0 END) AS FLOAT)
                 / NULLIF(COUNT(*), 0) >= 1.00 THEN 1 ELSE 0 END
        FROM   apuracao_promotores
        WHERE  equipe_id = :equipe_id
          AND  periodo   = :periodo
    """),

    # ── 04 — 50% Lojas com Cross Selling ─────────────────────────
    ("cross_50pct", """
        SELECT CASE
            WHEN CAST(SUM(CASE WHEN tem_cross = 1 THEN 1 ELSE 0 END) AS FLOAT)
                 / NULLIF(COUNT(*), 0) >= 0.50 THEN 1 ELSE 0 END
        FROM   apuracao_lojas
        WHERE  equipe_id = :equipe_id
          AND  periodo   = :periodo
    """),

    # ── 05 — 20% Lojas com Terminal ──────────────────────────────
    ("terminal_20pct", """
        SELECT CASE
            WHEN CAST(SUM(CASE WHEN tem_terminal = 1 THEN 1 ELSE 0 END) AS FLOAT)
                 / NULLIF(COUNT(*), 0) >= 0.20 THEN 1 ELSE 0 END
        FROM   apuracao_lojas
        WHERE  equipe_id = :equipe_id
          AND  periodo   = :periodo
    """),

    # ── 06 — 20% Lojas com Ilha ───────────────────────────────────
    ("ilha_20pct", """
        SELECT CASE
            WHEN CAST(SUM(CASE WHEN tem_ilha = 1 THEN 1 ELSE 0 END) AS FLOAT)
                 / NULLIF(COUNT(*), 0) >= 0.20 THEN 1 ELSE 0 END
        FROM   apuracao_lojas
        WHERE  equipe_id = :equipe_id
          AND  periodo   = :periodo
    """),

    # ── 07 — 2 Lojas Display/Turbilhão ───────────────────────────
    ("display_turbilhao", """
        SELECT CASE WHEN COUNT(*) >= 2 THEN 1 ELSE 0 END
        FROM   apuracao_lojas
        WHERE  equipe_id   = :equipe_id
          AND  periodo     = :periodo
          AND  tem_display = 1
    """),

    # ── 08 — 2 Lojas PE Criativo Copa ────────────────────────────
    ("criativo_copa_2", """
        SELECT CASE WHEN COUNT(*) >= 2 THEN 1 ELSE 0 END
        FROM   apuracao_lojas
        WHERE  equipe_id        = :equipe_id
          AND  periodo          = :periodo
          AND  tem_pe_copa      = 1
    """),

    # ── 09 — 4 Lojas PE Criativo Copa ────────────────────────────
    ("criativo_copa_4", """
        SELECT CASE WHEN COUNT(*) >= 4 THEN 1 ELSE 0 END
        FROM   apuracao_lojas
        WHERE  equipe_id   = :equipe_id
          AND  periodo     = :periodo
          AND  tem_pe_copa = 1
    """),

    # ── 10 — 2 Lojas PE Criativo São João ────────────────────────
    ("criativo_sj_2", """
        SELECT CASE WHEN COUNT(*) >= 2 THEN 1 ELSE 0 END
        FROM   apuracao_lojas
        WHERE  equipe_id     = :equipe_id
          AND  periodo       = :periodo
          AND  tem_pe_sjoao  = 1
    """),

    # ── 11 — 4 Lojas PE Criativo São João ────────────────────────
    ("criativo_sj_4", """
        SELECT CASE WHEN COUNT(*) >= 4 THEN 1 ELSE 0 END
        FROM   apuracao_lojas
        WHERE  equipe_id    = :equipe_id
          AND  periodo      = :periodo
          AND  tem_pe_sjoao = 1
    """),

    # ── 12 — 25% Lojas MPDV Leve ─────────────────────────────────
    ("mpdv_leve_25", """
        SELECT CASE
            WHEN CAST(SUM(CASE WHEN tem_mpdv_leve = 1 THEN 1 ELSE 0 END) AS FLOAT)
                 / NULLIF(COUNT(*), 0) >= 0.25 THEN 1 ELSE 0 END
        FROM   apuracao_lojas
        WHERE  equipe_id = :equipe_id
          AND  periodo   = :periodo
    """),

    # ── 13 — 50% Lojas MPDV Leve ─────────────────────────────────
    ("mpdv_leve_50", """
        SELECT CASE
            WHEN CAST(SUM(CASE WHEN tem_mpdv_leve = 1 THEN 1 ELSE 0 END) AS FLOAT)
                 / NULLIF(COUNT(*), 0) >= 0.50 THEN 1 ELSE 0 END
        FROM   apuracao_lojas
        WHERE  equipe_id = :equipe_id
          AND  periodo   = :periodo
    """),

    # ── 14 — 1 Loja MPDV Pesado ──────────────────────────────────
    ("mpdv_pesado_1", """
        SELECT CASE WHEN COUNT(*) >= 1 THEN 1 ELSE 0 END
        FROM   apuracao_lojas
        WHERE  equipe_id      = :equipe_id
          AND  periodo        = :periodo
          AND  tem_mpdv_pesado = 1
    """),

    # ── 15 — 2 Lojas MPDV Pesado ─────────────────────────────────
    ("mpdv_pesado_2", """
        SELECT CASE WHEN COUNT(*) >= 2 THEN 1 ELSE 0 END
        FROM   apuracao_lojas
        WHERE  equipe_id       = :equipe_id
          AND  periodo         = :periodo
          AND  tem_mpdv_pesado = 1
    """)
]

# Período atual da apuração (ajuste conforme necessário)
PERIODO_ATUAL = os.getenv("PERIODO_APURACAO", "2026-06")


# ══════════════════════════════════════════════════════════════════
#  4. CONEXÃO COM O BANCO
# ══════════════════════════════════════════════════════════════════

def criar_conexao():
    """Cria e retorna uma conexão com o banco configurado."""
    tipo = DB_CONFIG["tipo"]

    if tipo == "sqlserver":
        import pyodbc
        cfg = DB_CONFIG["sqlserver"]
        conn_str = (
            f"DRIVER={{{cfg['driver']}}};"
            f"SERVER={cfg['server']};"
            f"DATABASE={cfg['database']};"
            f"UID={cfg['username']};"
            f"PWD={cfg['password']};"
            "Encrypt=yes;TrustServerCertificate=yes;"
        )
        return pyodbc.connect(conn_str)

    elif tipo == "mysql":
        import pymysql
        cfg = DB_CONFIG["mysql"]
        return pymysql.connect(
            host=cfg["host"], port=cfg["port"],
            db=cfg["database"],
            user=cfg["username"], password=cfg["password"],
            charset="utf8mb4",
        )

    elif tipo == "postgresql":
        import psycopg2
        cfg = DB_CONFIG["postgresql"]
        return psycopg2.connect(
            host=cfg["host"], port=cfg["port"],
            dbname=cfg["database"],
            user=cfg["username"], password=cfg["password"],
        )

    elif tipo == "sqlite":
        import sqlite3
        return sqlite3.connect(DB_CONFIG["sqlite"]["arquivo"])

    else:
        raise ValueError(f"Tipo de banco desconhecido: {tipo}")


def executar_query(cursor, sql, equipe_id, periodo):
    """Executa uma query e retorna True se a meta foi atingida."""
    tipo = DB_CONFIG["tipo"]
    try:
        # Adapta placeholder para cada banco
        if tipo in ("mysql",):
            sql_adapted = sql.replace(":equipe_id", "%(equipe_id)s").replace(":periodo", "%(periodo)s")
            cursor.execute(sql_adapted, {"equipe_id": equipe_id, "periodo": periodo})
        else:
            cursor.execute(sql, {"equipe_id": equipe_id, "periodo": periodo})

        row = cursor.fetchone()
        if row is None:
            return False
        valor = row[0]
        return bool(valor) and valor not in (0, "0", False, None)
    except Exception as e:
        print(f"  ⚠️  Erro na query para {equipe_id}: {e}")
        return False


# ══════════════════════════════════════════════════════════════════
#  5. ATUALIZAÇÃO DO EXCEL
# ══════════════════════════════════════════════════════════════════

def atualizar_excel(resultados: dict, dry_run: bool = False):
    """
    Atualiza a planilha com os resultados.
    resultados = { "franca": {"efetividade_90": True, "pe_50pct_prom": False, ...}, ... }
    """
    from openpyxl import load_workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter

    print(f"\n📊 {'[DRY-RUN] ' if dry_run else ''}Atualizando planilha: {ARQUIVO_EXCEL.name}")

    wb = load_workbook(ARQUIVO_EXCEL)
    ws = wb[ABA_CONTROLE]

    # Cores
    VERDE  = PatternFill("solid", fgColor="C6EFCE")  # verde claro = atingiu
    BRANCO = PatternFill("solid", fgColor="FFFFFF")   # branco = não atingiu
    FONTE_SIM = Font(bold=True, color="276221")
    FONTE_NAO = Font(bold=False, color="999999")

    # IDs de figurinha na ordem das colunas (E, F, G... = col 5, 6, 7...)
    ids_figurinha = [fid for fid, _ in QUERIES_FIGURINHAS]

    alteracoes = 0

    for linha in ws.iter_rows(min_row=LINHA_INICIO):
        # Lê o ID da equipe na coluna D
        cell_id = linha[COL_ID_EQUIPE - 1]
        equipe_id = str(cell_id.value or "").strip().lower()

        if equipe_id not in resultados:
            continue

        equipe_resultado = resultados[equipe_id]
        conquistadas = 0

        for i, fig_id in enumerate(ids_figurinha):
            col_idx = COL_FIGURINHA_INICIO + i - 1  # 0-based index na linha
            cell = linha[col_idx]
            atingiu = equipe_resultado.get(fig_id, False)

            novo_valor = "SIM" if atingiu else ""
            valor_atual = str(cell.value or "").strip().upper()

            if novo_valor != valor_atual:
                alteracoes += 1
                status = "SIM ✅" if atingiu else "removido"
                print(f"  {equipe_id:12s} | #{i+1:02d} {fig_id:25s} → {status}")

            if not dry_run:
                cell.value = novo_valor
                if atingiu:
                    cell.fill = VERDE
                    cell.font = FONTE_SIM
                    cell.alignment = Alignment(horizontal="center")
                    conquistadas += 1
                else:
                    cell.fill = BRANCO
                    cell.font = FONTE_NAO
                    cell.alignment = Alignment(horizontal="center")
            else:
                if atingiu:
                    conquistadas += 1

        # Atualiza coluna PROGRESSO (última coluna de dados)
        col_progresso = COL_FIGURINHA_INICIO + len(ids_figurinha) - 1
        cell_prog = linha[col_progresso]
        if not dry_run:
            cell_prog.value = f"{conquistadas}/16"

    if dry_run:
        print(f"\n  Total de alterações que seriam feitas: {alteracoes}")
        print("  (modo dry-run — nenhum arquivo foi modificado)")
        return

    print(f"\n  ✅ {alteracoes} alterações realizadas.")

    # Atualiza timestamp na célula A1 ou B1
    try:
        ws["A2"] = (f"COPA BRACELL ABSOLUTO — CONTROLE DE FIGURINHAS 2026  "
                    f"| Última atualização: {datetime.now().strftime('%d/%m/%Y %H:%M')}")
    except Exception:
        pass

    wb.save(ARQUIVO_EXCEL)
    print(f"  💾 Planilha salva: {ARQUIVO_EXCEL}")


# ══════════════════════════════════════════════════════════════════
#  6. GERAÇÃO DO dados.json  (igual ao gerar_dados.py original)
# ══════════════════════════════════════════════════════════════════

def gerar_json(resultados: dict):
    """Gera o dados.json a partir dos resultados da apuração."""
    dados = {"equipes": {}, "gerado_em": datetime.now().isoformat()}

    for equipe_id, figurinhas in resultados.items():
        dados["equipes"][equipe_id] = {
            fig_id: True
            for fig_id, atingiu in figurinhas.items()
            if atingiu
        }

    ARQUIVO_JSON.write_text(
        json.dumps(dados, ensure_ascii=False, indent=2),
        encoding="utf-8"
    )
    print(f"  🎴 dados.json gerado: {ARQUIVO_JSON}")


# ══════════════════════════════════════════════════════════════════
#  7. ORQUESTRADOR PRINCIPAL
# ══════════════════════════════════════════════════════════════════

def apurar(equipes_filtro=None, dry_run=False):
    equipes = equipes_filtro or EQUIPES
    resultados = {}

    print(f"\n🔌 Conectando ao banco ({DB_CONFIG['tipo']})...")
    try:
        conn = criar_conexao()
    except Exception as e:
        print(f"❌ Falha na conexão: {e}")
        print("\nVerifique as credenciais em DB_CONFIG ou no arquivo .env")
        sys.exit(1)

    cursor = conn.cursor()
    print(f"✅ Conectado! Período: {PERIODO_ATUAL}\n")

    for equipe_id in equipes:
        print(f"🏳  Apurando equipe: {equipe_id}")
        resultados[equipe_id] = {}

        for fig_id, query in QUERIES_FIGURINHAS:
            atingiu = executar_query(cursor, query, equipe_id, PERIODO_ATUAL)
            resultados[equipe_id][fig_id] = atingiu
            simbolo = "✅" if atingiu else "·"
            print(f"   {simbolo} {fig_id}")

    cursor.close()
    conn.close()
    print("\n🔒 Conexão encerrada.")

    # Atualiza Excel
    atualizar_excel(resultados, dry_run=dry_run)

    # Gera JSON (só se não for dry-run)
    if not dry_run:
        print("\n📄 Gerando dados.json...")
        gerar_json(resultados)

    # Resumo
    print("\n" + "═" * 60)
    print("RESUMO DA APURAÇÃO")
    print("═" * 60)
    for eq, figs in resultados.items():
        total = sum(1 for v in figs.values() if v)
        print(f"  {eq:15s}  {total:2d}/16  {'█' * total}{'░' * (16-total)}")
    print("═" * 60)
    print(f"Apuração concluída: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")

