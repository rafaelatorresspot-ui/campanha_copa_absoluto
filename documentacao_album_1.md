# Documentação — Álbum Copa Bracell Absoluto
**Arquivo:** `index.html` · **Tamanho:** ~328 KB · **Gerado em:** Maio 2026

---

## Índice
1. [Visão Geral da Estrutura](#visão-geral)
2. [Seção CSS](#seção-css)
3. [Seção HTML — Capa](#seção-capa)
4. [Seção HTML — Álbum](#seção-álbum)
5. [Seção JavaScript — Dados](#javascript-dados)
6. [Seção JavaScript — Funções](#javascript-funções)
7. [Fluxo de Dados](#fluxo-de-dados)
8. [Estrutura de Arquivos do Projeto](#estrutura-de-arquivos)
9. [Como Adicionar Figurinhas](#como-adicionar-figurinhas)
10. [Como Alterar Dados](#como-alterar-dados)

---

## Visão Geral

O arquivo é um **SPA (Single Page Application)** em HTML/CSS/JS puro, sem frameworks.
Ele tem 4 grandes regiões:

```
┌─────────────────────────────────────────────────────────────┐
│  <style>          CSS global          chars ~387–21.345     │
├─────────────────────────────────────────────────────────────┤
│  <div id="cover"> Tela de capa        chars ~21.370–158.086 │
├─────────────────────────────────────────────────────────────┤
│  <div id="album"> Álbum interativo    chars ~158.086–207.565│
├─────────────────────────────────────────────────────────────┤
│  <script>         Toda a lógica JS    chars ~207.619–328.369│
└─────────────────────────────────────────────────────────────┘
```

A **capa** é estática (HTML puro). O **álbum** é gerado dinamicamente pelo JavaScript com base nos arrays `TEAMS`, `STICKERS`, `PANINI` e no arquivo externo `dados.json`.

---

## Seção CSS

### Grupos de classes

| Prefixo | Região | Exemplos |
|---------|--------|---------|
| `.cv-*` | Capa — conteúdo central | `.cv-book`, `.cv-title`, `.cv-cta`, `.cv-prog` |
| `.bk-*` | Capa — livro físico | `.bk-cover`, `.bk-spine`, `.bk-logo-b`, `.bk-footer` |
| `.hex` | Capa — hexágonos animados de fundo | `.hex:nth-child(n)` |
| `.al-nav` | Barra de navegação do álbum | `.al-nav`, `.nav-tab`, `.nav-back` |
| `.t-*` | Página de equipe — hero/strip | `.t-hero`, `.t-strip`, `.t-crest`, `.t-name`, `.t-sup` |
| `.tpb*` | Barra de progresso da equipe | `.tpb-bar`, `.tpb-fill`, `.tpb-g` |
| `.prizes*` `.pz*` | Painel de premiação | `.prizes-grid`, `.pz`, `.pz-val`, `.pz-tag` |
| `.sbar` | Barra de status (dados.json) | `.sbar`, `.sdot`, `.stxt`, `.srefresh` |
| `.album-spread` | Dupla página estilo álbum | `.album-spread`, `.panini-page` |
| `.panini-*` | Componentes das páginas Panini | `.panini-header`, `.panini-slots-grid`, `.panini-slot` |
| `.sc-*` | Partes internas da figurinha | `.sc-bar-panini`, `.sc-img-wrap-panini`, `.sc-lbl-panini` |
| `.afn-*` | Navegação inferior (Anterior/Próxima) | `.afn-btn`, `.afn-page`, `.afn-count` |
| `.pf*` | Rodapé da página | `.pf`, `.pf-copy` |

### Classes de estado

| Classe | Descrição |
|--------|-----------|
| `.panini-slot.earned` | Figurinha conquistada — imagem visível, barra/código ocultos |
| `.panini-slot.rare` | Raridade rara — estrela roxa |
| `.panini-slot.epic` | Raridade épica — estrela laranja |
| `.panini-slot.legendary` | Raridade lendária — borda dourada + efeito holográfico |
| `.nav-tab.active` | Aba de equipe selecionada |
| `.nav-tab.has-s` | Equipe com pelo menos 1 figurinha |
| `.tp.active` | Página de equipe visível |
| `.page-complete-badge.show` | Selo "Álbum Completo" visível |

### Breakpoints responsivos

| Breakpoint | Comportamento |
|------------|---------------|
| `≤ 1200px` | Padding e gap reduzidos, header compacto |
| `≤ 1024px` | Spread passa de 2 colunas para 1 coluna (páginas empilhadas) |
| `≤ 600px` | Layout mobile: fontes menores, capa reduzida a 82%, nav compacta |
| `≤ 380px` | Grid de figurinhas muda de 4 para 2 colunas |

---

## Seção Capa

**Localização:** `<div id="cover">`

### Subcomponentes

```
#cover
├── .cvbg              → fundo com gradiente e grade pontilhada (pointer-events:none)
├── .hex (×6)          → hexágonos flutuantes animados (animação floatHex)
└── .cv-inner          → container central (z-index:10)
    ├── .cv-book        → livro físico 3D
    │   ├── .bk-spine   → lombada lateral esquerda
    │   └── .bk-cover   → capa do livro
    │       └── <img>   → PNG da capa real (base64 embutido)
    ├── .cv-title       → "Copa / Bracell" (Bebas Neue)
    ├── .cv-sub         → subtítulo da campanha
    ├── .cv-prog        → barra de progresso global
    ├── #cv-teams       → pills de equipes (gerado por buildCoverPills())
    └── .cv-cta         → botão "Abrir Álbum"
```

> **Logos embutidos:** Os logos Bracell e SPOT estão em `data:image/png;base64,...` diretamente no HTML — não dependem de arquivos externos.

---

## Seção Álbum

**Localização:** `<div id="album">` — inicialmente `display:none`, exibido por `openAlbum()`

### Subcomponentes

```
#album
├── .al-nav (sticky)           → barra de navegação fixa no topo
│   ├── .nav-back              → botão "← Capa"
│   ├── #nav-tabs              → tabs das 11 equipes (gerado por buildNavTabs())
│   └── .nav-logos             → logos Bracell + SPOT
└── #team-pages                → container de todas as páginas (gerado por buildPages())
    └── .tp (×11)              → página de cada equipe (gerado por buildPage())
        ├── .t-hero            → hero com cor da seleção
        │   ├── .t-strip       → faixa colorida com bandeira + info
        │   └── .t-bbar        → faixa inferior com logos
        ├── .prizes            → 4 cards de premiação
        ├── .sbar              → barra de status + botão Atualizar
        └── .album-spread      → dupla página estilo Panini
            ├── .panini-page.left   → página esquerda (figurinhas 01–08)
            │   ├── .panini-header  → "WE ARE [PAÍS]" + bandeira + federação
            │   └── .panini-slots-grid → grid 4×2 com os slots
            │       └── .panini-slot (×8) → cada figurinha
            ├── .panini-page.right  → página direita (figurinhas 09–16)
            │   ├── .panini-header  → "ABSOLUTO / COPA BRACELL" alinhado à direita
            │   └── .panini-slots-grid → grid 4×2
            │       └── .panini-slot (×8)
            └── .page-complete-badge → selo "Álbum Completo" (oculto por padrão)
```

### Anatomia de um slot de figurinha

```
.panini-slot                    → célula do grid (aspect-ratio 3/4)
└── .panini-slot-inner          → preenche position:absolute;inset:0

  [VAZIA — não earned]
    ├── .panini-slot-code       → código ex: "ALE-01"
    ├── .panini-slot-num-bg     → número watermark (ex: "01")
    └── label do objetivo       → texto da meta

  [CONQUISTADA — .earned]
    ├── .sc-bar-panini          → barra colorida c/ número (OCULTA com earned)
    ├── .sc-img-wrap-panini     → área da imagem
    │   └── <img>               → assets/stickers/{team_id}/{sticker_id}.png
    ├── .sc-lbl-wrap-panini     → label do objetivo (OCULTO com earned)
    ├── .sc-panini-code-bot     → código inferior (OCULTO com earned)
    └── .sc-star-panini         → estrela dourada (visível com earned)
```

---

## JavaScript — Dados

### Constantes globais

#### `TEAMS` — Array com 11 equipes
```js
{ id, name, color, supervisor }
// id       → chave usada em todos os IDs do DOM e no dados.json
// color    → cor hex usada na barra hero e na barra do slot
// supervisor → nome exibido na página
```

#### `STICKERS` — Array com 16 figurinhas
```js
{ id, label, num, code, icon, rarity }
// id     → nome do arquivo PNG (sem extensão)
// num    → "01"–"16" (exibido no watermark)
// code   → "ABS-01"–"ABS-16" (código Panini)
// rarity → 0=comum, 1=rara, 2=épica, 3=lendária
```

| # | id | label | rarity |
|---|----|-------|--------|
| 01 | `efetividade_90` | 90% de Efetividade | rara |
| 02 | `pe_50pct_prom` | 50% Promotores c/ Ponto Extra | comum |
| 03 | `pe_100pct_prom` | 100% Promotores c/ P.Extra | épica |
| 04 | `cross_50pct` | 50% Lojas com Cross | comum |
| 05 | `terminal_20pct` | 20% Lojas com Terminal | comum |
| 06 | `ilha_20pct` | 20% Lojas com Ilha | rara |
| 07 | `display_turbilhao` | 2 Lojas Display/Turbilhão | comum |
| 08 | `criativo_copa_2` | 2 Lojas PE Criativo Copa | rara |
| 09 | `criativo_copa_4` | 4 Lojas PE Criativo Copa | épica |
| 10 | `criativo_sj_2` | 2 Lojas PE Criativo S.João | rara |
| 11 | `criativo_sj_4` | 4 Lojas PE Criativo S.João | épica |
| 12 | `mpdv_leve_25` | 25% Lojas MPDV Leve | comum |
| 13 | `mpdv_leve_50` | 50% Lojas MPDV Leve | rara |
| 14 | `mpdv_pesado_1` | 1 Loja MPDV Pesado | rara |
| 15 | `mpdv_pesado_2` | 2 Lojas MPDV Pesado | épica |
| 16 | `autofalante` | Anúncio no Autofalante | lendária |

#### `PANINI` — Temas visuais por equipe
```js
{
  b1, b2, b3,   // cores dos blobs SVG de fundo (bandeira)
  wa,           // cor do texto "WE ARE"
  cn,           // cor do nome do país
  sA, sB,       // cores alternadas dos slots vazios
  st,           // cor do texto nos slots vazios
  fed           // nome oficial da federação
}
```

#### `DADOS` — Estado das figurinhas (carregado de `dados.json`)
```js
{
  "equipes": {
    "mexico": {
      "efetividade_90": true,
      "pe_50pct_prom": false,
      // ... 16 chaves por equipe
    },
    // ... 11 equipes
  }
}
```

#### Variáveis de controle
```js
let DADOS = {}        // dados carregados do JSON
let imgTs = Date.now() // timestamp para cache-busting das imagens
let currentTeam = null // equipe atualmente selecionada
```

---

## JavaScript — Funções

### Carregamento de dados

| Função | Descrição |
|--------|-----------|
| `carregarDados()` | `async` — faz `fetch('dados.json?t=timestamp')`, popula `DADOS`. Se protocolo for `file://`, exibe aviso |
| `atualizarStatusBar(msg)` | Atualiza o texto em todos os elementos `.stxt` da página |

### Construção do DOM

| Função | Descrição |
|--------|-----------|
| `buildAll()` | Ponto de entrada — chama as 4 funções de build + refreshUI |
| `buildCoverPills()` | Gera os pills de equipe (`#cv-teams`) na capa com bandeira + % |
| `buildNavTabs()` | Gera as tabs (`#nav-tabs`) na barra de navegação do álbum |
| `buildPages()` | Gera todas as 11 páginas de equipe dentro de `#team-pages` |
| `buildPage(t, index)` | Gera o HTML completo de uma página de equipe (hero + prizes + spread) |
| `buildPaniniSlot(t, s, slotColor)` | Gera o HTML de um slot individual — vazio ou conquistado |

### Utilitários visuais

| Função | Descrição |
|--------|-----------|
| `blobSvg(p)` | Retorna o SVG das formas orgânicas de fundo usando as cores do tema `p` |
| `darken(hex, amt)` | Escurece uma cor hex em `amt` unidades RGB — usado no gradiente do hero |
| `launchConfetti()` | Lança 60 peças de confete animadas ao completar um álbum |

### Estado e atualização

| Função | Descrição |
|--------|-----------|
| `refreshUI()` | Atualiza **sem reconstruir o DOM** — percentuais, barras, classes `.earned`, badges |
| `isEarned(tid, sid)` | `true` se `DADOS.equipes[tid][sid] === true` |
| `teamProg(tid)` | Retorna `{earned, total, pct}` para uma equipe |
| `globalProg()` | Retorna `{earned, total, pct}` para o álbum inteiro |
| `recarregar()` | `async` — recarrega `dados.json` + reconstrói páginas + atualiza `src` das imagens com novo timestamp |

### Navegação

| Função | Descrição |
|--------|-----------|
| `openAlbum(teamId)` | Oculta `#cover`, exibe `#album`, chama `switchTeam()` |
| `goToCover()` | Exibe `#cover`, oculta `#album`, limpa `currentTeam` |
| `switchTeam(tid)` | Ativa a página `.tp` e tab correspondente, scroll para o topo |
| `navPage(tid, dir)` | Navega para equipe anterior (`dir=-1`) ou próxima (`dir=1`) |
| `updateNavBtns()` | Desabilita botão Anterior na 1ª equipe e Próxima na última |

---

## Fluxo de Dados

```
dados.json
    │
    ▼
carregarDados()        → popula DADOS{}
    │
    ▼
buildAll()
    ├── buildCoverPills()  → DOM: #cv-teams
    ├── buildNavTabs()     → DOM: #nav-tabs
    ├── buildPages()       → DOM: #team-pages
    │       └── buildPage() × 11
    │               └── buildPaniniSlot() × 16 por equipe
    └── refreshUI()        → atualiza %, barras, classes
```

```
Usuário clica ↻ Atualizar
    │
    ▼
recarregar()
    ├── carregarDados()   → recarrega dados.json
    ├── imgTs = Date.now() → novo timestamp
    ├── buildPages()       → reconstrói DOM
    ├── refreshUI()        → atualiza estado visual
    └── atualiza src das <img> com ?t=timestamp  → força recarregamento do cache
```

---

## Estrutura de Arquivos do Projeto

```
copa-bracell/
├── index.html                     ← este arquivo
├── dados.json                     ← estado das figurinhas {equipes:{team:{sticker:bool}}}
├── gerar_dados.py                 ← converte planilha → dados.json
├── controle_copa_bracell.xlsx     ← planilha de controle admin
├── README.md                      ← instruções de uso
└── assets/
    └── stickers/
        ├── mexico/
        │   ├── efetividade_90.png
        │   ├── pe_50pct_prom.png
        │   └── ... (16 PNGs)
        ├── canada/  (mesmos 16 nomes)
        ├── belgica/
        ├── alemanha/
        ├── holanda/
        ├── uruguai/
        ├── espanha/
        ├── franca/
        ├── japao/
        ├── portugal/
        └── inglaterra/
```

> **Caminho das imagens:** `assets/stickers/{team_id}/{sticker_id}.png`  
> Exemplo: `assets/stickers/alemanha/efetividade_90.png`

---

## Como Adicionar Figurinhas

Para que uma figurinha apareça colorida no álbum:

**Opção A — Via planilha (recomendado):**
1. Abrir `controle_copa_bracell.xlsx` → aba **Controle**
2. Digitar `SIM` na célula da equipe × figurinha
3. Salvar e rodar: `python3 gerar_dados.py`
4. Fazer commit do `dados.json` ou clicar **↻ Atualizar** no álbum

**Opção B — Editar `dados.json` diretamente:**
```json
{
  "equipes": {
    "alemanha": {
      "efetividade_90": true,   ← true = figurinha aparece
      "pe_50pct_prom": false    ← false = slot vazio
    }
  }
}
```

**Opção C — Via banco de dados:**
Script `query_to_excel.py` lê query SQL → escreve SIM na planilha → `gerar_dados.py` gera o JSON.

---

## Como Alterar Dados

### Adicionar/renomear equipe
```js
// Em TEAMS (linha ~207.800)
{ id:'nova_equipe', name:'Nova Equipe', color:'#RRGGBB', supervisor:'Nome' }

// Em PANINI (linha ~207.900) — adicionar tema visual
nova_equipe: { b1:'#cor1', b2:'#cor2', b3:'#cor3', wa:'#cor', cn:'#cor',
               sA:'#cor', sB:'#cor', st:'#cor', fed:'Federação X' }

// Em FLAG_CODES — código ISO da bandeira
nova_equipe: 'xx'   // código flagcdn.com
```

### Adicionar figurinha
```js
// Em STICKERS (linha ~207.840)
{ id:'novo_sticker', label:'Descrição', num:'17', code:'ABS-17', icon:'🎯', rarity:0 }
```

### Alterar supervisor
```js
// Em TEAMS, campo supervisor
{ id:'holanda', ..., supervisor:'Anaive Cristina' }
```

### Alterar tema visual de uma equipe
```js
// Em PANINI, objeto da equipe
holanda: {
  b1:'#AE1C28',  // blob principal (vermelho holandês)
  b2:'#21468B',  // blob secundário (azul)
  b3:'#F77F00',  // círculo acento (laranja)
  wa:'#AE1C28',  // cor "WE ARE"
  cn:'#F77F00',  // cor nome do país
  sA:'#FFE0B2',  // cor slot A (pastel laranja)
  sB:'#FFCC80',  // cor slot B
  st:'#E65100',  // cor texto slots vazios
  fed:'Koninklijke Nederlandse Voetbalbond'
}
```

---

*Documentação gerada automaticamente — Copa Bracell Absoluto 2026 · Bracell × SPOT*
