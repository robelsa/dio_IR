# 📐 Documentação Técnica — Agregador de Imposto de Renda

Este documento descreve **como a planilha foi construída** e **por que** cada decisão foi tomada. Ele complementa o `README.md`, que foca no uso.

---

## 1. Arquitetura geral

A pasta de trabalho segue o princípio de **separação de responsabilidades**, dividindo as abas em três camadas:

```
┌─────────────────────────────────────────────────────────────┐
│  CAMADA DE APRESENTAÇÃO / NAVEGAÇÃO                          │
│  • INÍCIO  (menu, instruções, links rápidos)                │
├─────────────────────────────────────────────────────────────┤
│  CAMADA DE ENTRADA DE DADOS                                  │
│  • DADOS PESSOAIS · RENDIMENTOS · BENS E DIREITOS            │
│  • DÍVIDAS E ÔNUS · PAGAMENTOS · DEPENDENTES · DOCUMENTOS    │
├─────────────────────────────────────────────────────────────┤
│  CAMADA DE APOIO E SAÍDA                                     │
│  • LISTAS  (fonte das validações)  →  alimenta as entradas   │
│  • RESUMO  (dashboard)             ←  consome as entradas    │
└─────────────────────────────────────────────────────────────┘
```

**Por quê?** Essa separação torna a planilha sustentável: alterar uma opção de lista (aba **LISTAS**) reflete em todas as validações; o dashboard (**RESUMO**) apenas lê as abas de entrada, sem duplicar dados. A entrada nunca depende da saída, evitando referências circulares.

---

## 2. Abas em detalhe

### 2.1 INÍCIO
- Capa com título, descrição e **instruções numeradas**.
- **Menu de navegação**: 8 botões com **hyperlinks internos** (`'ABA'!A1`).
- **Links rápidos**: hyperlinks externos para e-CAC, Meu Imposto de Renda, consulta de CPF e perguntas & respostas.
- Rodapé com versão, autoria e aviso de dados fictícios.
- **Decisão:** navegação por hyperlinks (e não por botões de macro) para manter o arquivo **100% compatível** e livre de VBA — abre em qualquer Excel, LibreOffice ou Google Sheets sem alertas de segurança.

### 2.2 DADOS PESSOAIS
- Layout **rótulo → campo** (coluna B = rótulo, coluna C = entrada), com uma coluna de **exemplo fictício** ao lado.
- Formatos personalizados: **CPF** (`000"."000"."000"-"00`), **CEP** (`00000"-"000`), **data** (`dd/mm/yyyy`).
- **Idade** calculada: `=SEERRO(DATEDIF(C7;HOJE();"y");"")`.
- Validações Sim/Não e listas de **UF** e **ocupação**.

### 2.3 a 2.7 Abas de lançamento (Rendimentos, Bens, Dívidas, Pagamentos, Dependentes)
- Cada uma usa uma **Tabela estruturada** (`TB_*`) com filtro automático e linhas zebradas.
- Cabeçalho fixo (freeze panes) e **linha de TOTAL** destacada com `SOMA` sobre o corpo da tabela.
- **Colunas calculadas** (live):
  - Bens → **Variação** = `=SE(E(E{r}="";F{r}="");"";SEERRO(F{r}-E{r};""))`
  - Dívidas → **Amortizado no ano** = `=SE(E(D{r}="";E{r}="");"";SEERRO(D{r}-E{r};""))`
- **Decisão:** ranges de corpo **fixos** (ex.: `F7:F31`) em vez de referências estruturadas com acento, para garantir que o dashboard funcione igual em qualquer versão/idioma do Excel.

### 2.8 DOCUMENTOS
- Estrutura pedida no desafio: `Documento | Categoria | Ano | Status | Observação | Link`.
- Coluna **Link** com hyperlinks (aponta para o arquivo do comprovante em nuvem/pasta).
- **Formatação condicional** por status: `Pendente` (vermelho), `Solicitado` (âmbar), `Recebido`/`Anexado` (verde).

### 2.9 RESUMO (Dashboard)
- **8 KPIs** em cartões: total de rendimentos, imposto retido, total de bens, total de dívidas, patrimônio líquido, despesas dedutíveis, dependentes e documentos pendentes.
- **Quebra por categoria** (rendimentos) e **por grupo** (bens) com `SOMASE`.
- **Controle de documentos**: cadastrados, concluídos, pendentes e **% de conclusão**, com **barra de dados**.

### 2.10 LISTAS
- Uma coluna por lista, com cabeçalho = **nome do intervalo nomeado**.
- 12 intervalos nomeados: `SIM_NAO`, `UF`, `TITDEP`, `CAT_REND`, `GRUPO_BENS`, `TIPO_DIVIDA`, `CAT_PAG`, `PARENTESCO`, `STATUS_DOC`, `CAT_DOC`, `OCUPACAO`, `BANCOS`.

---

## 3. Tabelas estruturadas

| Tabela | Aba | Papel |
|---|---|---|
| `TB_RENDIMENTOS` | RENDIMENTOS | Fontes de renda |
| `TB_BENS_E_DIREITOS` | BENS E DIREITOS | Patrimônio |
| `TB_DIVIDAS_E_ONUS` | DÍVIDAS E ÔNUS | Passivos |
| `TB_PAGAMENTOS` | PAGAMENTOS | Despesas dedutíveis |
| `TB_DEPENDENTES` | DEPENDENTES | Dependentes |
| `TB_DOCUMENTOS` | DOCUMENTOS | Comprovantes |

**Por quê tabelas?** Filtro/ordenação nativos, faixas zebradas automáticas e expansão fácil (ao digitar uma nova linha, o estilo é herdado).

---

## 4. Fórmulas utilizadas (e a razão de cada uma)

| Função | Onde | Para quê |
|---|---|---|
| `SOMA` | Totais das abas e KPIs do dashboard | Somar colunas de valores |
| `SOMASE` | Dashboard (por categoria / por grupo) | Somar condicionado à categoria/grupo |
| `CONT.SE` | Documentos pendentes/concluídos | Contar por status |
| `CONT.VALORES` | Dependentes e documentos cadastrados | Contar itens preenchidos |
| `SE` + `E` | Colunas calculadas (variação/amortização) | Ocultar resultado em linha vazia |
| `SEERRO` | Idade, variação, % de conclusão | Blindar contra erros (`#DIV/0!`, `#VALOR!`) |
| `DATEDIF` + `HOJE` | Idade em DADOS PESSOAIS | Idade automática a partir do nascimento |

**Princípio adotado:** fórmulas **simples e legíveis**. Nenhuma fórmula matricial ou construção desnecessariamente complexa — qualquer usuário consegue entender e manter.

### Exemplos reais do arquivo
```
Total de rendimentos (RESUMO!B7):
  =SOMA(RENDIMENTOS!F7:F31)+SOMA(RENDIMENTOS!G7:G31)

Rendimento por categoria (RESUMO!E17):
  =SOMASE(RENDIMENTOS!B7:B31;B17;RENDIMENTOS!F7:F31)
  +SOMASE(RENDIMENTOS!B7:B31;B17;RENDIMENTOS!G7:G31)

Documentos pendentes (RESUMO!H11):
  =CONT.SE(DOCUMENTOS!E7:E31;"Pendente")+CONT.SE(DOCUMENTOS!E7:E31;"Solicitado")

% de conclusão (RESUMO):
  =SEERRO((CONT.SE(...;"Anexado")+CONT.SE(...;"Recebido"))/CONT.VALORES(...);0)
```
> Observação: os nomes das funções aparecem em inglês (`SUM`, `SUMIF`, `COUNTIF`…) no motor de cálculo; o Excel exibe automaticamente na versão em português.

---

## 5. Validações de dados

| Lista (intervalo nomeado) | Usada em |
|---|---|
| `SIM_NAO` | Dados pessoais (booleanos), Pagamentos (Dedutível?) |
| `UF` | Dados pessoais |
| `OCUPACAO` | Dados pessoais |
| `CAT_REND` | Rendimentos |
| `GRUPO_BENS` | Bens e Direitos |
| `TIPO_DIVIDA` | Dívidas e Ônus |
| `CAT_PAG` | Pagamentos |
| `PARENTESCO` | Dependentes |
| `TITDEP` | Rendimentos, Pagamentos |
| `STATUS_DOC` | Rendimentos, Bens, Dívidas, Pagamentos, Documentos |
| `CAT_DOC` | Documentos |
| `BANCOS` | disponível para uso (lista de apoio, herdada do material da DIO) |

Todas as validações do tipo **lista** apontam para um **intervalo nomeado** na aba LISTAS, com **mensagem de erro** ("Selecione um item da lista suspensa") — o usuário não digita livremente o que poderia escolher.

---

## 6. Regras e formatação condicional

- **Status dos documentos** (e comprovantes das demais abas): cor por valor — vermelho (Pendente), âmbar (Solicitado), verde (Recebido/Anexado).
- **% de conclusão dos documentos**: **barra de dados** verde (0 → 100%).
- **Patrimônio líquido**: exibido em verde no cartão do dashboard.

---

## 7. Navegação

- **Menu principal** na aba INÍCIO (8 hyperlinks internos).
- **Botão "Início"** no canto superior direito de **todas** as abas de conteúdo.
- **Links rápidos** externos para sites oficiais.
- Painéis congelados (*freeze panes*) mantêm cabeçalhos visíveis na rolagem.

---

## 8. Indicadores (dashboard)

| Indicador | Fonte |
|---|---|
| Total de rendimentos | `SOMA` tributável + isento (RENDIMENTOS) |
| Imposto retido na fonte | `SOMA` (RENDIMENTOS) |
| Total de bens (31/12) | `SOMA` situação atual (BENS) |
| Total de dívidas (31/12) | `SOMA` saldo atual (DÍVIDAS) |
| Patrimônio líquido | Bens − Dívidas |
| Despesas dedutíveis | `SOMA` (PAGAMENTOS) |
| Dependentes | `CONT.VALORES` (DEPENDENTES) |
| Documentos pendentes | `CONT.SE` Pendente + Solicitado (DOCUMENTOS) |
| % de conclusão dos documentos | Concluídos ÷ cadastrados |

---

## 9. Decisões de projeto (resumo)

1. **Sem VBA/macros** → compatibilidade e segurança máximas (arquivo `.xlsx`, não `.xlsm`).
2. **Ranges fixos** no dashboard → previsibilidade em qualquer idioma/versão do Excel.
3. **Aba LISTAS única** → fonte única de verdade para todas as validações.
4. **`SEERRO` em toda fórmula sensível** → o arquivo nunca "quebra" com abas vazias.
5. **Dados de exemplo fictícios e sinalizados** → demonstram o funcionamento sem confundir com dados reais.
6. **Solução própria** inspirada na referência da DIO, porém reorganizada e ampliada (o projeto de referência tinha 4 abas; esta versão tem 9 + apoio, seguindo o escopo do desafio).

---

## 10. Limitações conhecidas

- As áreas das tabelas têm um **número fixo de linhas** de exemplo (20–25). Para mais lançamentos, basta **arrastar/estender a tabela** — as validações e o estilo são herdados, mas os ranges do dashboard podem precisar de ajuste se ultrapassarem o intervalo previsto.
- **Não substitui o programa oficial** da Receita nem faz cálculo de imposto devido/restituição — é uma ferramenta de **organização** dos dados.
- Não há importação automática de informes (ver *Melhorias futuras* no README).

---

## 11. Como o arquivo foi produzido

A planilha foi construída de forma **reprodutível** (via biblioteca de manipulação de `.xlsx`), o que garante consistência de estilos, fórmulas e validações. Em seguida foi feita uma **auditoria automatizada**: recálculo de todas as fórmulas e conferência dos totais do dashboard contra os dados de exemplo (todos bateram). As imagens em `/images` são capturas reais dessa versão final.
