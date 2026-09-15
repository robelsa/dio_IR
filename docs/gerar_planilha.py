# -*- coding: utf-8 -*-
"""
Agregador de Imposto de Renda - DIO
Gera a ferramenta Excel completa (9 abas + Listas) com validacoes,
formulas, navegacao, dashboard e design profissional.
"""
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.worksheet.hyperlink import Hyperlink
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName

# ------------------------------------------------------------------ PALETA
NAVY   = "14304A"
NAVY2  = "1E4A6D"
TEAL   = "1E88A8"
TEALDK = "13657E"
LIGHT  = "F4F7FA"
BAND   = "E8EEF4"
CARD   = "FFFFFF"
LINE   = "C7D3E0"
TXT    = "1B2A3A"
MUTED  = "6B7A8D"
GREEN  = "2E7D46"
GREENL = "E4F1E8"
RED    = "C0392B"
REDL   = "FBE7E5"
AMBER  = "B77800"
AMBERL = "FBF0D9"
WHITE  = "FFFFFF"
INPUT  = "FFFFFF"

F = "Calibri"
def font(sz=11, b=False, color=TXT, italic=False):
    return Font(name=F, size=sz, bold=b, color=color, italic=italic)
def fill(c):
    return PatternFill("solid", fgColor=c)
def side(c=LINE, style="thin"):
    return Side(style=style, color=c)
def box(c=LINE, style="thin"):
    s = side(c, style)
    return Border(left=s, right=s, top=s, bottom=s)
def align(h="left", v="center", wrap=False):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)

thin = box()
CUR = 'R$ #,##0.00'
CUR_RED = 'R$ #,##0.00;[Red]-R$ #,##0.00'
DATEF = 'dd/mm/yyyy'
CPF = '000"."000"."000"-"00'
CNPJ = '00"."000"."000"/"0000"-"00'
CEP = '00000"-"000'
PCT = '0.0%'

wb = Workbook()

# ------------------------------------------------------------------ LISTAS
LISTS = {
    "SIM_NAO": ["Sim", "Nao"],
    "UF": ["AC","AL","AP","AM","BA","CE","DF","ES","GO","MA","MT","MS","MG","PA",
           "PB","PR","PE","PI","RJ","RN","RS","RO","RR","SC","SP","SE","TO"],
    "TITDEP": ["Titular", "Dependente"],
    "CAT_REND": [
        "Salario / CLT","13o salario","Aposentadoria / Pensao (INSS)",
        "Pensao alimenticia recebida","Trabalho autonomo / PJ","Alugueis",
        "Rendimentos de aplicacoes financeiras","Dividendos / Lucros",
        "Rendimentos recebidos do exterior","Outros rendimentos"],
    "GRUPO_BENS": [
        "Imovel","Veiculo","Conta corrente / poupanca",
        "Aplicacao financeira","Participacao societaria","Criptoativo",
        "Bens moveis / Outros"],
    "TIPO_DIVIDA": [
        "Financiamento imobiliario","Financiamento de veiculo",
        "Emprestimo pessoal","Cartao de credito","Consorcio","Outras dividas"],
    "CAT_PAG": [
        "Despesas medicas","Plano de saude","Educacao",
        "Previdencia privada (PGBL)","Pensao alimenticia judicial",
        "Contribuicao INSS","Doacoes incentivadas","Outras despesas"],
    "PARENTESCO": [
        "Conjuge / Companheiro(a)","Filho(a)","Enteado(a)",
        "Pais / Avos / Bisavos","Neto(a) / Bisneto(a)","Irmao(a)",
        "Menor sob guarda judicial","Outros"],
    "STATUS_DOC": ["Pendente","Solicitado","Recebido","Anexado","Nao se aplica"],
    "CAT_DOC": ["Dados Pessoais","Rendimentos","Bens e Direitos","Dividas",
                "Pagamentos / Despesas","Dependentes","Outros"],
    "OCUPACAO": ["Empregado(a) CLT","Servidor(a) publico(a)","Autonomo(a) / MEI",
                 "Empresario(a)","Aposentado(a) / Pensionista","Estudante","Outros"],
    "BANCOS": [
        "1 - Banco do Brasil","33 - Banco Santander","77 - Banco Inter",
        "104 - Caixa Economica Federal","208 - Banco BTG Pactual",
        "212 - Banco Original","237 - Banco Bradesco","260 - Nubank",
        "290 - PagBank","336 - C6 Bank","341 - Itau Unibanco","380 - PicPay",
        "422 - Banco Safra","748 - Sicredi","102 - XP Investimentos","Outros"],
}
lst = wb.active
lst.title = "LISTAS"
lst.sheet_properties.tabColor = MUTED
lst.sheet_view.showGridLines = False
lst["A1"] = "TABELAS DE APOIO / LISTAS DE VALIDACAO"
lst["A1"].font = font(12, True, NAVY)
lst["A2"] = ("Esta aba alimenta os menus suspensos (validacao de dados) das demais abas. "
             "Nao renomeie os cabecalhos. Voce pode acrescentar itens ao final de cada lista.")
lst["A2"].font = font(9, False, MUTED, italic=True)

ranges = {}
col = 1
for name, values in LISTS.items():
    L = get_column_letter(col)
    h = lst.cell(row=4, column=col, value=name)
    h.font = font(10, True, WHITE); h.fill = fill(NAVY2); h.alignment = align("center")
    h.border = thin
    for i, v in enumerate(values):
        c = lst.cell(row=5+i, column=col, value=v)
        c.font = font(10); c.border = box(BAND)
    lst.column_dimensions[L].width = 30
    rng = f"LISTAS!${L}$5:${L}${5+len(values)-1}"
    ranges[name] = rng
    # named range
    dn = DefinedName(name, attr_text=rng)
    wb.defined_names.add(dn)
    col += 1
lst.column_dimensions["A"].width = 30
lst.freeze_panes = "A5"

# ------------------------------------------------------------------ HELPERS
def dv_list(ws, name, cell_range, prompt=None):
    dv = DataValidation(type="list", formula1=f"={name}", allow_blank=True,
                        showErrorMessage=True, showInputMessage=bool(prompt))
    dv.errorTitle = "Valor invalido"
    dv.error = "Selecione um item da lista suspensa."
    if prompt:
        dv.promptTitle = "Selecione"
        dv.prompt = prompt
    ws.add_data_validation(dv)
    dv.add(cell_range)
    return dv

def dv_type(ws, cell_range, dtype, op=None, f1=None, f2=None, title="", msg=""):
    dv = DataValidation(type=dtype, operator=op, formula1=f1, formula2=f2,
                        allow_blank=True, showErrorMessage=True)
    dv.errorTitle = title or "Valor invalido"
    dv.error = msg or "Verifique o valor digitado."
    ws.add_data_validation(dv)
    dv.add(cell_range)
    return dv

def internal_link(ws, coord, text, target_sheet, target_cell="A1",
                  fsz=11, fb=True, fc=WHITE, bg=TEAL, h="center"):
    c = ws[coord]
    c.value = text
    c.hyperlink = Hyperlink(ref=coord, location=f"'{target_sheet}'!{target_cell}", display=text)
    c.font = font(fsz, fb, fc)
    c.fill = fill(bg)
    c.alignment = align(h, "center")
    c.border = box(bg)
    return c

def ext_link(ws, coord, text, url, fc=TEAL):
    c = ws[coord]
    c.value = text
    c.hyperlink = url
    c.font = Font(name=F, size=10, bold=True, color=fc, underline="single")
    c.alignment = align("left", "center")
    return c

def header_band(ws, section_no, section_title, subtitle, ncols=8, back=True):
    """Faixa de titulo padrao no topo de cada aba de conteudo."""
    last = get_column_letter(ncols)
    band_last_i = ncols - 1 if back else ncols
    band_last = get_column_letter(band_last_i)
    # linha 1: barra navy com nome do projeto
    ws.merge_cells(f"B1:{band_last}1")
    t = ws["B1"]
    t.value = "AGREGADOR DE IMPOSTO DE RENDA"
    t.font = font(10, True, "AECBE0"); t.fill = fill(NAVY); t.alignment = align("left")
    for col in range(2, ncols+1):
        ws.cell(row=1, column=col).fill = fill(NAVY)
    ws.row_dimensions[1].height = 20
    # linha 2-3: titulo da secao
    ws.merge_cells(f"B2:{band_last}3")
    s = ws["B2"]
    s.value = f"  {section_no}  {section_title}"
    s.font = font(20, True, NAVY); s.fill = fill(BAND); s.alignment = align("left", "center")
    for r in (2, 3):
        for col in range(2, ncols+1):
            ws.cell(row=r, column=col).fill = fill(BAND)
    ws.row_dimensions[2].height = 20; ws.row_dimensions[3].height = 22
    # linha 4: subtitulo/instrucao
    ws.merge_cells(f"B4:{last}4")
    sub = ws["B4"]; sub.value = "   " + subtitle
    sub.font = font(10, False, MUTED, italic=True); sub.alignment = align("left", "center")
    ws.row_dimensions[4].height = 18
    # botao voltar (coluna reservada a direita, linhas 1-3)
    if back:
        ws.merge_cells(f"{last}1:{last}3")
        internal_link(ws, f"{last}1", "Inicio", "INICIO", "A1",
                      fsz=11, fb=True, fc=WHITE, bg=TEAL, h="center")
        ws[f"{last}1"].alignment = align("center", "center")

def col_widths(ws, widths, start="B"):
    c0 = openpyxl.utils.column_index_from_string(start)
    for i, w in enumerate(widths):
        ws.column_dimensions[get_column_letter(c0+i)].width = w

def table_header(ws, row, start_col, headers):
    for i, htext in enumerate(headers):
        c = ws.cell(row=row, column=start_col+i, value=htext)
        c.font = font(10.5, True, WHITE); c.fill = fill(NAVY2)
        c.alignment = align("center", "center", wrap=True); c.border = box(NAVY)
    ws.row_dimensions[row].height = 30

def style_body(ws, r0, r1, c0, c1, banded=True):
    for r in range(r0, r1+1):
        for c in range(c0, c1+1):
            cell = ws.cell(row=r, column=c)
            cell.border = box(LINE)
            cell.font = font(10.5)
            if banded and (r - r0) % 2 == 1:
                cell.fill = fill(LIGHT)
            else:
                cell.fill = fill(WHITE)

def total_row(ws, row, label_col, first_col, last_col, sum_cols, r0, r1, label="TOTAL"):
    lc = ws.cell(row=row, column=label_col, value=label)
    lc.font = font(11, True, WHITE); lc.fill = fill(TEALDK); lc.alignment = align("right")
    for c in range(first_col, last_col+1):
        cell = ws.cell(row=row, column=c)
        cell.fill = fill(TEALDK); cell.border = box(TEALDK)
        if c in sum_cols:
            L = get_column_letter(c)
            cell.value = f"=SUM({L}{r0}:{L}{r1})"
            cell.number_format = CUR
            cell.font = font(11, True, WHITE); cell.alignment = align("center")
    ws.row_dimensions[row].height = 22

def note(ws, coord, text):
    c = ws[coord]; c.value = text
    c.font = font(9, False, MUTED, italic=True)

# ================================================================== INICIO
ws = wb.create_sheet("INICIO")
ws.sheet_properties.tabColor = NAVY
ws.sheet_view.showGridLines = False
col_widths(ws, [2,26,26,26,4], start="A")  # A..E
for i in range(1, 40):
    ws.row_dimensions[i].height = 18

# capa
ws.merge_cells("B2:D2")
ws["B2"] = "AGREGADOR DE"
ws["B2"].font = font(16, True, TEAL); ws["B2"].alignment = align("center")
ws.merge_cells("B3:D4")
ws["B3"] = "IMPOSTO DE RENDA"
ws["B3"].font = font(34, True, NAVY); ws["B3"].alignment = align("center", "center")
ws.merge_cells("B5:D5")
ws["B5"] = "Organizador de dados para a Declaracao de IRPF  -  Pessoa Fisica"
ws["B5"].font = font(11, False, MUTED); ws["B5"].alignment = align("center")
ws.merge_cells("B6:D6")
ws["B6"] = "Desafio DIO  |  Microsoft Excel"
ws["B6"].font = font(9, True, TEALDK); ws["B6"].alignment = align("center")

# descricao / instrucoes
ws.merge_cells("B8:D8")
ws["B8"] = "COMO USAR"
ws["B8"].font = font(12, True, WHITE); ws["B8"].fill = fill(NAVY2); ws["B8"].alignment = align("center")
instr = [
    "1. Comece pela aba DADOS PESSOAIS e preencha a identificacao do contribuinte.",
    "2. Registre suas informacoes nas abas de RENDIMENTOS, BENS, DIVIDAS e PAGAMENTOS.",
    "3. Cadastre seus DEPENDENTES e controle os COMPROVANTES na aba DOCUMENTOS.",
    "4. Acompanhe os totais consolidados na aba RESUMO (dashboard).",
    "5. Prefira sempre os menus suspensos; campos com * sao obrigatorios.",
]
r = 9
for line in instr:
    ws.merge_cells(f"B{r}:D{r}")
    ws[f"B{r}"] = line
    ws[f"B{r}"].font = font(10.5); ws[f"B{r}"].alignment = align("left", "center")
    ws.row_dimensions[r].height = 20
    r += 1

# menu de navegacao (cards)
ws.merge_cells("B15:D15")
ws["B15"] = "MENU DE NAVEGACAO"
ws["B15"].font = font(12, True, WHITE); ws["B15"].fill = fill(NAVY2); ws["B15"].alignment = align("center")

menu = [
    ("1  Dados Pessoais", "DADOS PESSOAIS"),
    ("2  Rendimentos", "RENDIMENTOS"),
    ("3  Bens e Direitos", "BENS E DIREITOS"),
    ("4  Dividas e Onus", "DIVIDAS E ONUS"),
    ("5  Pagamentos / Despesas", "PAGAMENTOS"),
    ("6  Dependentes", "DEPENDENTES"),
    ("7  Documentos", "DOCUMENTOS"),
    ("8  Resumo (Dashboard)", "RESUMO"),
]
positions = [("B",16),("C",16),("D",16),
             ("B",18),("C",18),("D",18),
             ("B",20),("C",20)]
# place 3 per row
cells = ["B16","C16","D16","B18","C18","D18","B20","C20"]
for (text, target), coord in zip(menu, cells):
    internal_link(ws, coord, text, target, "A1", fsz=11, fb=True, fc=WHITE, bg=TEAL, h="center")
    ws[coord].alignment = align("center", "center", wrap=True)
for rr in (16,18,20):
    ws.row_dimensions[rr].height = 34
    ws.row_dimensions[rr+1].height = 6

# links rapidos externos
ws.merge_cells("B23:D23")
ws["B23"] = "LINKS RAPIDOS  (sites oficiais)"
ws["B23"].font = font(12, True, WHITE); ws["B23"].fill = fill(NAVY2); ws["B23"].alignment = align("center")
links = [
    ("Portal e-CAC - Receita Federal", "https://cav.receita.fazenda.gov.br/"),
    ("Programa IRPF / Meu Imposto de Renda", "https://www.gov.br/receitafederal/pt-br/assuntos/meu-imposto-de-renda"),
    ("Consulta CPF / Situacao cadastral", "https://www.gov.br/receitafederal/pt-br/assuntos/meu-cpf"),
    ("Perguntas e respostas IRPF", "https://www.gov.br/receitafederal/pt-br/centrais-de-conteudo/publicacoes/perguntas-e-respostas/dirpf"),
]
r = 24
for text, url in links:
    ws.merge_cells(f"B{r}:D{r}")
    ext_link(ws, f"B{r}", "   " + text, url)
    ws.row_dimensions[r].height = 18
    r += 1

# rodape
ws.merge_cells("B29:D29")
ws["B29"] = "Versao 1.0  -  Os dados de exemplo sao ficticios e servem apenas para demonstracao."
ws["B29"].font = font(9, False, MUTED, italic=True); ws["B29"].alignment = align("center")
ws.merge_cells("B30:D30")
ws["B30"] = "Autor: Bel  |  Projeto do desafio DIO"
ws["B30"].font = font(9, True, MUTED); ws["B30"].alignment = align("center")

# moldura branca no bloco
for rr in range(2, 31):
    ws.cell(row=rr, column=1).fill = fill(WHITE)
    ws.cell(row=rr, column=5).fill = fill(WHITE)

print("INICIO ok")

# ================================================================== DADOS PESSOAIS
ws = wb.create_sheet("DADOS PESSOAIS")
ws.sheet_properties.tabColor = TEAL
ws.sheet_view.showGridLines = False
col_widths(ws, [30, 34, 6, 30, 20], start="B")  # B..F
header_band(ws, "1.", "DADOS PESSOAIS", "Identificacao do contribuinte (titular). Campos com * sao obrigatorios.", ncols=6)

fields = [
    ("Nome completo *", "text", None),
    ("CPF *", "cpf", None),
    ("Data de nascimento *", "date", None),
    ("Idade", "formula", "=IFERROR(DATEDIF(C7,TODAY(),\"y\"),\"\")"),
    ("Titulo de eleitor", "text", None),
    ("Nome da mae", "text", None),
    ("Ocupacao principal", "list:OCUPACAO", None),
    ("E-mail", "text", None),
    ("Telefone / Celular", "text", None),
    ("CEP", "cep", None),
    ("Logradouro (rua, av.)", "text", None),
    ("Numero", "text", None),
    ("Complemento", "text", None),
    ("Bairro", "text", None),
    ("Cidade", "text", None),
    ("UF", "list:UF", None),
    ("Primeira declaracao?", "list:SIM_NAO", None),
    ("Houve alteracoes desde a ultima entrega?", "list:SIM_NAO", None),
    ("Possui conjuge / companheiro(a)?", "list:SIM_NAO", None),
    ("Residente no exterior?", "list:SIM_NAO", None),
]
r0 = 6
for i, (label, kind, extra) in enumerate(fields):
    r = r0 + i
    lc = ws.cell(row=r, column=2, value=label)  # B
    lc.font = font(10.5, True, NAVY); lc.alignment = align("left", "center")
    lc.fill = fill(BAND); lc.border = box(LINE)
    vc = ws.cell(row=r, column=3)  # C
    vc.fill = fill(INPUT); vc.border = box(LINE); vc.alignment = align("left", "center")
    vc.font = font(10.5)
    ws.row_dimensions[r].height = 21
    ref = f"C{r}"
    if kind == "cpf":
        vc.number_format = CPF
    elif kind == "cep":
        vc.number_format = CEP
    elif kind == "date":
        vc.number_format = DATEF
    elif kind == "formula":
        vc.value = extra; vc.fill = fill(GREENL); vc.font = font(10.5, True, GREEN)
    elif kind.startswith("list:"):
        name = kind.split(":")[1]
        dv_list(ws, name, ref, prompt="Selecione na lista")

# exemplo (ficticio) ao lado
ws.merge_cells("E6:F6")
ws["E6"] = "EXEMPLO (ficticio)"
ws["E6"].font = font(10, True, WHITE); ws["E6"].fill = fill(AMBER); ws["E6"].alignment = align("center")
examples = [
    "MARIA FICTICIA DA SILVA","111.444.777-35","10/05/1990","(auto)","1234567890 SP",
    "JOANA FICTICIA","Autonomo(a) / MEI","maria.exemplo@email.com","(11) 99999-0000",
    "01310-100","Av. das Demonstracoes","1000","Sala 5","Centro","Sao Paulo","SP",
    "Nao","Sim","Sim","Nao",
]
for i, ex in enumerate(examples):
    r = r0 + i
    ws.merge_cells(f"E{r}:F{r}")
    c = ws.cell(row=r, column=5, value=ex)
    c.font = font(9.5, False, MUTED, italic=True); c.alignment = align("left", "center")
note(ws, "B27", "* Preencha antes de iniciar a declaracao no programa da Receita. O CPF e a data de nascimento sao obrigatorios para o acesso ao e-CAC.")
ws.freeze_panes = "B5"
print("DADOS ok")

# ================================================================== RENDIMENTOS
def entry_sheet(sheet, tabcolor, no, title, subtitle, headers, widths,
                money_cols, example_rows, list_cols, name_ranges,
                body_rows=25, computed=None, totals=True, status_col=None):
    ws = wb.create_sheet(sheet)
    ws.sheet_properties.tabColor = tabcolor
    ws.sheet_view.showGridLines = False
    ncols = 1 + len(headers)  # começa na coluna B
    header_band(ws, no, title, subtitle, ncols=ncols+0)
    col_widths(ws, widths, start="B")
    hrow = 6
    start_col = 2  # B
    table_header(ws, hrow, start_col, headers)
    r0 = hrow + 1
    r1 = r0 + body_rows - 1
    style_body(ws, r0, r1, start_col, start_col+len(headers)-1)
    # money formats
    for mc in money_cols:
        for r in range(r0, r1+1):
            ws.cell(row=r, column=start_col+mc).number_format = CUR
            ws.cell(row=r, column=start_col+mc).alignment = align("right")
    # dropdowns
    for ci, name in list_cols.items():
        L = get_column_letter(start_col+ci)
        dv_list(ws, name, f"{L}{r0}:{L}{r1}", prompt="Selecione")
    # example rows
    for ridx, row in enumerate(example_rows):
        r = r0 + ridx
        for ci, val in enumerate(row):
            c = ws.cell(row=r, column=start_col+ci)
            if val is not None:
                c.value = val
            c.font = font(10.5, False, TXT)
    # computed columns (dict col_index -> formula template using {r}) as live column
    if computed:
        for ci, tmpl in computed.items():
            for r in range(r0, r1+1):
                cc = ws.cell(row=r, column=start_col+ci)
                cc.value = tmpl.format(r=r)
                cc.number_format = CUR
                cc.alignment = align("right")
    # tag exemplo
    tag = ws.cell(row=r0, column=start_col+len(headers)+1)
    # (skip visual tag to keep clean)
    # total row
    total_at = r1 + 1
    if totals:
        sumcols = [start_col+mc for mc in money_cols]
        total_row(ws, total_at, start_col, start_col, start_col+len(headers)-1,
                  sumcols, r0, r1)
    # excel table
    last_col_letter = get_column_letter(start_col+len(headers)-1)
    ref = f"B{hrow}:{last_col_letter}{r1}"
    tab = Table(displayName="TB_"+sheet.replace(" ","_").replace("/","").replace("-",""), ref=ref)
    tab.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True,
                                        showColumnStripes=False, showFirstColumn=False,
                                        showLastColumn=False)
    ws.add_table(tab)
    # status conditional format
    if status_col is not None:
        L = get_column_letter(start_col+status_col)
        rng = f"{L}{r0}:{L}{r1}"
        ws.conditional_formatting.add(rng, CellIsRule(operator="equal",
            formula=['"Pendente"'], fill=fill(REDL), font=Font(name=F, color=RED, bold=True)))
        ws.conditional_formatting.add(rng, CellIsRule(operator="equal",
            formula=['"Solicitado"'], fill=fill(AMBERL), font=Font(name=F, color=AMBER, bold=True)))
        for ok in ("Recebido","Anexado"):
            ws.conditional_formatting.add(rng, CellIsRule(operator="equal",
                formula=[f'"{ok}"'], fill=fill(GREENL), font=Font(name=F, color=GREEN, bold=True)))
    ws.freeze_panes = f"B{hrow+1}"
    note(ws, f"B{total_at+2}", "Dados de exemplo sao ficticios. Preencha novas linhas dentro da tabela; os totais se atualizam automaticamente.")
    return ws, r0, r1, total_at

# --- RENDIMENTOS
entry_sheet(
    "RENDIMENTOS", TEAL, "2.", "RENDIMENTOS",
    "Registre todas as fontes de renda do ano-base, separando titular e dependentes.",
    ["Categoria","Fonte pagadora","CNPJ/CPF da fonte","Titular/Dep.",
     "Valor tributavel (R$)","Rend. isento (R$)","Imposto retido (R$)","Comprovante"],
    [26,26,20,14,18,18,18,15],
    money_cols=[4,5,6],
    example_rows=[
        ["Salario / CLT","Empresa Exemplo LTDA","12.345.678/0001-99","Titular",54000,0,4200,"Anexado"],
        ["Aposentadoria / Pensao (INSS)","INSS","29.979.036/0001-40","Titular",0,28000,0,"Recebido"],
        ["Alugueis","Imovel Rua das Demonstracoes (aluguel)","",  "Titular",14400,0,0,"Pendente"],
    ],
    list_cols={0:"CAT_REND", 3:"TITDEP", 7:"STATUS_DOC"},
    name_ranges=ranges, body_rows=25, status_col=7,
)
# fix the 3rd example row (shifted) -> rewrite cleanly
print("RENDIMENTOS ok")

# --- BENS E DIREITOS
entry_sheet(
    "BENS E DIREITOS", "2E7D8A", "3.", "BENS E DIREITOS",
    "Relacione os bens e direitos em 31/12. A variacao (atual - anterior) e calculada automaticamente.",
    ["Grupo / Tipo","Descricao do bem","Localizacao / Instituicao",
     "Situacao 31/12 ANTERIOR (R$)","Situacao 31/12 ATUAL (R$)","Variacao (R$)","Comprovante"],
    [24,30,24,20,20,16,15],
    money_cols=[3,4,5],
    example_rows=[
        ["Imovel","Apartamento 2 quartos (exemplo)","Sao Paulo/SP",180000,180000,None,"Anexado"],
        ["Veiculo","Automovel popular 2019","Sao Paulo/SP",45000,40000,None,"Recebido"],
        ["Conta corrente / poupanca","Conta corrente","260 - Nubank",8000,12500,None,"Pendente"],
    ],
    list_cols={0:"GRUPO_BENS", 6:"STATUS_DOC"},
    name_ranges=ranges, body_rows=25,
    computed={5:"=IF(AND(E{r}=\"\",F{r}=\"\"),\"\",IFERROR(F{r}-E{r},\"\"))"},
    status_col=6,
)
print("BENS ok")

# --- DIVIDAS
entry_sheet(
    "DIVIDAS E ONUS", "B77800", "4.", "DIVIDAS E ONUS",
    "Registre financiamentos, emprestimos e demais dividas com saldo em 31/12.",
    ["Tipo","Credor / Instituicao","Descricao",
     "Saldo 31/12 ANTERIOR (R$)","Saldo 31/12 ATUAL (R$)","Amortizado no ano (R$)","Comprovante"],
    [24,26,28,20,20,18,15],
    money_cols=[3,4,5],
    example_rows=[
        ["Financiamento imobiliario","104 - Caixa Economica Federal","Financiamento do apartamento",120000,110000,None,"Anexado"],
        ["Financiamento de veiculo","341 - Itau Unibanco","Financiamento do automovel",30000,20000,None,"Pendente"],
    ],
    list_cols={0:"TIPO_DIVIDA", 6:"STATUS_DOC"},
    name_ranges=ranges, body_rows=20,
    computed={5:"=IF(AND(D{r}=\"\",E{r}=\"\"),\"\",IFERROR(D{r}-E{r},\"\"))"},
    status_col=6,
)
print("DIVIDAS ok")

# --- PAGAMENTOS
entry_sheet(
    "PAGAMENTOS", "7A5FA6", "5.", "PAGAMENTOS / DESPESAS DEDUTIVEIS",
    "Registre despesas dedutiveis (medicas, educacao, previdencia, pensao) com comprovante.",
    ["Categoria","Beneficiario / Prestador","CNPJ/CPF","Titular/Dep.",
     "Valor pago (R$)","Dedutivel?","Comprovante"],
    [26,28,20,14,18,14,15],
    money_cols=[4],
    example_rows=[
        ["Despesas medicas","Clinica Exemplo","01.234.567/0001-00","Titular",3200,"Sim","Anexado"],
        ["Educacao","Escola Modelo","09.876.543/0001-11","Dependente",9600,"Sim","Recebido"],
        ["Plano de saude","Operadora Demonstracao","11.222.333/0001-44","Titular",7200,"Sim","Pendente"],
    ],
    list_cols={0:"CAT_PAG", 3:"TITDEP", 5:"SIM_NAO", 6:"STATUS_DOC"},
    name_ranges=ranges, body_rows=25, status_col=6,
)
print("PAGAMENTOS ok")

# --- DEPENDENTES
entry_sheet(
    "DEPENDENTES", "C0553B", "6.", "DEPENDENTES",
    "Cadastre os dependentes declarados. CPF e obrigatorio para qualquer idade.",
    ["Nome do dependente","CPF","Data de nascimento","Grau de parentesco","Observacoes"],
    [30,18,18,26,30],
    money_cols=[],
    example_rows=[
        ["Pedro Ficticio da Silva","222.333.444-05","15/03/2015","Filho(a)","Estudante"],
        ["Ana Ficticia da Silva","333.444.555-06","20/07/2018","Filho(a)",""],
    ],
    list_cols={3:"PARENTESCO"},
    name_ranges=ranges, body_rows=15, totals=False,
)
# format CPF & date columns on DEPENDENTES
wsd = wb["DEPENDENTES"]
for r in range(7, 7+15):
    wsd.cell(row=r, column=3).number_format = CPF   # C
    wsd.cell(row=r, column=4).number_format = DATEF # D
print("DEPENDENTES ok")

# --- DOCUMENTOS
entry_sheet(
    "DOCUMENTOS", "166B85", "7.", "DOCUMENTOS / COMPROVANTES",
    "Controle central dos comprovantes. Use a coluna Link para apontar o arquivo (nuvem/pasta).",
    ["Documento","Categoria","Ano","Status","Observacao","Link do arquivo"],
    [34,24,10,15,28,26],
    money_cols=[],
    example_rows=[
        ["Informe de rendimentos - Empresa Exemplo","Rendimentos",2024,"Anexado","Fonte principal","https://exemplo.com/informe.pdf"],
        ["Informe INSS","Rendimentos",2024,"Recebido","Aposentadoria","https://exemplo.com/inss.pdf"],
        ["Recibos medicos","Pagamentos / Despesas",2024,"Pendente","Reunir recibos do ano",""],
        ["Comprovante IPTU / imovel","Bens e Direitos",2024,"Solicitado","Prefeitura",""],
    ],
    list_cols={1:"CAT_DOC", 3:"STATUS_DOC"},
    name_ranges=ranges, body_rows=25, totals=False, status_col=3,
)
# make link column real hyperlinks for the example rows
wsdoc = wb["DOCUMENTOS"]
for r in range(7, 7+25):
    v = wsdoc.cell(row=r, column=7).value  # G = link
    if isinstance(v, str) and v.startswith("http"):
        cell = wsdoc.cell(row=r, column=7)
        cell.hyperlink = v
        cell.font = Font(name=F, size=10, color=TEAL, underline="single")
    wsdoc.cell(row=r, column=4).number_format = "0"  # Ano as plain
print("DOCUMENTOS ok")

print("entry sheets done")

# ================================================================== RESUMO
ws = wb.create_sheet("RESUMO")
ws.sheet_properties.tabColor = NAVY
ws.sheet_view.showGridLines = False
col_widths(ws, [18,18,18,18,18,18,18,18], start="B")  # B..I
header_band(ws, "8.", "RESUMO (DASHBOARD)",
            "Visao consolidada das informacoes cadastradas. Os valores se atualizam automaticamente.", ncols=9)

R = {  # ranges auxiliares
 "rend_trib":"RENDIMENTOS!F7:F31","rend_isen":"RENDIMENTOS!G7:G31",
 "rend_imp":"RENDIMENTOS!H7:H31","rend_cat":"RENDIMENTOS!B7:B31",
 "bens_val":"'BENS E DIREITOS'!F7:F31","bens_grp":"'BENS E DIREITOS'!B7:B31",
 "div_val":"'DIVIDAS E ONUS'!F7:F26",
 "pag_val":"PAGAMENTOS!F7:F31",
 "dep_nome":"DEPENDENTES!B7:B21",
 "doc_nome":"DOCUMENTOS!B7:B31","doc_stat":"DOCUMENTOS!E7:E31",
}

def card(r, c, span, label, formula, numfmt=CUR, accent=TEAL, big=18, valcolor=NAVY):
    L0=get_column_letter(c); L1=get_column_letter(c+span-1)
    ws.merge_cells(f"{L0}{r}:{L1}{r}")
    lc=ws[f"{L0}{r}"]; lc.value=label; lc.font=font(9,True,WHITE)
    lc.fill=fill(accent); lc.alignment=align("center","center",True)
    ws.merge_cells(f"{L0}{r+1}:{L1}{r+2}")
    vc=ws[f"{L0}{r+1}"]; vc.value=formula; vc.number_format=numfmt
    vc.font=font(big,True,valcolor); vc.fill=fill(WHITE); vc.alignment=align("center","center")
    for rr in range(r,r+3):
        for cc in range(c,c+span):
            ws.cell(row=rr,column=cc).border=box(LINE)
    ws.row_dimensions[r].height=24; ws.row_dimensions[r+1].height=16; ws.row_dimensions[r+2].height=16

# Linha 1 de cards (row 6)
card(6,2,2,"TOTAL DE RENDIMENTOS",   f"=SUM({R['rend_trib']})+SUM({R['rend_isen']})", accent=TEAL)
card(6,4,2,"IMPOSTO RETIDO NA FONTE",f"=SUM({R['rend_imp']})", accent=TEAL)
card(6,6,2,"TOTAL DE BENS (31/12)",  f"=SUM({R['bens_val']})", accent=TEALDK)
card(6,8,2,"TOTAL DE DIVIDAS (31/12)",f"=SUM({R['div_val']})", accent=AMBER)
# Linha 2 de cards (row 10)
card(10,2,2,"PATRIMONIO LIQUIDO",    f"=SUM({R['bens_val']})-SUM({R['div_val']})", accent=NAVY2, valcolor=GREEN)
card(10,4,2,"DESPESAS DEDUTIVEIS",   f"=SUM({R['pag_val']})", accent=NAVY2)
card(10,6,2,"DEPENDENTES",           f"=COUNTA({R['dep_nome']})", numfmt="0", accent=NAVY2)
card(10,8,2,"DOCUMENTOS PENDENTES",  f"=COUNTIF({R['doc_stat']},\"Pendente\")+COUNTIF({R['doc_stat']},\"Solicitado\")",
     numfmt="0", accent=RED, valcolor=RED)

# ---- Rendimentos por categoria
sr = 15
ws.merge_cells(f"B{sr}:E{sr}")
ws[f"B{sr}"]="RENDIMENTOS POR CATEGORIA"
ws[f"B{sr}"].font=font(11,True,WHITE); ws[f"B{sr}"].fill=fill(NAVY2); ws[f"B{sr}"].alignment=align("center")
ws.merge_cells(f"B{sr+1}:D{sr+1}"); ws[f"B{sr+1}"]="Categoria"
ws[f"B{sr+1}"].font=font(10,True,NAVY); ws[f"B{sr+1}"].fill=fill(BAND); ws[f"B{sr+1}"].border=thin
ws.merge_cells(f"B{sr+1}:D{sr+1}")
ws[f"E{sr+1}"]="Valor (R$)"; ws[f"E{sr+1}"].font=font(10,True,NAVY); ws[f"E{sr+1}"].fill=fill(BAND)
ws[f"E{sr+1}"].alignment=align("center"); ws[f"E{sr+1}"].border=thin
for i,cat in enumerate(LISTS["CAT_REND"]):
    r=sr+2+i
    ws.merge_cells(f"B{r}:D{r}")
    lc=ws[f"B{r}"]; lc.value=cat; lc.font=font(10); lc.border=box(LINE); lc.alignment=align("left")
    ws[f"C{r}"].border=box(LINE); ws[f"D{r}"].border=box(LINE)
    vc=ws[f"E{r}"]
    vc.value=f"=SUMIF({R['rend_cat']},B{r},{R['rend_trib']})+SUMIF({R['rend_cat']},B{r},{R['rend_isen']})"
    vc.number_format=CUR; vc.font=font(10); vc.border=box(LINE); vc.alignment=align("right")
    if (i%2)==1:
        for L in ("B","C","D","E"): ws[f"{L}{r}"].fill=fill(LIGHT)

# ---- Bens por grupo
ws.merge_cells(f"G{sr}:I{sr}")
ws[f"G{sr}"]="BENS POR GRUPO"
ws[f"G{sr}"].font=font(11,True,WHITE); ws[f"G{sr}"].fill=fill(NAVY2); ws[f"G{sr}"].alignment=align("center")
ws.merge_cells(f"G{sr+1}:H{sr+1}"); ws[f"G{sr+1}"]="Grupo"
ws[f"G{sr+1}"].font=font(10,True,NAVY); ws[f"G{sr+1}"].fill=fill(BAND); ws[f"G{sr+1}"].border=thin
ws[f"H{sr+1}"].fill=fill(BAND); ws[f"H{sr+1}"].border=thin
ws[f"I{sr+1}"]="Valor (R$)"; ws[f"I{sr+1}"].font=font(10,True,NAVY); ws[f"I{sr+1}"].fill=fill(BAND)
ws[f"I{sr+1}"].alignment=align("center"); ws[f"I{sr+1}"].border=thin
for i,grp in enumerate(LISTS["GRUPO_BENS"]):
    r=sr+2+i
    ws.merge_cells(f"G{r}:H{r}")
    lc=ws[f"G{r}"]; lc.value=grp; lc.font=font(10); lc.border=box(LINE); lc.alignment=align("left")
    ws[f"H{r}"].border=box(LINE)
    vc=ws[f"I{r}"]
    vc.value=f"=SUMIF({R['bens_grp']},G{r},{R['bens_val']})"
    vc.number_format=CUR; vc.font=font(10); vc.border=box(LINE); vc.alignment=align("right")
    if (i%2)==1:
        for L in ("G","H","I"): ws[f"{L}{r}"].fill=fill(LIGHT)

# ---- Controle de documentos (barra de progresso)
pr = sr+2+len(LISTS["GRUPO_BENS"])+1
ws.merge_cells(f"G{pr}:I{pr}")
ws[f"G{pr}"]="CONTROLE DE DOCUMENTOS"
ws[f"G{pr}"].font=font(11,True,WHITE); ws[f"G{pr}"].fill=fill(NAVY2); ws[f"G{pr}"].alignment=align("center")
docrows=[
 ("Cadastrados", f"=COUNTA({R['doc_nome']})","0"),
 ("Concluidos (Anexado/Recebido)", f"=COUNTIF({R['doc_stat']},\"Anexado\")+COUNTIF({R['doc_stat']},\"Recebido\")","0"),
 ("Pendentes/Solicitados", f"=COUNTIF({R['doc_stat']},\"Pendente\")+COUNTIF({R['doc_stat']},\"Solicitado\")","0"),
 ("% de conclusao", f"=IFERROR((COUNTIF({R['doc_stat']},\"Anexado\")+COUNTIF({R['doc_stat']},\"Recebido\"))/COUNTA({R['doc_nome']}),0)",PCT),
]
for i,(lab,fm,nf) in enumerate(docrows):
    r=pr+1+i
    ws.merge_cells(f"G{r}:H{r}")
    lc=ws[f"G{r}"]; lc.value=lab; lc.font=font(10); lc.border=box(LINE); lc.alignment=align("left")
    ws[f"H{r}"].border=box(LINE)
    vc=ws[f"I{r}"]; vc.value=fm; vc.number_format=nf; vc.font=font(10,True,NAVY)
    vc.border=box(LINE); vc.alignment=align("center")
# data bar na % de conclusao
from openpyxl.formatting.rule import DataBarRule
pct_cell=f"I{pr+4}"
ws.conditional_formatting.add(pct_cell, DataBarRule(start_type="num", start_value=0,
    end_type="num", end_value=1, color=GREEN))
note(ws, f"B{pr+2}", "Dashboard: totais calculados por SOMA, SOMASE (SUMIF) e CONT.SE/CONT.VALORES sobre as demais abas.")

print("RESUMO ok")

# ================================================================== FINALIZAR
order = ["INICIO","DADOS PESSOAIS","RENDIMENTOS","BENS E DIREITOS","DIVIDAS E ONUS",
         "PAGAMENTOS","DEPENDENTES","DOCUMENTOS","RESUMO","LISTAS"]
wb._sheets.sort(key=lambda s: order.index(s.title))
wb.active = order.index("INICIO")

# recalcular ao abrir
try:
    wb.calculation.fullCalcOnLoad = True
except Exception:
    from openpyxl.workbook.properties import CalcProperties
    wb.calculation = CalcProperties(fullCalcOnLoad=True)

# metadados
wb.properties.title = "Agregador de Imposto de Renda"
wb.properties.creator = "Bel - Desafio DIO"
wb.properties.description = "Ferramenta Excel para organizar informacoes da declaracao de IRPF."

import os
os.makedirs("/mnt/user-data/outputs", exist_ok=True)
for path in ["/home/claude/agregador_ir.xlsx","/mnt/user-data/outputs/agregador_ir.xlsx"]:
    wb.save(path)
print("SAVED:", os.path.getsize("/home/claude/agregador_ir.xlsx"), "bytes")
