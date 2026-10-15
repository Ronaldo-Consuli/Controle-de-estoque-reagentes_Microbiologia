"""Monitoramento mensal de estoque — fonte para GitHub/Streamlit.

Instalação: pip install -r requirements.txt
Execução: streamlit run monitoramento.py
No Streamlit, selecione este arquivo como arquivo principal (Python 3.12).

Importa Excel .xlsx no formato Forms: uma linha por resposta, uma coluna
por reagente iniciada por RE, coluna Mês e datas de resposta opcionais.
Conserva todas as abas do arquivo original sem modificá-lo. A visualização
transforma as respostas em estoque mensal; não calcula consumo.
Para manter histórico, importe o Excel acumulado ou todos os arquivos mensais.
Os uploads não são gravados no servidor nem no GitHub. Ao encerrar a sessão,
reimporte os arquivos. Exportações são relatórios, não backups dos originais.

Classificação administrativa exclusivamente pela referência do usuário.
Itens ausentes recebem —, inclusive misturas comerciais não listadas.
Líquidos cadastrados com cabeçalho g têm unidade corrigida para mL sem
alteração do valor: o usuário confirmou que as medições já são volumes.
"""

from __future__ import annotations

import csv
import hashlib
import html
import io
import json
import math
import re
import unicodedata

from datetime import date, datetime

import openpyxl
import streamlit as st
import streamlit.components.v1 as components


def normalizar(value):
    text = unicodedata.normalize("NFKD", str(value or "")).encode("ascii", "ignore").decode()
    return re.sub(r"[^a-z0-9]+", " ", text.lower()).strip()


# Produto | Polícia Federal | Polícia Civil | Exército. 1 = Sim.
REFERENCIA = """Acetaldeído|1|1|0
Acetato de Etila|1|1|0
Acetona|1|1|0
Ácido Acético|1|1|0
Ácido Benzílico|0|1|1
Ácido Benzóico|1|1|0
Ácido Bórico|1|1|0
Ácido Clorídrico|1|1|0
Ácido Fórmico|1|1|0
Ácido Fosfórico|1|1|0
Ácido Nítrico|0|1|1
Ácido Perclórico|0|1|1
Ácido Sulfúrico 98%|1|1|0
Ácido Sulfúrico Fumegante|1|1|0
Ácido Sulfúrico Merck|1|1|0
Álcool Etílico|1|0|0
Álcool Isobutílico|1|1|0
Álcool Isopropílico|1|1|0
Álcool Metílico|1|1|0
Álcool N-butílico|1|1|0
Álcool N-propílico|0|1|0
Álcool Sec-butílico|1|1|0
Amônia|1|1|0
Anidrido Acético|1|1|0
Azída de Sódio|0|1|1
Benzeno|1|1|0
Bicarbonato de Potássio|1|1|0
Bicarbonato de Sódio|1|1|0
Bicarbonato de Sódio Anidro|1|1|0
Borohidreto de Sódio|1|1|0
Carbonato de Sódio|1|1|0
Cianeto de Potássio|0|1|1
Ciclohexano|1|1|0
Clorato de Potássio|0|1|1
Cloreto de Alumínio|1|1|0
Cloreto de Amônio|1|1|0
Cloreto de Etíla|1|1|0
Cloreto de Mercúrio|1|1|0
Cloreto de Mercúrio II|1|1|0
Cloreto de Metíleno|1|1|0
Cloreto de Potássio|0|1|0
Cloridrato de Hidroxilamina|0|1|0
Clorofórmio|1|1|0
Cromato de Potássio|1|0|0
Dicromato de Potássio|1|1|0
Dicromato de Sódio|1|1|0
Dietanolamina|0|1|0
Dióxido de Cloro|1|1|0
Enxofre (Lentilhado)|0|1|0
Éter Etílico|1|1|0
Etilenoglicol|0|1|0
Gasolina|1|0|0
Glicerina|0|1|0
Glicerol|0|1|0
Hidróxido de Amônio|1|1|0
Hidróxido de Cálcio|1|0|0
Hidróxido de Potássio|0|1|0
Hidróxido de Sódio|1|1|0
Hidroxilamina|1|1|0
Hipoclorito de Sódio|1|1|0
Manitol|1|1|0
Metanol|1|1|0
Nitrato de Amônio|0|1|1
Óleo Diesel|1|0|0
Óxido de cálcio|1|0|0
Perclorato de Sódio|0|1|0
Permanganato de Potássio|1|1|0
Peróxido de Hidrogênio|1|1|0
Piridina|1|1|0
Querosene|1|0|0
Soda Cáustica (escamas)|1|1|0
Soda Cáustica (líquida)|1|1|0
Sulfato de Magnésio|0|1|0
Sulfato de Sódio|1|1|0
Tolueno|1|1|0
Trietanolamina|0|1|1
Ureia|1|0|0
Xileno|1|1|0"""
BASE = {}
for line in REFERENCIA.splitlines():
    name, pf, pc, eb = line.split("|")
    BASE[normalizar(name)] = (name, pf == "1", pc == "1", eb == "1")

# Correspondências explícitas dos nomes usados no inventário fornecido.
ALIASES = {
    "ácido acético glacial": "ácido acético",
    "ácido sulfúrico": "ácido sulfúrico 98%",
    "álcool etílico absoluto": "álcool etílico",
    "álcool iso-butílico": "álcool isobutílico",
    "1 - propanol": "álcool n-propílico",
    "carbonato de sódio anidro": "carbonato de sódio",
    "bicarbonato de sódio sodium hydrogen carbonate": "bicarbonato de sódio",
    "hidróxido de sódio em pérolas": "hidróxido de sódio",
    "solução hidróxido de sódio 50%": "hidróxido de sódio",
    "solução cloreto de potássio 3m": "cloreto de potássio",
}
ALIASES = {normalizar(k): normalizar(v) for k, v in ALIASES.items()}


def classificar(label):
    clean = re.sub(r"^\s*RE\s*\d+(?:-\d+)?\s*-\s*", "", str(label), flags=re.I)
    clean = re.sub(r"\((?:g|kg|mg|ml|l)\)\s*$", "", clean, flags=re.I)
    key = normalizar(clean)
    key = re.sub(r"\s+p a$", "", key)
    match = BASE.get(ALIASES.get(key, key))
    if match:
        ref, pf, pc, eb = match
        return {"referencia": ref, "pf": pf, "pc": pc, "eb": eb}
    return {"referencia": "—", "pf": False, "pc": False, "eb": False}


MESES = ["janeiro", "fevereiro", "marco", "abril", "maio", "junho", "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"]

# Cadastro de líquidos/soluções do inventário. Não é conversão por densidade.
# Reagentes comerciais de formulação não identificada não são inferidos.
LIQUIDOS_ML = {
    "332", "701", "701-1", "702", "703", "705", "705-1", "706", "707", "709",
    "800", "804", "806", "808", "810", "812", "820", "830", "850", "851",
    "860", "861", "870", "880", "881", "890", "910", "921", "923", "930",
}


def data_resposta(value):
    if isinstance(value, datetime):
        return value.replace(tzinfo=None)
    if isinstance(value, date):
        return datetime.combine(value, datetime.min.time())
    for fmt in ("%d/%m/%Y %H:%M:%S", "%d/%m/%Y %H:%M", "%d/%m/%Y", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d"):
        try:
            return datetime.strptime(str(value), fmt)
        except ValueError:
            pass
    return None


def periodo(value, ano):
    if isinstance(value, (datetime, date)):
        return value.strftime("%Y-%m")
    text = normalizar(value)
    nums = re.findall(r"\d+", text)
    year = next((int(n) for n in nums if len(n) == 4), ano)
    month = None
    for i, name in enumerate(MESES, 1):
        if name in text.split() or name[:3] in text.split():
            month = i
            break
    if month is None:
        month = next((int(n) for n in nums if len(n) <= 2 and 1 <= int(n) <= 12), None)
    if month is None:
        raise ValueError("Mês não reconhecido")
    return date(year, month, 1).strftime("%Y-%m")


def quantidade(value):
    if value is None or str(value).strip() == "":
        return None
    if isinstance(value, bool):
        raise ValueError("Quantidade inválida")
    if isinstance(value, (int, float)):
        result = float(value)
    else:
        text = str(value).strip().replace("\u00a0", "").replace(" ", "")
        # Texto segue a convenção brasileira; células numéricas não são alteradas.
        if "," in text:
            text = text.replace(".", "").replace(",", ".")
        elif re.fullmatch(r"\d{1,3}(?:\.\d{3})+", text):
            text = text.replace(".", "")
        result = float(text)
    if not math.isfinite(result) or result < 0:
        raise ValueError("Quantidade deve ser finita e não negativa")
    return result


def ler_excel(blob, arquivo, local, ano_padrao, unidade_padrao, liquidos_ml=None):
    """Retorna respostas completas; nunca soma duas respostas do mesmo local/mês."""
    liquidos_ml = LIQUIDOS_ML if liquidos_ml is None else set(liquidos_ml)
    wb = openpyxl.load_workbook(io.BytesIO(blob), data_only=True, read_only=True)
    snapshots, avisos = [], []
    try:
        for ws in wb:
            # Alguns exports Forms declaram dimensão A1 apesar de conterem
            # centenas de colunas; percorra o XML real, não essa dimensão.
            ws.reset_dimensions()
            rows = ws.iter_rows(values_only=True)
            header = next(rows, ())
            # Formato Forms: cabeçalhos na primeira linha.
            columns = [(i, str(h).strip()) for i, h in enumerate(header) if h is not None and re.match(r"^RE\s", str(h).strip(), re.I)]
            if not columns:
                continue
            names = {normalizar(h): i for i, h in enumerate(header) if h is not None}
            def get(row, name):
                idx = names.get(normalizar(name))
                return row[idx] if idx is not None and idx < len(row) else None
            for rownum, row in enumerate(rows, 2):
                if not any(v is not None for v in row):
                    continue
                row_local = str(get(row, "Local") or get(row, "Armário") or local).strip()
                stamp = next((d for n in ("Hora da última modificação", "Hora de conclusão", "Hora de início") if (d := data_resposta(get(row, n)))), None)
                year = stamp.year if stamp else ano_padrao
                if get(row, "Ano") is not None:
                    try:
                        year = int(get(row, "Ano"))
                    except (ValueError, TypeError):
                        avisos.append(f"{arquivo}/{ws.title}, linha {rownum}: ano inválido; resposta ignorada.")
                        continue
                month_value = get(row, "Mês")
                try:
                    month = periodo(month_value if month_value is not None else stamp, year)
                except (ValueError, TypeError):
                    avisos.append(f"{arquivo}/{ws.title}, linha {rownum}: mês inválido; resposta ignorada.")
                    continue
                items = []
                seen = set()
                for idx, label in columns:
                    code = re.match(r"^RE\s*(\d+(?:-\d+)?)\s*-", label, re.I)
                    key = "RE " + code.group(1) if code else normalizar(label)
                    if key in seen:
                        raise ValueError(f"{arquivo}/{ws.title}: reagente duplicado no cabeçalho: {label}")
                    seen.add(key)
                    unit_match = re.search(r"\((g|kg|mg|ml|l)\)\s*$", label, re.I)
                    unit = unit_match.group(1).lower() if unit_match else unidade_padrao.lower()
                    corrected = bool(code and code.group(1) in liquidos_ml and (not unit_match or unit == "g"))
                    if corrected:
                        unit = "ml"
                    target, factor = {"g": ("g", 1), "kg": ("g", 1000), "mg": ("g", .001), "ml": ("mL", 1), "l": ("mL", 1000)}[unit]
                    if not unit_match and not corrected:
                        avisos.append(f"{arquivo}/{ws.title}: {label} sem unidade; usada {unidade_padrao}.")
                    try:
                        qty = quantidade(row[idx] if idx < len(row) else None)
                        qty = qty * factor if qty is not None else None
                    except (ValueError, TypeError):
                        qty = None
                        avisos.append(f"{arquivo}/{ws.title}, linha {rownum}: quantidade inválida em {label}; marcada não informada.")
                    display_label = (label[:unit_match.start()].strip() if unit_match else label.strip()) + f" ({target})"
                    items.append({"id": key, "produto": display_label, "produto_original": label,
                                  "unidade_corrigida": corrected, "mes": month, "local": row_local, "aba": ws.title,
                                  "quantidade": qty, "unidade": target, **classificar(label)})
                snapshots.append({"local": row_local, "aba": ws.title, "mes": month, "data": stamp or datetime.min,
                                  "linha": rownum, "arquivo": arquivo, "itens": items})
    finally:
        wb.close()
    if not snapshots:
        avisos.append(f"{arquivo}: nenhuma resposta válida com colunas RE foi encontrada.")
    return snapshots, list(dict.fromkeys(avisos))


def consolidar(snapshots):
    latest, avisos = {}, []
    for s in snapshots:
        # Cada aba representa uma seção de inventário independente.
        key = (s["local"], s["aba"], s["mes"])
        old = latest.get(key)
        if old:
            avisos.append(f"{s['local']} / {s['aba']} / {s['mes']}: respostas repetidas; mantida a mais recente. Em empate, a última importada.")
        if old is None or (s["data"], s["linha"]) >= (old["data"], old["linha"]):
            latest[key] = s
    return [i for s in latest.values() for i in s["itens"]], list(dict.fromkeys(avisos))


STYLE = """
<style>
*{box-sizing:border-box}body{font-family:Arial,sans-serif;background:#F1F5F8;color:#20364B;margin:0;padding:20px}
.wrap{max-width:1500px;margin:auto}.header{background:linear-gradient(115deg,#087F62,#075783 65%,#113D61);border-radius:18px;padding:28px 32px;color:white;box-shadow:0 8px 28px #14325718}
.eyebrow{text-transform:uppercase;letter-spacing:2px;font-size:10px;font-weight:700;margin-bottom:12px;color:#C9EEE5}h1{font-size:30px;letter-spacing:-.7px;margin:0 0 9px}h2{font-size:18px;letter-spacing:-.2px;color:#143257;margin:0 0 8px}.note{font-size:12px;color:#64788B;line-height:1.6}.header .note{color:#DBEBF3;margin-top:12px}
.panel,.card{background:white;border:1px solid #E1E8EF;border-radius:14px;box-shadow:0 3px 14px #183C6106;padding:22px;margin-top:18px}
.panel-head{display:flex;align-items:center;justify-content:space-between;gap:14px;margin-bottom:18px}.panel-head h2{margin:0}.filters{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:16px}label{font-size:12px;font-weight:bold;color:#526B81}select,input{display:block;margin-top:7px;width:100%;height:44px;border:1px solid #D6E0E9;border-radius:8px;padding:10px;background:#FAFCFE;color:#20364B;font-size:13px}select:focus,input:focus,button:focus-visible{outline:3px solid #91CDDF;outline-offset:2px}
button{border:1px solid #D6E0E9;background:white;color:#245775;border-radius:8px;padding:10px 14px;cursor:pointer;font-weight:bold;font-size:12px}button:hover{background:#EFF6FA}.kpis{display:grid;grid-template-columns:repeat(5,minmax(0,1fr));gap:14px}.secondary{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:14px}.card{text-align:center;border-top:3px solid #D2DFE8;padding:20px}.card:nth-child(2){border-top-color:#1167A4}.card:nth-child(3){border-top-color:#009878}.value{white-space:pre-line;font-size:32px;font-weight:700;letter-spacing:-1px;color:#143257;margin-top:12px;overflow-wrap:anywhere}.label{font-size:12px;color:#526B81}.secondary .card{display:flex;flex-direction:column;justify-content:center;align-items:center;gap:12px;border-top:1px solid #E1E8EF;margin-top:14px;padding:16px 22px}.secondary .value{font-size:26px;margin:0}.context{margin:22px 2px -6px;color:#526B81;font-size:12px}.context strong{color:#143257}
.scroll{max-height:480px;overflow:auto;border:1px solid #E1E8EF;border-radius:10px}table{border-collapse:collapse;width:100%;font-size:12px}th{background:#143257;color:white;position:sticky;top:0;z-index:1;white-space:nowrap}th,td{padding:14px 12px;border-bottom:1px solid #E8ECF0;text-align:center}td:first-child{text-align:left;min-width:260px}tbody tr:nth-child(even){background:#F7FAFC}tbody tr:hover{background:#ECF5FA}.yes{color:#00795F;font-weight:bold}.empty{color:#A65B00}.zero{color:#9B3E26;font-weight:bold}.badge{display:inline-block;border-radius:20px;padding:5px 9px;font-size:11px;white-space:nowrap}.controlled{background:#EAF2FC;color:#155C99}.uncontrolled{background:#E5F5EE;color:#16724F}.emptybox{padding:34px 18px;text-align:center;color:#738596;background:#F8FAFC;border-radius:10px;font-size:13px}
@media(max-width:1000px){.kpis{grid-template-columns:repeat(3,minmax(0,1fr))}.value{font-size:28px}}
@media(max-width:650px){body{padding:8px}.header{padding:23px}.filters{grid-template-columns:1fr 1fr}.kpis{grid-template-columns:1fr 1fr}.secondary{grid-template-columns:1fr}.secondary .card{margin-top:8px}.panel{padding:16px}.panel-head{align-items:flex-start;flex-direction:column}h1{font-size:25px}}
@media(max-width:650px){.filters{grid-template-columns:1fr}}
</style>
"""


def _painel_html_base(records):
    """Exporta um painel HTML interativo, autônomo e com CSV do estoque mensal."""
    payload = json.dumps(records, ensure_ascii=False, allow_nan=False).replace("<", "\\u003c").replace(">", "\\u003e").replace("&", "\\u0026")
    return "<!doctype html><html lang='pt-BR'><head><meta charset='utf-8'><meta name='viewport' content='width=device-width, initial-scale=1'><title>Monitoramento de reagentes</title>" + STYLE + """</head>
<body><div class="wrap">
<div class="header"><div class="eyebrow">Laboratório industrial · Gestão de estoque</div><h1>Monitoramento de reagentes</h1><div>Estoque mensal do laboratório</div></div>
<div class="panel"><div class="panel-head"><h2>Filtros de acompanhamento</h2><button id="limpar" type="button">Limpar filtros</button></div><div class="filters">
<label>Classificação<select id="classe"><option value="">Todas</option><option value="sim">Controlados</option><option value="nao">Não controlados</option></select></label>
<label>Órgão<select id="orgao"><option value="">Todos</option><option value="pf">Polícia Federal</option><option value="pc">Polícia Civil</option><option value="eb">Exército</option><option value="nao">Não controlado</option></select></label>
</div></div>
<div class="context" id="contexto"></div><div class="kpis" id="cards"></div><div class="secondary" id="agencias"></div>
<div class="panel"><div class="panel-head"><div><h2>Estoque mensal</h2><div class="note" id="linhasTabela"></div></div><button id="baixarTabela" type="button">↓ Baixar tabela mensal (CSV)</button></div>
<div class="scroll"><table><thead><tr id="cabecalhoTabela"></tr></thead><tbody id="tabela"></tbody></table></div></div>
</div>
<script>const dados=""" + payload + """;
const $=id=>document.getElementById(id);
const fmt=n=>n.toLocaleString('pt-BR',{maximumFractionDigits:3});
const control=r=>r.pf||r.pc||r.eb;
const esc=s=>String(s).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const uniq=a=>[...new Set(a)];
const months=uniq(dados.map(r=>r.mes)).sort();
const monthLatest=months.at(-1)||'';
const cycleOf=m=>Number(m.slice(0,4))-(Number(m.slice(5))<4?1:0);
const cycle=monthLatest?cycleOf(monthLatest):new Date().getFullYear();
const cycleMonths=months.filter(m=>cycleOf(m)===cycle);
const monthNames=['Janeiro','Fevereiro','Março','Abril','Maio','Junho','Julho','Agosto','Setembro','Outubro','Novembro','Dezembro'];
const monthLabel=m=>m.slice(5)+'/'+m.slice(0,4);
const card=(label,value)=>'<div class="card"><div class="label">'+esc(label)+'</div><div class="value">'+esc(value)+'</div></div>';
const csvEscape=value=>{
  let s=String(value??'');
  // Evita execução de fórmulas se o arquivo for aberto em uma planilha.
  if (/^\\s*[=+@-]/.test(s)) s="'"+s;
  return '"'+s.replace(/"/g,'""')+'"';
};
let csvMensal='';
function render(){
  const base=dados.filter(r=>(!$('orgao').value||($('orgao').value==='nao'?!control(r):r[$('orgao').value]))&&(!$('classe').value||control(r)===($('classe').value==='sim')));
  const rows=base.filter(r=>r.mes===monthLatest);
  const count=rs=>uniq(rs.map(r=>r.id)).length;
  const estoque=u=>fmt(rows.filter(r=>r.unidade===u).reduce((s,r)=>s+(r.quantidade??0),0));
  $('contexto').textContent='Referência: '+(monthLatest?monthLabel(monthLatest):'—')+' · '+count(rows)+' reagentes na seleção';
  $('cards').innerHTML=card('Reagentes',count(rows))+card('Controlados',count(rows.filter(control)))+card('Não controlados',count(rows.filter(r=>!control(r))))+card('Estoque (g)',estoque('g'))+card('Estoque (mL)',estoque('mL'));
  $('agencias').innerHTML=['pf','pc','eb'].map((k,i)=>card(['Polícia Federal','Polícia Civil','Exército'][i],count(rows.filter(r=>r[k])))).join('')+card('Não controlados',count(rows.filter(r=>!control(r))));
  const matrix=new Map();
  base.filter(r=>cycleOf(r.mes)===cycle).forEach(r=>{
    const key=JSON.stringify([r.id,r.unidade]);
    if(!matrix.has(key))matrix.set(key,{...r,cells:new Map()});
    const g=matrix.get(key);
    ['pf','pc','eb'].forEach(k=>g[k]=g[k]||r[k]);
    if(!g.cells.has(r.mes))g.cells.set(r.mes,{v:0,n:0,missing:false});
    const cell=g.cells.get(r.mes);
    if(r.quantidade===null)cell.missing=true;
    else{cell.v+=r.quantidade;cell.n++;}
  });
  const tableRows=[...matrix.values()];
  $('linhasTabela').textContent=tableRows.length+' reagentes · Abril/'+cycle+' a março/'+(cycle+1)+' · '+cycleMonths.length+' meses importados';
  const headers=['Reagente','Classificação','Federal','Civil','Exército',...cycleMonths.map(m=>monthNames[Number(m.slice(5))-1]+'/'+m.slice(0,4))];
  $('cabecalhoTabela').innerHTML=headers.map(h=>'<th>'+esc(h)+'</th>').join('');
  const tableValue=(r,m)=>{const c=r.cells.get(m);return !c?.n?'—':fmt(c.v)+(c.missing?'*':'');};
  $('tabela').innerHTML=tableRows.map(r=>'<tr><td>'+esc(r.produto)+'</td><td><span class="badge '+(control(r)?'controlled':'uncontrolled')+'">'+(control(r)?'Controlado':'Não controlado')+'</span></td>'+['pf','pc','eb'].map(k=>'<td class="'+(r[k]?'yes':'')+'">'+(r[k]?'Sim':'—')+'</td>').join('')+cycleMonths.map(m=>{const c=r.cells.get(m);return '<td class="'+(!c?.n?'empty':c.v===0?'zero':'')+'">'+tableValue(r,m)+'</td>';}).join('')+'</tr>').join('')||'<tr><td colspan="'+headers.length+'">Nenhum reagente para os filtros selecionados.</td></tr>';
  // O download corresponde à mesma tabela exibida, respeitando os dois filtros.
  const csvRows=[headers,...tableRows.map(r=>[
    r.produto,control(r)?'Controlado':'Não controlado',
    ...['pf','pc','eb'].map(k=>r[k]?'Sim':'—'),
    ...cycleMonths.map(m=>tableValue(r,m))
  ])];
  csvMensal=csvRows.map(row=>row.map(csvEscape).join(';')).join('\\r\\n');
}
$('classe').addEventListener('change',render);
$('orgao').addEventListener('change',render);
$('limpar').addEventListener('click',()=>{$('classe').value='';$('orgao').value='';render();});
$('baixarTabela').addEventListener('click',()=>{
  const blob=new Blob(['\\uFEFF'+csvMensal],{type:'text/csv;charset=utf-8'});
  const url=URL.createObjectURL(blob);
  const a=document.createElement('a');
  a.href=url;a.download='estoque_mensal_'+cycle+'_'+(cycle+1)+'.csv';
  document.body.appendChild(a);a.click();a.remove();
  setTimeout(()=>URL.revokeObjectURL(url),1000);
});
render();
</script></body></html>"""


# Gerador QR incorporado ao HTML exportado: funciona sem CDN e sem Streamlit.
# Biblioteca QRCode for JavaScript © 2009 Kazuhiko Arase (MIT).
QR_LIB_JS = '/* QR generator (c) Kazuhiko Arase, MIT license; bundled from QRCode for JavaScript. */\n(function(global){\n"use strict";\nconst modules = {\n"./QRMode": function(module, exports, require) {\nmodule.exports = {\n    MODE_NUMBER :       1 << 0,\n    MODE_ALPHA_NUM :    1 << 1,\n    MODE_8BIT_BYTE :    1 << 2,\n    MODE_KANJI :        1 << 3\n};\n\n},\n"./QRErrorCorrectLevel": function(module, exports, require) {\nmodule.exports = {\n\tL : 1,\n\tM : 0,\n\tQ : 3,\n\tH : 2\n};\n\n\n},\n"./QRMaskPattern": function(module, exports, require) {\nmodule.exports = {\n\tPATTERN000 : 0,\n\tPATTERN001 : 1,\n\tPATTERN010 : 2,\n\tPATTERN011 : 3,\n\tPATTERN100 : 4,\n\tPATTERN101 : 5,\n\tPATTERN110 : 6,\n\tPATTERN111 : 7\n};\n\n},\n"./QRMath": function(module, exports, require) {\nvar QRMath = {\n\n\tglog : function(n) {\n\t\n\t\tif (n < 1) {\n\t\t\tthrow new Error("glog(" + n + ")");\n\t\t}\n\t\t\n\t\treturn QRMath.LOG_TABLE[n];\n\t},\n\t\n\tgexp : function(n) {\n\t\n\t\twhile (n < 0) {\n\t\t\tn += 255;\n\t\t}\n\t\n\t\twhile (n >= 256) {\n\t\t\tn -= 255;\n\t\t}\n\t\n\t\treturn QRMath.EXP_TABLE[n];\n\t},\n\t\n\tEXP_TABLE : new Array(256),\n\t\n\tLOG_TABLE : new Array(256)\n\n};\n\t\nfor (var i = 0; i < 8; i++) {\n\tQRMath.EXP_TABLE[i] = 1 << i;\n}\nfor (var i = 8; i < 256; i++) {\n\tQRMath.EXP_TABLE[i] = QRMath.EXP_TABLE[i - 4]\n\t\t^ QRMath.EXP_TABLE[i - 5]\n\t\t^ QRMath.EXP_TABLE[i - 6]\n\t\t^ QRMath.EXP_TABLE[i - 8];\n}\nfor (var i = 0; i < 255; i++) {\n\tQRMath.LOG_TABLE[QRMath.EXP_TABLE[i] ] = i;\n}\n\nmodule.exports = QRMath;\n\n},\n"./QRPolynomial": function(module, exports, require) {\nvar QRMath = require(\'./QRMath\');\n\nfunction QRPolynomial(num, shift) {\n\tif (num.length === undefined) {\n\t\tthrow new Error(num.length + "/" + shift);\n\t}\n\n\tvar offset = 0;\n\n\twhile (offset < num.length && num[offset] === 0) {\n\t\toffset++;\n\t}\n\n\tthis.num = new Array(num.length - offset + shift);\n\tfor (var i = 0; i < num.length - offset; i++) {\n\t\tthis.num[i] = num[i + offset];\n\t}\n}\n\nQRPolynomial.prototype = {\n\n\tget : function(index) {\n\t\treturn this.num[index];\n\t},\n\t\n\tgetLength : function() {\n\t\treturn this.num.length;\n\t},\n\t\n\tmultiply : function(e) {\n\t\n\t\tvar num = new Array(this.getLength() + e.getLength() - 1);\n\t\n\t\tfor (var i = 0; i < this.getLength(); i++) {\n\t\t\tfor (var j = 0; j < e.getLength(); j++) {\n\t\t\t\tnum[i + j] ^= QRMath.gexp(QRMath.glog(this.get(i) ) + QRMath.glog(e.get(j) ) );\n\t\t\t}\n\t\t}\n\t\n\t\treturn new QRPolynomial(num, 0);\n\t},\n\t\n\tmod : function(e) {\n\t\n\t\tif (this.getLength() - e.getLength() < 0) {\n\t\t\treturn this;\n\t\t}\n\t\n\t\tvar ratio = QRMath.glog(this.get(0) ) - QRMath.glog(e.get(0) );\n\t\n\t\tvar num = new Array(this.getLength() );\n\t\t\n\t\tfor (var i = 0; i < this.getLength(); i++) {\n\t\t\tnum[i] = this.get(i);\n\t\t}\n\t\t\n\t\tfor (var x = 0; x < e.getLength(); x++) {\n\t\t\tnum[x] ^= QRMath.gexp(QRMath.glog(e.get(x) ) + ratio);\n\t\t}\n\t\n\t\t// recursive call\n\t\treturn new QRPolynomial(num, 0).mod(e);\n\t}\n};\n\nmodule.exports = QRPolynomial;\n\n},\n"./QRRSBlock": function(module, exports, require) {\nvar QRErrorCorrectLevel = require(\'./QRErrorCorrectLevel\');\n\nfunction QRRSBlock(totalCount, dataCount) {\n\tthis.totalCount = totalCount;\n\tthis.dataCount  = dataCount;\n}\n\nQRRSBlock.RS_BLOCK_TABLE = [\n\n\t// L\n\t// M\n\t// Q\n\t// H\n\n\t// 1\n\t[1, 26, 19],\n\t[1, 26, 16],\n\t[1, 26, 13],\n\t[1, 26, 9],\n\t\n\t// 2\n\t[1, 44, 34],\n\t[1, 44, 28],\n\t[1, 44, 22],\n\t[1, 44, 16],\n\n\t// 3\n\t[1, 70, 55],\n\t[1, 70, 44],\n\t[2, 35, 17],\n\t[2, 35, 13],\n\n\t// 4\t\t\n\t[1, 100, 80],\n\t[2, 50, 32],\n\t[2, 50, 24],\n\t[4, 25, 9],\n\t\n\t// 5\n\t[1, 134, 108],\n\t[2, 67, 43],\n\t[2, 33, 15, 2, 34, 16],\n\t[2, 33, 11, 2, 34, 12],\n\t\n\t// 6\n\t[2, 86, 68],\n\t[4, 43, 27],\n\t[4, 43, 19],\n\t[4, 43, 15],\n\t\n\t// 7\t\t\n\t[2, 98, 78],\n\t[4, 49, 31],\n\t[2, 32, 14, 4, 33, 15],\n\t[4, 39, 13, 1, 40, 14],\n\t\n\t// 8\n\t[2, 121, 97],\n\t[2, 60, 38, 2, 61, 39],\n\t[4, 40, 18, 2, 41, 19],\n\t[4, 40, 14, 2, 41, 15],\n\t\n\t// 9\n\t[2, 146, 116],\n\t[3, 58, 36, 2, 59, 37],\n\t[4, 36, 16, 4, 37, 17],\n\t[4, 36, 12, 4, 37, 13],\n\t\n\t// 10\t\t\n\t[2, 86, 68, 2, 87, 69],\n\t[4, 69, 43, 1, 70, 44],\n\t[6, 43, 19, 2, 44, 20],\n\t[6, 43, 15, 2, 44, 16],\n\n\t// 11\n\t[4, 101, 81],\n\t[1, 80, 50, 4, 81, 51],\n\t[4, 50, 22, 4, 51, 23],\n\t[3, 36, 12, 8, 37, 13],\n\n\t// 12\n\t[2, 116, 92, 2, 117, 93],\n\t[6, 58, 36, 2, 59, 37],\n\t[4, 46, 20, 6, 47, 21],\n\t[7, 42, 14, 4, 43, 15],\n\n\t// 13\n\t[4, 133, 107],\n\t[8, 59, 37, 1, 60, 38],\n\t[8, 44, 20, 4, 45, 21],\n\t[12, 33, 11, 4, 34, 12],\n\n\t// 14\n\t[3, 145, 115, 1, 146, 116],\n\t[4, 64, 40, 5, 65, 41],\n\t[11, 36, 16, 5, 37, 17],\n\t[11, 36, 12, 5, 37, 13],\n\n\t// 15\n\t[5, 109, 87, 1, 110, 88],\n\t[5, 65, 41, 5, 66, 42],\n\t[5, 54, 24, 7, 55, 25],\n\t[11, 36, 12],\n\n\t// 16\n\t[5, 122, 98, 1, 123, 99],\n\t[7, 73, 45, 3, 74, 46],\n\t[15, 43, 19, 2, 44, 20],\n\t[3, 45, 15, 13, 46, 16],\n\n\t// 17\n\t[1, 135, 107, 5, 136, 108],\n\t[10, 74, 46, 1, 75, 47],\n\t[1, 50, 22, 15, 51, 23],\n\t[2, 42, 14, 17, 43, 15],\n\n\t// 18\n\t[5, 150, 120, 1, 151, 121],\n\t[9, 69, 43, 4, 70, 44],\n\t[17, 50, 22, 1, 51, 23],\n\t[2, 42, 14, 19, 43, 15],\n\n\t// 19\n\t[3, 141, 113, 4, 142, 114],\n\t[3, 70, 44, 11, 71, 45],\n\t[17, 47, 21, 4, 48, 22],\n\t[9, 39, 13, 16, 40, 14],\n\n\t// 20\n\t[3, 135, 107, 5, 136, 108],\n\t[3, 67, 41, 13, 68, 42],\n\t[15, 54, 24, 5, 55, 25],\n\t[15, 43, 15, 10, 44, 16],\n\n\t// 21\n\t[4, 144, 116, 4, 145, 117],\n\t[17, 68, 42],\n\t[17, 50, 22, 6, 51, 23],\n\t[19, 46, 16, 6, 47, 17],\n\n\t// 22\n\t[2, 139, 111, 7, 140, 112],\n\t[17, 74, 46],\n\t[7, 54, 24, 16, 55, 25],\n\t[34, 37, 13],\n\n\t// 23\n\t[4, 151, 121, 5, 152, 122],\n\t[4, 75, 47, 14, 76, 48],\n\t[11, 54, 24, 14, 55, 25],\n\t[16, 45, 15, 14, 46, 16],\n\n\t// 24\n\t[6, 147, 117, 4, 148, 118],\n\t[6, 73, 45, 14, 74, 46],\n\t[11, 54, 24, 16, 55, 25],\n\t[30, 46, 16, 2, 47, 17],\n\n\t// 25\n\t[8, 132, 106, 4, 133, 107],\n\t[8, 75, 47, 13, 76, 48],\n\t[7, 54, 24, 22, 55, 25],\n\t[22, 45, 15, 13, 46, 16],\n\n\t// 26\n\t[10, 142, 114, 2, 143, 115],\n\t[19, 74, 46, 4, 75, 47],\n\t[28, 50, 22, 6, 51, 23],\n\t[33, 46, 16, 4, 47, 17],\n\n\t// 27\n\t[8, 152, 122, 4, 153, 123],\n\t[22, 73, 45, 3, 74, 46],\n\t[8, 53, 23, 26, 54, 24],\n\t[12, 45, 15, 28, 46, 16],\n\n\t// 28\n\t[3, 147, 117, 10, 148, 118],\n\t[3, 73, 45, 23, 74, 46],\n\t[4, 54, 24, 31, 55, 25],\n\t[11, 45, 15, 31, 46, 16],\n\n\t// 29\n\t[7, 146, 116, 7, 147, 117],\n\t[21, 73, 45, 7, 74, 46],\n\t[1, 53, 23, 37, 54, 24],\n\t[19, 45, 15, 26, 46, 16],\n\n\t// 30\n\t[5, 145, 115, 10, 146, 116],\n\t[19, 75, 47, 10, 76, 48],\n\t[15, 54, 24, 25, 55, 25],\n\t[23, 45, 15, 25, 46, 16],\n\n\t// 31\n\t[13, 145, 115, 3, 146, 116],\n\t[2, 74, 46, 29, 75, 47],\n\t[42, 54, 24, 1, 55, 25],\n\t[23, 45, 15, 28, 46, 16],\n\n\t// 32\n\t[17, 145, 115],\n\t[10, 74, 46, 23, 75, 47],\n\t[10, 54, 24, 35, 55, 25],\n\t[19, 45, 15, 35, 46, 16],\n\n\t// 33\n\t[17, 145, 115, 1, 146, 116],\n\t[14, 74, 46, 21, 75, 47],\n\t[29, 54, 24, 19, 55, 25],\n\t[11, 45, 15, 46, 46, 16],\n\n\t// 34\n\t[13, 145, 115, 6, 146, 116],\n\t[14, 74, 46, 23, 75, 47],\n\t[44, 54, 24, 7, 55, 25],\n\t[59, 46, 16, 1, 47, 17],\n\n\t// 35\n\t[12, 151, 121, 7, 152, 122],\n\t[12, 75, 47, 26, 76, 48],\n\t[39, 54, 24, 14, 55, 25],\n\t[22, 45, 15, 41, 46, 16],\n\n\t// 36\n\t[6, 151, 121, 14, 152, 122],\n\t[6, 75, 47, 34, 76, 48],\n\t[46, 54, 24, 10, 55, 25],\n\t[2, 45, 15, 64, 46, 16],\n\n\t// 37\n\t[17, 152, 122, 4, 153, 123],\n\t[29, 74, 46, 14, 75, 47],\n\t[49, 54, 24, 10, 55, 25],\n\t[24, 45, 15, 46, 46, 16],\n\n\t// 38\n\t[4, 152, 122, 18, 153, 123],\n\t[13, 74, 46, 32, 75, 47],\n\t[48, 54, 24, 14, 55, 25],\n\t[42, 45, 15, 32, 46, 16],\n\n\t// 39\n\t[20, 147, 117, 4, 148, 118],\n\t[40, 75, 47, 7, 76, 48],\n\t[43, 54, 24, 22, 55, 25],\n\t[10, 45, 15, 67, 46, 16],\n\n\t// 40\n\t[19, 148, 118, 6, 149, 119],\n\t[18, 75, 47, 31, 76, 48],\n\t[34, 54, 24, 34, 55, 25],\n\t[20, 45, 15, 61, 46, 16]\n];\n\nQRRSBlock.getRSBlocks = function(typeNumber, errorCorrectLevel) {\n\t\n\tvar rsBlock = QRRSBlock.getRsBlockTable(typeNumber, errorCorrectLevel);\n\t\n\tif (rsBlock === undefined) {\n\t\tthrow new Error("bad rs block @ typeNumber:" + typeNumber + "/errorCorrectLevel:" + errorCorrectLevel);\n\t}\n\n\tvar length = rsBlock.length / 3;\n\t\n\tvar list = [];\n\t\n\tfor (var i = 0; i < length; i++) {\n\n\t\tvar count = rsBlock[i * 3 + 0];\n\t\tvar totalCount = rsBlock[i * 3 + 1];\n\t\tvar dataCount  = rsBlock[i * 3 + 2];\n\n\t\tfor (var j = 0; j < count; j++) {\n\t\t\tlist.push(new QRRSBlock(totalCount, dataCount) );\t\n\t\t}\n\t}\n\t\n\treturn list;\n};\n\nQRRSBlock.getRsBlockTable = function(typeNumber, errorCorrectLevel) {\n\n\tswitch(errorCorrectLevel) {\n\tcase QRErrorCorrectLevel.L :\n\t\treturn QRRSBlock.RS_BLOCK_TABLE[(typeNumber - 1) * 4 + 0];\n\tcase QRErrorCorrectLevel.M :\n\t\treturn QRRSBlock.RS_BLOCK_TABLE[(typeNumber - 1) * 4 + 1];\n\tcase QRErrorCorrectLevel.Q :\n\t\treturn QRRSBlock.RS_BLOCK_TABLE[(typeNumber - 1) * 4 + 2];\n\tcase QRErrorCorrectLevel.H :\n\t\treturn QRRSBlock.RS_BLOCK_TABLE[(typeNumber - 1) * 4 + 3];\n\tdefault :\n\t\treturn undefined;\n\t}\n};\n\nmodule.exports = QRRSBlock;\n\n},\n"./QRBitBuffer": function(module, exports, require) {\nfunction QRBitBuffer() {\n\tthis.buffer = [];\n\tthis.length = 0;\n}\n\nQRBitBuffer.prototype = {\n\n\tget : function(index) {\n\t\tvar bufIndex = Math.floor(index / 8);\n\t\treturn ( (this.buffer[bufIndex] >>> (7 - index % 8) ) & 1) == 1;\n\t},\n\t\n\tput : function(num, length) {\n\t\tfor (var i = 0; i < length; i++) {\n\t\t\tthis.putBit( ( (num >>> (length - i - 1) ) & 1) == 1);\n\t\t}\n\t},\n\t\n\tgetLengthInBits : function() {\n\t\treturn this.length;\n\t},\n\t\n\tputBit : function(bit) {\n\t\n\t\tvar bufIndex = Math.floor(this.length / 8);\n\t\tif (this.buffer.length <= bufIndex) {\n\t\t\tthis.buffer.push(0);\n\t\t}\n\t\n\t\tif (bit) {\n\t\t\tthis.buffer[bufIndex] |= (0x80 >>> (this.length % 8) );\n\t\t}\n\t\n\t\tthis.length++;\n\t}\n};\n\nmodule.exports = QRBitBuffer;\n\n},\n"./QR8bitByte": function(module, exports, require) {\nvar QRMode = require(\'./QRMode\');\n\nfunction QR8bitByte(data) {\n\tthis.mode = QRMode.MODE_8BIT_BYTE;\n\tthis.data = data;\n}\n\nQR8bitByte.prototype = {\n\n\tgetLength : function() {\n\t\treturn this.data.length;\n\t},\n\t\n\twrite : function(buffer) {\n\t\tfor (var i = 0; i < this.data.length; i++) {\n\t\t\t// not JIS ...\n\t\t\tbuffer.put(this.data.charCodeAt(i), 8);\n\t\t}\n\t}\n};\n\nmodule.exports = QR8bitByte;\n\n},\n"./QRUtil": function(module, exports, require) {\nvar QRMode = require(\'./QRMode\');\nvar QRPolynomial = require(\'./QRPolynomial\');\nvar QRMath = require(\'./QRMath\');\nvar QRMaskPattern = require(\'./QRMaskPattern\');\n\nvar QRUtil = {\n\n    PATTERN_POSITION_TABLE : [\n        [],\n        [6, 18],\n        [6, 22],\n        [6, 26],\n        [6, 30],\n        [6, 34],\n        [6, 22, 38],\n        [6, 24, 42],\n        [6, 26, 46],\n        [6, 28, 50],\n        [6, 30, 54],        \n        [6, 32, 58],\n        [6, 34, 62],\n        [6, 26, 46, 66],\n        [6, 26, 48, 70],\n        [6, 26, 50, 74],\n        [6, 30, 54, 78],\n        [6, 30, 56, 82],\n        [6, 30, 58, 86],\n        [6, 34, 62, 90],\n        [6, 28, 50, 72, 94],\n        [6, 26, 50, 74, 98],\n        [6, 30, 54, 78, 102],\n        [6, 28, 54, 80, 106],\n        [6, 32, 58, 84, 110],\n        [6, 30, 58, 86, 114],\n        [6, 34, 62, 90, 118],\n        [6, 26, 50, 74, 98, 122],\n        [6, 30, 54, 78, 102, 126],\n        [6, 26, 52, 78, 104, 130],\n        [6, 30, 56, 82, 108, 134],\n        [6, 34, 60, 86, 112, 138],\n        [6, 30, 58, 86, 114, 142],\n        [6, 34, 62, 90, 118, 146],\n        [6, 30, 54, 78, 102, 126, 150],\n        [6, 24, 50, 76, 102, 128, 154],\n        [6, 28, 54, 80, 106, 132, 158],\n        [6, 32, 58, 84, 110, 136, 162],\n        [6, 26, 54, 82, 110, 138, 166],\n        [6, 30, 58, 86, 114, 142, 170]\n    ],\n\n    G15 : (1 << 10) | (1 << 8) | (1 << 5) | (1 << 4) | (1 << 2) | (1 << 1) | (1 << 0),\n    G18 : (1 << 12) | (1 << 11) | (1 << 10) | (1 << 9) | (1 << 8) | (1 << 5) | (1 << 2) | (1 << 0),\n    G15_MASK : (1 << 14) | (1 << 12) | (1 << 10)    | (1 << 4) | (1 << 1),\n\n    getBCHTypeInfo : function(data) {\n        var d = data << 10;\n        while (QRUtil.getBCHDigit(d) - QRUtil.getBCHDigit(QRUtil.G15) >= 0) {\n            d ^= (QRUtil.G15 << (QRUtil.getBCHDigit(d) - QRUtil.getBCHDigit(QRUtil.G15) ) );    \n        }\n        return ( (data << 10) | d) ^ QRUtil.G15_MASK;\n    },\n\n    getBCHTypeNumber : function(data) {\n        var d = data << 12;\n        while (QRUtil.getBCHDigit(d) - QRUtil.getBCHDigit(QRUtil.G18) >= 0) {\n            d ^= (QRUtil.G18 << (QRUtil.getBCHDigit(d) - QRUtil.getBCHDigit(QRUtil.G18) ) );    \n        }\n        return (data << 12) | d;\n    },\n\n    getBCHDigit : function(data) {\n\n        var digit = 0;\n\n        while (data !== 0) {\n            digit++;\n            data >>>= 1;\n        }\n\n        return digit;\n    },\n\n    getPatternPosition : function(typeNumber) {\n        return QRUtil.PATTERN_POSITION_TABLE[typeNumber - 1];\n    },\n\n    getMask : function(maskPattern, i, j) {\n        \n        switch (maskPattern) {\n            \n        case QRMaskPattern.PATTERN000 : return (i + j) % 2 === 0;\n        case QRMaskPattern.PATTERN001 : return i % 2 === 0;\n        case QRMaskPattern.PATTERN010 : return j % 3 === 0;\n        case QRMaskPattern.PATTERN011 : return (i + j) % 3 === 0;\n        case QRMaskPattern.PATTERN100 : return (Math.floor(i / 2) + Math.floor(j / 3) ) % 2 === 0;\n        case QRMaskPattern.PATTERN101 : return (i * j) % 2 + (i * j) % 3 === 0;\n        case QRMaskPattern.PATTERN110 : return ( (i * j) % 2 + (i * j) % 3) % 2 === 0;\n        case QRMaskPattern.PATTERN111 : return ( (i * j) % 3 + (i + j) % 2) % 2 === 0;\n\n        default :\n            throw new Error("bad maskPattern:" + maskPattern);\n        }\n    },\n\n    getErrorCorrectPolynomial : function(errorCorrectLength) {\n\n        var a = new QRPolynomial([1], 0);\n\n        for (var i = 0; i < errorCorrectLength; i++) {\n            a = a.multiply(new QRPolynomial([1, QRMath.gexp(i)], 0) );\n        }\n\n        return a;\n    },\n\n    getLengthInBits : function(mode, type) {\n\n        if (1 <= type && type < 10) {\n\n            // 1 - 9\n\n            switch(mode) {\n            case QRMode.MODE_NUMBER     : return 10;\n            case QRMode.MODE_ALPHA_NUM  : return 9;\n            case QRMode.MODE_8BIT_BYTE  : return 8;\n            case QRMode.MODE_KANJI      : return 8;\n            default :\n                throw new Error("mode:" + mode);\n            }\n\n        } else if (type < 27) {\n\n            // 10 - 26\n\n            switch(mode) {\n            case QRMode.MODE_NUMBER     : return 12;\n            case QRMode.MODE_ALPHA_NUM  : return 11;\n            case QRMode.MODE_8BIT_BYTE  : return 16;\n            case QRMode.MODE_KANJI      : return 10;\n            default :\n                throw new Error("mode:" + mode);\n            }\n\n        } else if (type < 41) {\n\n            // 27 - 40\n\n            switch(mode) {\n            case QRMode.MODE_NUMBER     : return 14;\n            case QRMode.MODE_ALPHA_NUM  : return 13;\n            case QRMode.MODE_8BIT_BYTE  : return 16;\n            case QRMode.MODE_KANJI      : return 12;\n            default :\n                throw new Error("mode:" + mode);\n            }\n\n        } else {\n            throw new Error("type:" + type);\n        }\n    },\n\n    getLostPoint : function(qrCode) {\n        \n        var moduleCount = qrCode.getModuleCount();\n        var lostPoint = 0;\n        var row = 0; \n        var col = 0;\n\n        \n        // LEVEL1\n        \n        for (row = 0; row < moduleCount; row++) {\n\n            for (col = 0; col < moduleCount; col++) {\n\n                var sameCount = 0;\n                var dark = qrCode.isDark(row, col);\n\n                for (var r = -1; r <= 1; r++) {\n\n                    if (row + r < 0 || moduleCount <= row + r) {\n                        continue;\n                    }\n\n                    for (var c = -1; c <= 1; c++) {\n\n                        if (col + c < 0 || moduleCount <= col + c) {\n                            continue;\n                        }\n\n                        if (r === 0 && c === 0) {\n                            continue;\n                        }\n\n                        if (dark === qrCode.isDark(row + r, col + c) ) {\n                            sameCount++;\n                        }\n                    }\n                }\n\n                if (sameCount > 5) {\n                    lostPoint += (3 + sameCount - 5);\n                }\n            }\n        }\n\n        // LEVEL2\n\n        for (row = 0; row < moduleCount - 1; row++) {\n            for (col = 0; col < moduleCount - 1; col++) {\n                var count = 0;\n                if (qrCode.isDark(row,     col    ) ) count++;\n                if (qrCode.isDark(row + 1, col    ) ) count++;\n                if (qrCode.isDark(row,     col + 1) ) count++;\n                if (qrCode.isDark(row + 1, col + 1) ) count++;\n                if (count === 0 || count === 4) {\n                    lostPoint += 3;\n                }\n            }\n        }\n\n        // LEVEL3\n\n        for (row = 0; row < moduleCount; row++) {\n            for (col = 0; col < moduleCount - 6; col++) {\n                if (qrCode.isDark(row, col) && \n                        !qrCode.isDark(row, col + 1) && \n                         qrCode.isDark(row, col + 2) && \n                         qrCode.isDark(row, col + 3) && \n                         qrCode.isDark(row, col + 4) && \n                        !qrCode.isDark(row, col + 5) && \n                         qrCode.isDark(row, col + 6) ) {\n                    lostPoint += 40;\n                }\n            }\n        }\n\n        for (col = 0; col < moduleCount; col++) {\n            for (row = 0; row < moduleCount - 6; row++) {\n                if (qrCode.isDark(row, col) &&\n                        !qrCode.isDark(row + 1, col) &&\n                         qrCode.isDark(row + 2, col) &&\n                         qrCode.isDark(row + 3, col) &&\n                         qrCode.isDark(row + 4, col) &&\n                        !qrCode.isDark(row + 5, col) &&\n                         qrCode.isDark(row + 6, col) ) {\n                    lostPoint += 40;\n                }\n            }\n        }\n\n        // LEVEL4\n        \n        var darkCount = 0;\n\n        for (col = 0; col < moduleCount; col++) {\n            for (row = 0; row < moduleCount; row++) {\n                if (qrCode.isDark(row, col) ) {\n                    darkCount++;\n                }\n            }\n        }\n        \n        var ratio = Math.abs(100 * darkCount / moduleCount / moduleCount - 50) / 5;\n        lostPoint += ratio * 10;\n\n        return lostPoint;       \n    }\n\n};\n\nmodule.exports = QRUtil;\n\n},\n"./index": function(module, exports, require) {\n//---------------------------------------------------------------------\n// QRCode for JavaScript\n//\n// Copyright (c) 2009 Kazuhiko Arase\n//\n// URL: http://www.d-project.com/\n//\n// Licensed under the MIT license:\n//   http://www.opensource.org/licenses/mit-license.php\n//\n// The word "QR Code" is registered trademark of \n// DENSO WAVE INCORPORATED\n//   http://www.denso-wave.com/qrcode/faqpatent-e.html\n//\n//---------------------------------------------------------------------\n// Modified to work in node for this project (and some refactoring)\n//---------------------------------------------------------------------\n\nvar QR8bitByte = require(\'./QR8bitByte\');\nvar QRUtil = require(\'./QRUtil\');\nvar QRPolynomial = require(\'./QRPolynomial\');\nvar QRRSBlock = require(\'./QRRSBlock\');\nvar QRBitBuffer = require(\'./QRBitBuffer\');\n\nfunction QRCode(typeNumber, errorCorrectLevel) {\n\tthis.typeNumber = typeNumber;\n\tthis.errorCorrectLevel = errorCorrectLevel;\n\tthis.modules = null;\n\tthis.moduleCount = 0;\n\tthis.dataCache = null;\n\tthis.dataList = [];\n}\n\nQRCode.prototype = {\n\t\n\taddData : function(data) {\n\t\tvar newData = new QR8bitByte(data);\n\t\tthis.dataList.push(newData);\n\t\tthis.dataCache = null;\n\t},\n\t\n\tisDark : function(row, col) {\n\t\tif (row < 0 || this.moduleCount <= row || col < 0 || this.moduleCount <= col) {\n\t\t\tthrow new Error(row + "," + col);\n\t\t}\n\t\treturn this.modules[row][col];\n\t},\n\n\tgetModuleCount : function() {\n\t\treturn this.moduleCount;\n\t},\n\t\n\tmake : function() {\n\t\t// Calculate automatically typeNumber if provided is < 1\n\t\tif (this.typeNumber < 1 ){\n\t\t\tvar typeNumber = 1;\n\t\t\tfor (typeNumber = 1; typeNumber < 40; typeNumber++) {\n\t\t\t\tvar rsBlocks = QRRSBlock.getRSBlocks(typeNumber, this.errorCorrectLevel);\n\n\t\t\t\tvar buffer = new QRBitBuffer();\n\t\t\t\tvar totalDataCount = 0;\n\t\t\t\tfor (var i = 0; i < rsBlocks.length; i++) {\n\t\t\t\t\ttotalDataCount += rsBlocks[i].dataCount;\n\t\t\t\t}\n\n\t\t\t\tfor (var x = 0; x < this.dataList.length; x++) {\n\t\t\t\t\tvar data = this.dataList[x];\n\t\t\t\t\tbuffer.put(data.mode, 4);\n\t\t\t\t\tbuffer.put(data.getLength(), QRUtil.getLengthInBits(data.mode, typeNumber) );\n\t\t\t\t\tdata.write(buffer);\n\t\t\t\t}\n\t\t\t\tif (buffer.getLengthInBits() <= totalDataCount * 8)\n\t\t\t\t\tbreak;\n\t\t\t}\n\t\t\tthis.typeNumber = typeNumber;\n\t\t}\n\t\tthis.makeImpl(false, this.getBestMaskPattern() );\n\t},\n\t\n\tmakeImpl : function(test, maskPattern) {\n\t\t\n\t\tthis.moduleCount = this.typeNumber * 4 + 17;\n\t\tthis.modules = new Array(this.moduleCount);\n\t\t\n\t\tfor (var row = 0; row < this.moduleCount; row++) {\n\t\t\t\n\t\t\tthis.modules[row] = new Array(this.moduleCount);\n\t\t\t\n\t\t\tfor (var col = 0; col < this.moduleCount; col++) {\n\t\t\t\tthis.modules[row][col] = null;//(col + row) % 3;\n\t\t\t}\n\t\t}\n\t\n\t\tthis.setupPositionProbePattern(0, 0);\n\t\tthis.setupPositionProbePattern(this.moduleCount - 7, 0);\n\t\tthis.setupPositionProbePattern(0, this.moduleCount - 7);\n\t\tthis.setupPositionAdjustPattern();\n\t\tthis.setupTimingPattern();\n\t\tthis.setupTypeInfo(test, maskPattern);\n\t\t\n\t\tif (this.typeNumber >= 7) {\n\t\t\tthis.setupTypeNumber(test);\n\t\t}\n\t\n\t\tif (this.dataCache === null) {\n\t\t\tthis.dataCache = QRCode.createData(this.typeNumber, this.errorCorrectLevel, this.dataList);\n\t\t}\n\t\n\t\tthis.mapData(this.dataCache, maskPattern);\n\t},\n\n\tsetupPositionProbePattern : function(row, col)  {\n\t\t\n\t\tfor (var r = -1; r <= 7; r++) {\n\t\t\t\n\t\t\tif (row + r <= -1 || this.moduleCount <= row + r) continue;\n\t\t\t\n\t\t\tfor (var c = -1; c <= 7; c++) {\n\t\t\t\t\n\t\t\t\tif (col + c <= -1 || this.moduleCount <= col + c) continue;\n\t\t\t\t\n\t\t\t\tif ( (0 <= r && r <= 6 && (c === 0 || c === 6) ) || \n                     (0 <= c && c <= 6 && (r === 0 || r === 6) ) || \n                     (2 <= r && r <= 4 && 2 <= c && c <= 4) ) {\n\t\t\t\t\tthis.modules[row + r][col + c] = true;\n\t\t\t\t} else {\n\t\t\t\t\tthis.modules[row + r][col + c] = false;\n\t\t\t\t}\n\t\t\t}\t\t\n\t\t}\t\t\n\t},\n\t\n\tgetBestMaskPattern : function() {\n\t\n\t\tvar minLostPoint = 0;\n\t\tvar pattern = 0;\n\t\n\t\tfor (var i = 0; i < 8; i++) {\n\t\t\t\n\t\t\tthis.makeImpl(true, i);\n\t\n\t\t\tvar lostPoint = QRUtil.getLostPoint(this);\n\t\n\t\t\tif (i === 0 || minLostPoint >  lostPoint) {\n\t\t\t\tminLostPoint = lostPoint;\n\t\t\t\tpattern = i;\n\t\t\t}\n\t\t}\n\t\n\t\treturn pattern;\n\t},\n\t\n\tcreateMovieClip : function(target_mc, instance_name, depth) {\n\t\n\t\tvar qr_mc = target_mc.createEmptyMovieClip(instance_name, depth);\n\t\tvar cs = 1;\n\t\n\t\tthis.make();\n\n\t\tfor (var row = 0; row < this.modules.length; row++) {\n\t\t\t\n\t\t\tvar y = row * cs;\n\t\t\t\n\t\t\tfor (var col = 0; col < this.modules[row].length; col++) {\n\t\n\t\t\t\tvar x = col * cs;\n\t\t\t\tvar dark = this.modules[row][col];\n\t\t\t\n\t\t\t\tif (dark) {\n\t\t\t\t\tqr_mc.beginFill(0, 100);\n\t\t\t\t\tqr_mc.moveTo(x, y);\n\t\t\t\t\tqr_mc.lineTo(x + cs, y);\n\t\t\t\t\tqr_mc.lineTo(x + cs, y + cs);\n\t\t\t\t\tqr_mc.lineTo(x, y + cs);\n\t\t\t\t\tqr_mc.endFill();\n\t\t\t\t}\n\t\t\t}\n\t\t}\n\t\t\n\t\treturn qr_mc;\n\t},\n\n\tsetupTimingPattern : function() {\n\t\t\n\t\tfor (var r = 8; r < this.moduleCount - 8; r++) {\n\t\t\tif (this.modules[r][6] !== null) {\n\t\t\t\tcontinue;\n\t\t\t}\n\t\t\tthis.modules[r][6] = (r % 2 === 0);\n\t\t}\n\t\n\t\tfor (var c = 8; c < this.moduleCount - 8; c++) {\n\t\t\tif (this.modules[6][c] !== null) {\n\t\t\t\tcontinue;\n\t\t\t}\n\t\t\tthis.modules[6][c] = (c % 2 === 0);\n\t\t}\n\t},\n\t\n\tsetupPositionAdjustPattern : function() {\n\t\n\t\tvar pos = QRUtil.getPatternPosition(this.typeNumber);\n\t\t\n\t\tfor (var i = 0; i < pos.length; i++) {\n\t\t\n\t\t\tfor (var j = 0; j < pos.length; j++) {\n\t\t\t\n\t\t\t\tvar row = pos[i];\n\t\t\t\tvar col = pos[j];\n\t\t\t\t\n\t\t\t\tif (this.modules[row][col] !== null) {\n\t\t\t\t\tcontinue;\n\t\t\t\t}\n\t\t\t\t\n\t\t\t\tfor (var r = -2; r <= 2; r++) {\n\t\t\t\t\n\t\t\t\t\tfor (var c = -2; c <= 2; c++) {\n\t\t\t\t\t\n\t\t\t\t\t\tif (Math.abs(r) === 2 || \n                            Math.abs(c) === 2 ||\n                            (r === 0 && c === 0) ) {\n\t\t\t\t\t\t\tthis.modules[row + r][col + c] = true;\n\t\t\t\t\t\t} else {\n\t\t\t\t\t\t\tthis.modules[row + r][col + c] = false;\n\t\t\t\t\t\t}\n\t\t\t\t\t}\n\t\t\t\t}\n\t\t\t}\n\t\t}\n\t},\n\t\n\tsetupTypeNumber : function(test) {\n\t\n\t\tvar bits = QRUtil.getBCHTypeNumber(this.typeNumber);\n        var mod;\n\t\n\t\tfor (var i = 0; i < 18; i++) {\n\t\t\tmod = (!test && ( (bits >> i) & 1) === 1);\n\t\t\tthis.modules[Math.floor(i / 3)][i % 3 + this.moduleCount - 8 - 3] = mod;\n\t\t}\n\t\n\t\tfor (var x = 0; x < 18; x++) {\n\t\t\tmod = (!test && ( (bits >> x) & 1) === 1);\n\t\t\tthis.modules[x % 3 + this.moduleCount - 8 - 3][Math.floor(x / 3)] = mod;\n\t\t}\n\t},\n\t\n\tsetupTypeInfo : function(test, maskPattern) {\n\t\n\t\tvar data = (this.errorCorrectLevel << 3) | maskPattern;\n\t\tvar bits = QRUtil.getBCHTypeInfo(data);\n        var mod;\n\t\n\t\t// vertical\t\t\n\t\tfor (var v = 0; v < 15; v++) {\n\t\n\t\t\tmod = (!test && ( (bits >> v) & 1) === 1);\n\t\n\t\t\tif (v < 6) {\n\t\t\t\tthis.modules[v][8] = mod;\n\t\t\t} else if (v < 8) {\n\t\t\t\tthis.modules[v + 1][8] = mod;\n\t\t\t} else {\n\t\t\t\tthis.modules[this.moduleCount - 15 + v][8] = mod;\n\t\t\t}\n\t\t}\n\t\n\t\t// horizontal\n\t\tfor (var h = 0; h < 15; h++) {\n\t\n\t\t\tmod = (!test && ( (bits >> h) & 1) === 1);\n\t\t\t\n\t\t\tif (h < 8) {\n\t\t\t\tthis.modules[8][this.moduleCount - h - 1] = mod;\n\t\t\t} else if (h < 9) {\n\t\t\t\tthis.modules[8][15 - h - 1 + 1] = mod;\n\t\t\t} else {\n\t\t\t\tthis.modules[8][15 - h - 1] = mod;\n\t\t\t}\n\t\t}\n\t\n\t\t// fixed module\n\t\tthis.modules[this.moduleCount - 8][8] = (!test);\n\t\n\t},\n\t\n\tmapData : function(data, maskPattern) {\n\t\t\n\t\tvar inc = -1;\n\t\tvar row = this.moduleCount - 1;\n\t\tvar bitIndex = 7;\n\t\tvar byteIndex = 0;\n\t\t\n\t\tfor (var col = this.moduleCount - 1; col > 0; col -= 2) {\n\t\n\t\t\tif (col === 6) col--;\n\t\n\t\t\twhile (true) {\n\t\n\t\t\t\tfor (var c = 0; c < 2; c++) {\n\t\t\t\t\t\n\t\t\t\t\tif (this.modules[row][col - c] === null) {\n\t\t\t\t\t\t\n\t\t\t\t\t\tvar dark = false;\n\t\n\t\t\t\t\t\tif (byteIndex < data.length) {\n\t\t\t\t\t\t\tdark = ( ( (data[byteIndex] >>> bitIndex) & 1) === 1);\n\t\t\t\t\t\t}\n\t\n\t\t\t\t\t\tvar mask = QRUtil.getMask(maskPattern, row, col - c);\n\t\n\t\t\t\t\t\tif (mask) {\n\t\t\t\t\t\t\tdark = !dark;\n\t\t\t\t\t\t}\n\t\t\t\t\t\t\n\t\t\t\t\t\tthis.modules[row][col - c] = dark;\n\t\t\t\t\t\tbitIndex--;\n\t\n\t\t\t\t\t\tif (bitIndex === -1) {\n\t\t\t\t\t\t\tbyteIndex++;\n\t\t\t\t\t\t\tbitIndex = 7;\n\t\t\t\t\t\t}\n\t\t\t\t\t}\n\t\t\t\t}\n\t\t\t\t\t\t\t\t\n\t\t\t\trow += inc;\n\t\n\t\t\t\tif (row < 0 || this.moduleCount <= row) {\n\t\t\t\t\trow -= inc;\n\t\t\t\t\tinc = -inc;\n\t\t\t\t\tbreak;\n\t\t\t\t}\n\t\t\t}\n\t\t}\n\t\t\n\t}\n\n};\n\nQRCode.PAD0 = 0xEC;\nQRCode.PAD1 = 0x11;\n\nQRCode.createData = function(typeNumber, errorCorrectLevel, dataList) {\n\t\n\tvar rsBlocks = QRRSBlock.getRSBlocks(typeNumber, errorCorrectLevel);\n\t\n\tvar buffer = new QRBitBuffer();\n\t\n\tfor (var i = 0; i < dataList.length; i++) {\n\t\tvar data = dataList[i];\n\t\tbuffer.put(data.mode, 4);\n\t\tbuffer.put(data.getLength(), QRUtil.getLengthInBits(data.mode, typeNumber) );\n\t\tdata.write(buffer);\n\t}\n\n\t// calc num max data.\n\tvar totalDataCount = 0;\n\tfor (var x = 0; x < rsBlocks.length; x++) {\n\t\ttotalDataCount += rsBlocks[x].dataCount;\n\t}\n\n\tif (buffer.getLengthInBits() > totalDataCount * 8) {\n\t\tthrow new Error("code length overflow. (" + \n            buffer.getLengthInBits() + \n            ">" +  \n            totalDataCount * 8 + \n            ")");\n\t}\n\n\t// end code\n\tif (buffer.getLengthInBits() + 4 <= totalDataCount * 8) {\n\t\tbuffer.put(0, 4);\n\t}\n\n\t// padding\n\twhile (buffer.getLengthInBits() % 8 !== 0) {\n\t\tbuffer.putBit(false);\n\t}\n\n\t// padding\n\twhile (true) {\n\t\t\n\t\tif (buffer.getLengthInBits() >= totalDataCount * 8) {\n\t\t\tbreak;\n\t\t}\n\t\tbuffer.put(QRCode.PAD0, 8);\n\t\t\n\t\tif (buffer.getLengthInBits() >= totalDataCount * 8) {\n\t\t\tbreak;\n\t\t}\n\t\tbuffer.put(QRCode.PAD1, 8);\n\t}\n\n\treturn QRCode.createBytes(buffer, rsBlocks);\n};\n\nQRCode.createBytes = function(buffer, rsBlocks) {\n\n\tvar offset = 0;\n\t\n\tvar maxDcCount = 0;\n\tvar maxEcCount = 0;\n\t\n\tvar dcdata = new Array(rsBlocks.length);\n\tvar ecdata = new Array(rsBlocks.length);\n\t\n\tfor (var r = 0; r < rsBlocks.length; r++) {\n\n\t\tvar dcCount = rsBlocks[r].dataCount;\n\t\tvar ecCount = rsBlocks[r].totalCount - dcCount;\n\n\t\tmaxDcCount = Math.max(maxDcCount, dcCount);\n\t\tmaxEcCount = Math.max(maxEcCount, ecCount);\n\t\t\n\t\tdcdata[r] = new Array(dcCount);\n\t\t\n\t\tfor (var i = 0; i < dcdata[r].length; i++) {\n\t\t\tdcdata[r][i] = 0xff & buffer.buffer[i + offset];\n\t\t}\n\t\toffset += dcCount;\n\t\t\n\t\tvar rsPoly = QRUtil.getErrorCorrectPolynomial(ecCount);\n\t\tvar rawPoly = new QRPolynomial(dcdata[r], rsPoly.getLength() - 1);\n\n\t\tvar modPoly = rawPoly.mod(rsPoly);\n\t\tecdata[r] = new Array(rsPoly.getLength() - 1);\n\t\tfor (var x = 0; x < ecdata[r].length; x++) {\n            var modIndex = x + modPoly.getLength() - ecdata[r].length;\n\t\t\tecdata[r][x] = (modIndex >= 0)? modPoly.get(modIndex) : 0;\n\t\t}\n\n\t}\n\t\n\tvar totalCodeCount = 0;\n\tfor (var y = 0; y < rsBlocks.length; y++) {\n\t\ttotalCodeCount += rsBlocks[y].totalCount;\n\t}\n\n\tvar data = new Array(totalCodeCount);\n\tvar index = 0;\n\n\tfor (var z = 0; z < maxDcCount; z++) {\n\t\tfor (var s = 0; s < rsBlocks.length; s++) {\n\t\t\tif (z < dcdata[s].length) {\n\t\t\t\tdata[index++] = dcdata[s][z];\n\t\t\t}\n\t\t}\n\t}\n\n\tfor (var xx = 0; xx < maxEcCount; xx++) {\n\t\tfor (var t = 0; t < rsBlocks.length; t++) {\n\t\t\tif (xx < ecdata[t].length) {\n\t\t\t\tdata[index++] = ecdata[t][xx];\n\t\t\t}\n\t\t}\n\t}\n\n\treturn data;\n\n};\n\nmodule.exports = QRCode;\n\n}\n};\nconst cache = Object.create(null);\nfunction require(name) {\n  if (!Object.prototype.hasOwnProperty.call(modules,name)) throw new Error(\'Módulo QR desconhecido: \'+name);\n  if (!cache[name]) {\n    const module={exports:{}};\n    cache[name]=module;\n    modules[name](module,module.exports,require);\n  }\n  return cache[name].exports;\n}\nglobal.QRCodeLocal = require(\'./index\');\n})(window);\n'
QR_PANEL_CSS = '\n.share-tools{display:flex;flex-wrap:wrap;align-items:end;gap:12px}\n.share-tools label{flex:1;min-width:260px}\n.share-tools input{width:100%;margin-top:7px}\n.share-tools button{min-height:44px}\n.qr-result{margin-top:18px;display:flex;flex-wrap:wrap;gap:18px;align-items:center}\n.qr-result[hidden]{display:none}\n.qr-result canvas{max-width:100%;height:auto;border:1px solid #e1e8ef;border-radius:9px}\n.qr-note{margin-top:12px}\n.qr-status{margin-top:8px;font-size:12px;color:#526B81}\n'
QR_PANEL_HTML = '\n<div class="panel" id="compartilharRelatorio">\n  <div class="panel-head"><h2>Compartilhar relatório por QR Code</h2></div>\n  <div class="share-tools">\n    <label>Endereço público do HTML\n      <input id="urlRelatorioQR" type="url" inputmode="url" autocomplete="url" spellcheck="false" value="https://ronaldo-consuli.github.io/Controle-de-estoque-reagentes_Corta-fogo-1/monitoramento_reagentes.html" placeholder="https://ronaldo-consuli.github.io/Controle-de-estoque-reagentes_Corta-fogo-1/monitoramento_reagentes.html" aria-describedby="orientacaoQR">\n    </label>\n    <button id="gerarQR" type="button">▦ Gerar QR Code</button>\n  </div>\n  <div id="orientacaoQR" class="note qr-note">Endereço do GitHub Pages preenchido automaticamente. Para funcionar no celular, publique o arquivo monitoramento_reagentes.html na raiz do repositório e ative Pages (main / root). O QR Code não publica arquivos locais.</div>\n  <div class="qr-status" id="statusQR" role="status" aria-live="polite"></div>\n  <div class="qr-result" id="resultadoQR" hidden>\n    <canvas id="canvasQR" role="img" aria-label="QR Code para abrir este relatório"></canvas>\n    <div><button id="baixarQR" type="button">↓ Baixar QR Code (PNG)</button></div>\n  </div>\n</div>\n'
QR_PANEL_SCRIPT = "\n// Esta função é executada somente no HTML exportado, não dentro do iframe Streamlit.\n(function prepararQRCode(){\n  const campo=$('urlRelatorioQR'), status=$('statusQR'), resultado=$('resultadoQR');\n  const canvas=$('canvasQR');\n  if (window.location.protocol==='https:') {\n    // Quando publicado, usa a URL real desta página, mesmo que mude de lugar.\n    campo.value=window.location.href.split('#')[0];\n  } else {\n    atualizarStatus('Endereço previsto preenchido. Publique o HTML no GitHub Pages antes de compartilhar o QR Code.');\n  }\n  function atualizarStatus(texto, erro=false) {\n    status.textContent=texto;\n    status.style.color=erro?'#A13A2F':'#526B81';\n  }\n  $('gerarQR').addEventListener('click',()=>{\n    resultado.hidden=true;\n    const valor=campo.value.trim();\n    let url;\n    try {url=new URL(valor);} catch (_) {\n      atualizarStatus('Informe o endereço HTTPS do HTML publicado.',true);return;\n    }\n    const hostname=url.hostname.toLowerCase();\n    if(url.protocol!=='https:'||!hostname||hostname==='localhost'||hostname==='127.0.0.1'||hostname==='::1'||url.username||url.password){\n      atualizarStatus('Use um endereço público HTTPS, que possa ser aberto pelo celular.',true);return;\n    }\n    if(['github.com','www.github.com','raw.githubusercontent.com','gist.github.com'].includes(hostname)){\n      atualizarStatus('Informe o endereço do site publicado (ex.: usuario.github.io), e não o link do arquivo no GitHub.',true);return;\n    }\n    if(url.href.length>1000){\n      atualizarStatus('O endereço é muito longo para gerar um QR Code legível.',true);return;\n    }\n    try {\n      // Esta biblioteca está incluída no HTML, sem chamada a sites externos.\n      const qr=new window.QRCodeLocal(0, 0); // nível de correção M.\n      qr.addData(url.href);\n      qr.make();\n      const escala=8, margem=4, n=qr.getModuleCount();\n      canvas.width=(n+2*margem)*escala;\n      canvas.height=(n+2*margem)*escala;\n      const ctx=canvas.getContext('2d');\n      ctx.fillStyle='#ffffff';ctx.fillRect(0,0,canvas.width,canvas.height);\n      ctx.fillStyle='#143257';\n      for(let linha=0;linha<n;linha++)for(let coluna=0;coluna<n;coluna++){\n        if(qr.isDark(linha,coluna))ctx.fillRect((coluna+margem)*escala,(linha+margem)*escala,escala,escala);\n      }\n      resultado.hidden=false;\n      atualizarStatus('QR Code pronto. Confira se o link publicado abre no celular.');\n    } catch(e) {\n      atualizarStatus('Não foi possível gerar o QR Code para esse endereço.',true);\n    }\n  });\n  $('baixarQR').addEventListener('click',()=>{\n    if(resultado.hidden)return;\n    canvas.toBlob(blob=>{\n      if(!blob){atualizarStatus('Não foi possível salvar a imagem do QR Code.',true);return;}\n      const endereco=URL.createObjectURL(blob), a=document.createElement('a');\n      a.href=endereco;a.download='qr_estoque_mensal.png';\n      document.body.appendChild(a);a.click();a.remove();\n      setTimeout(()=>URL.revokeObjectURL(endereco),1200);\n    },'image/png');\n  });\n  campo.addEventListener('input',()=>{resultado.hidden=true;atualizarStatus('');});\n})();\n"


def painel_html(records, incluir_qr=True):
    """O HTML exportado permite gerar/baixar QR; a prévia Streamlit não exibe esse bloco."""
    pagina = _painel_html_base(records)
    if not incluir_qr:
        return pagina
    pagina = pagina.replace("</style>", QR_PANEL_CSS + "</style>", 1)
    marcador = '</div>\n<script>const dados='
    if marcador not in pagina:
        raise RuntimeError('Falha ao localizar o fechamento do painel para inserir o QR Code.')
    pagina = pagina.replace(marcador,
        QR_PANEL_HTML + '</div>\n<script>' + QR_LIB_JS + '</script>\n<script>const dados=', 1)
    marcador_js = 'render();\n</script></body></html>'
    if marcador_js not in pagina:
        raise RuntimeError('Falha ao localizar a inicialização do painel.')
    pagina = pagina.replace(marcador_js,'render();\n' + QR_PANEL_SCRIPT + '\n</script></body></html>', 1)
    return pagina


def csv_relatorio(records):
    out = io.StringIO()
    writer = csv.writer(out, delimiter=";")
    writer.writerow(["Mês", "Local", "Aba", "Produto", "Quantidade", "Unidade", "Polícia Federal", "Polícia Civil", "Exército", "Correspondência na referência", "Cabeçalho original", "Correção de unidade sem alterar valor"])
    for r in records:
        cells = [r["mes"], r["local"], r["aba"], r["produto"], "" if r["quantidade"] is None else str(r["quantidade"]).replace(".", ","), r["unidade"], *["Sim" if r[k] else "—" for k in ("pf", "pc", "eb")], r["referencia"]]
        cells.extend([r.get("produto_original", r["produto"]), "Sim" if r.get("unidade_corrigida") else "—"])
        writer.writerow(["'" + c if str(c).lstrip().startswith(("=", "+", "-", "@")) else c for c in cells])
    return out.getvalue().encode("utf-8-sig")




def main():
    st.set_page_config(page_title="Estoque de reagentes", page_icon="🧪", layout="wide")
    st.markdown("<style>.stApp{background:#F1F5F8}[data-testid='stMainBlockContainer']{padding-top:1.5rem;padding-bottom:1rem;max-width:1600px}[data-testid='stSidebar']{background:#fff}footer{visibility:hidden}</style>", unsafe_allow_html=True)
    with st.sidebar:
        files = st.file_uploader("Importar arquivos Excel", type=["xlsx"], accept_multiple_files=True)
    year, unit = date.today().year, "g"
    liquid_codes = LIQUIDOS_ML
    if not files:
        st.title("Monitoramento mensal de reagentes")
        st.info("Importe a planilha Excel na lateral. O formato esperado é o exportado pelo Forms: Mês e uma coluna RE por reagente, com cabeçalhos na primeira linha.")
        return
    snapshots, notices, hashes = [], [], set()
    for index, upload in enumerate(files):
        blob = upload.getvalue()
        digest = hashlib.sha256(blob).hexdigest()
        if digest in hashes:
            notices.append(f"{upload.name}: arquivo idêntico já importado; duplicata ignorada.")
            continue
        hashes.add(digest)
        default = re.sub(r"\([^()]*\)(?:\([^()]*\))?\s*$", "", upload.name.rsplit(".", 1)[0]).strip()
        # O local vem da coluna Local/Armário quando presente; caso contrário,
        # é identificado pelo nome do arquivo, sem controles extras na lateral.
        local = default or "Laboratório"
        try:
            items, issues = ler_excel(blob, upload.name, local, int(year), unit, liquid_codes)
            snapshots.extend(items)
            notices.extend(issues)
        except Exception as exc:
            st.error(f"Não foi possível importar {upload.name}: {exc}")
            return
    records, issues = consolidar(snapshots)
    notices.extend(issues)
    # Observações rotineiras não ocupam o painel. Só sinalizar perdas de dados.
    important = [n for n in notices if any(t in n for t in ("inválid", "ignorada", "nenhuma resposta"))]
    if important:
        st.warning("Alguns dados não puderam ser importados: " + " | ".join(dict.fromkeys(important)))
    if not records:
        st.warning("Nenhum estoque válido encontrado. Confira os cabeçalhos e o mês.")
        return
    # O HTML contém os dados importados e funciona localmente, sem Streamlit.
    html_painel = painel_html(records, incluir_qr=True)
    components.html(painel_html(records, incluir_qr=False), height=1250, scrolling=True)

    # Download ao final da página, depois do painel completo.
    st.download_button(
        "📥 Baixar painel em HTML",
        data=html_painel.encode("utf-8"),
        file_name="monitoramento_reagentes.html",
        mime="text/html",
        help="Baixa uma versão interativa do painel, com os dados atualmente importados.",
    )



if __name__ == "__main__":
    main()
