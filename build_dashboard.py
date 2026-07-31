"""
Gera o dashboard HTML do portfólio de produtos digitais SESI a partir da
planilha Planilha_Portfolio_Produtos_v1.xlsx, e publica no GitHub Pages
(branch main) se o repositório git já estiver configurado nesta pasta.

Uso:
    python3 build_dashboard.py
"""
import base64
import os
import subprocess
from datetime import date, datetime, timedelta
from hashlib import pbkdf2_hmac
from html import escape
from pathlib import Path

from openpyxl import load_workbook

BASE_DIR = Path(__file__).resolve().parent
PORTFOLIO_FILE = BASE_DIR / "Planilha_Portfolio_Produtos_v1.xlsx"
OUTPUT_FILE = BASE_DIR / "index.html"
PASSWORD_FILE = BASE_DIR / "senha_dashboard.txt"  # fora do Git (.gitignore)
LOGO_FILE = BASE_DIR / "fenix_icone.svg"          # só a fênix (transparente)
PBKDF2_ITER = 220000


def logo_data_uri():
    """Fênix como data URI (embutida na página — self-contained)."""
    svg = LOGO_FILE.read_bytes()
    return "data:image/svg+xml;base64," + base64.b64encode(svg).decode("ascii")


# Troca de abas (passado como argumento — chaves não são interpretadas pelo .format)
TAB_SCRIPT = """
(function(){
  var tabs = document.querySelectorAll('.tab');
  var panels = document.querySelectorAll('.panel');
  for (var i = 0; i < tabs.length; i++) {
    tabs[i].addEventListener('click', function(){
      for (var j = 0; j < tabs.length; j++) { tabs[j].classList.remove('active'); }
      for (var k = 0; k < panels.length; k++) { panels[k].classList.add('hidden'); }
      this.classList.add('active');
      document.getElementById('panel-' + this.getAttribute('data-tab')).classList.remove('hidden');
      window.scrollTo(0, 0);
    });
  }
})();
"""

NAVY = "#0D2137"
ORANGE = "#F97316"

# Etapas do fluxo (usadas nos cards e KPIs). Ordem = progressão do pipeline.
# (chave, coluna na aba Portfólio, cor, rótulo singular, rótulo plural)
STAGES = [
    ("nao_iniciada",    "Não iniciada",    "#E2E8F0", "não iniciada",    "não iniciadas"),
    ("planejamento",    "Planejamento",    "#7DD3FC", "planejamento",    "planejamento"),
    ("refinamento",     "Refinamento",     "#38BDF8", "refinamento",     "refinamento"),
    ("desenvolvimento", "Desenvolvimento", "#3B82F6", "desenvolvimento", "desenvolvimento"),
    ("homologacao",     "Homologação",     "#6366F1", "homologação",     "homologação"),
    ("treinamento",     "Treinamento",     "#8B5CF6", "treinamento",     "treinamento"),
    ("bloqueada",       "Bloqueada",       "#EF4444", "bloqueada",       "bloqueadas"),
    ("pausada",         "Pausada",         "#F59E0B", "pausada",         "pausadas"),
    ("concluida",       "Concluída",       "#22C55E", "concluída",       "concluídas"),
]
EXECUCAO_KEYS = ["planejamento", "refinamento", "desenvolvimento", "homologacao", "treinamento"]
PARADA_KEYS = ["bloqueada", "pausada"]

# Fases do épico no roadmap (Gantt). Barra dividida por semanas de cada fase.
PHASE_COLORS = {
    "discovery":   ("#E7D8A6", "#7C6A34", "discovery"),
    "development": ("#B3CEF2", "#2C5590", "development"),
    "delivery":    ("#A8E3BB", "#2C7A48", "delivery"),
    "plain":       ("#CBD5E1", "#475569", ""),
}

SINAL_COLORS = {"Verde": "#22C55E", "Amarelo": "#F59E0B", "Vermelho": "#EF4444"}
IMPACT_COLORS = {"Alto": "#DC2626", "Médio": "#F59E0B", "Baixo": "#64748B"}
AVATAR_PALETTE = ["#F97316", "#3B82F6", "#22C55E", "#A855F7", "#EF4444",
                  "#0EA5E9", "#EC4899", "#14B8A6", "#EAB308", "#6366F1"]

GANTT_YEAR = 2026
GANTT_START = date(GANTT_YEAR, 1, 1)
GANTT_END = date(GANTT_YEAR, 12, 31)
GANTT_DAYS = (GANTT_END - GANTT_START).days + 1
MONTHS = ["JAN", "FEV", "MAR", "ABR", "MAI", "JUN", "JUL", "AGO", "SET", "OUT", "NOV", "DEZ"]
MONTH_DAYS = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]

SITUACAO_ATIVAS = {"Ativo", "Concluído"}  # escopo aprovado: só portfólio ativo
EPICOS_OCULTOS_ROADMAP = {"GMUD"}          # épicos que não aparecem no roadmap


# ---------------------------------------------------------------------------
# Helpers de leitura de planilha
# ---------------------------------------------------------------------------

def find_row_with(ws, col, value, start=1, end=250):
    for r in range(start, end + 1):
        cell_val = ws.cell(row=r, column=col).value
        if cell_val is not None and str(cell_val).strip() == value:
            return r
    raise ValueError(f"Não encontrei '{value}' na coluna {col} da aba '{ws.title}'")


def read_table(ws, header_row, key_col=1):
    """Lê tabela a partir de header_row, parando quando a coluna-chave fica vazia."""
    max_col = ws.max_column
    headers = {}
    for c in range(1, max_col + 1):
        val = ws.cell(row=header_row, column=c).value
        if val is not None and str(val).strip() != "":
            headers[c] = str(val).strip()
    rows = []
    r = header_row + 1
    while True:
        key_val = ws.cell(row=r, column=key_col).value
        if key_val is None or str(key_val).strip() == "":
            break
        rows.append({headers[c]: ws.cell(row=r, column=c).value for c in headers})
        r += 1
    return rows


def to_date(value):
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, (int, float)):
        return (datetime(1899, 12, 30) + timedelta(days=value)).date()
    return None


def num(value, default=0.0):
    try:
        return float(value)
    except (TypeError, ValueError):
        return default


def label_value(ws, label, value_col, label_col=1, end=80):
    r = find_row_with(ws, label_col, label, 1, end)
    return ws.cell(row=r, column=value_col).value


def fmt_pf(value):
    return f"{value:,.1f}".replace(",", "_").replace(".", ",").replace("_", ".")


def fmt_pf0(value):
    return f"{value:,.0f}".replace(",", ".")


def fmt_pct(value):
    return f"{value * 100:.0f}%"


def fmt_dm(d):
    return d.strftime("%d/%m") if d else "—"


# ---------------------------------------------------------------------------
# Extração
# ---------------------------------------------------------------------------

def extract_financeiro(ws):
    contratado = num(label_value(ws, "Total contratado (PF)", 2))
    vig = find_row_with(ws, 1, "Vigência")
    vig_ini = to_date(ws.cell(row=vig, column=2).value)
    vig_fim = to_date(ws.cell(row=vig, column=3).value)

    linhas = []
    for nome in ["Retroativos (2025)", "Novos desenv. (2026)", "TOTAL"]:
        r = find_row_with(ws, 1, nome)
        linhas.append((nome, {
            "previsto": num(ws.cell(row=r, column=2).value),
            "faturado": num(ws.cell(row=r, column=3).value),
            "pct": num(ws.cell(row=r, column=4).value),
        }))

    faturado_total = dict(linhas)["TOTAL"]["faturado"]

    # Janelas de faturamento (tabela abaixo de "JANELAS DE FATURAMENTO").
    jan_row = find_row_with(ws, 1, "JANELAS DE FATURAMENTO")
    janelas = []
    r = jan_row + 1
    while True:
        janela = to_date(ws.cell(row=r, column=1).value)
        if janela is None:
            break
        janelas.append({
            "janela": janela,
            "enviar_ate": to_date(ws.cell(row=r, column=2).value),
            "situacao": str(ws.cell(row=r, column=3).value or "").strip(),
            "iniciativas": ws.cell(row=r, column=4).value,
            "pf": ws.cell(row=r, column=5).value,
        })
        r += 1

    # Proposta a negociar para o próximo ciclo (bloco final da aba).
    prop_row = find_row_with(ws, 1, "PROPOSTA A NEGOCIAR PARA O PRÓXIMO CICLO")
    proposta, prop_nota = [], ""
    r = prop_row + 1
    while True:
        c1 = ws.cell(row=r, column=1).value
        if c1 is None or str(c1).strip() == "":
            break
        label = str(c1).strip()
        val = ws.cell(row=r, column=5).value
        if val is None or str(val).strip() == "":
            prop_nota = label  # linha de observação (sem valor)
        else:
            proposta.append({"label": label, "valor": num(val)})
        r += 1

    return {
        "contratado": contratado,
        "vig_ini": vig_ini,
        "vig_fim": vig_fim,
        "linhas": linhas,
        "faturado_total": faturado_total,
        "faturado_pct": num(label_value(ws, "Faturado sobre o contratado", 2)),
        "saldo": num(label_value(ws, "Saldo a faturar (PF)", 2)),
        "meses": int(num(label_value(ws, "Meses restantes", 2))),
        "media_mes": num(label_value(ws, "Média PF por mês necessário", 2)),
        "pf_a_faturar": num(label_value(ws, "PF a faturar", 2)),
        "item_antigo": int(num(label_value(ws, "Item mais antigo parado (dias)", 2))),
        "janelas": janelas,
        "total_faturavel": num(label_value(ws, "TOTAL faturável dentro do ciclo", 5)),
        "teto": num(label_value(ws, "Teto do contrato (já faturado + faturável)", 5)),
        "proposta": proposta,
        "prop_nota": prop_nota,
    }


def extract_portfolio(ws):
    header = find_row_with(ws, 1, "Produto")
    rows = read_table(ws, header, key_col=1)
    produtos = []
    for r in rows:
        nome = str(r.get("Produto") or "").strip()
        if nome in ("", "TOTAL"):
            continue
        total = int(num(r.get("Iniciativas")))
        if total == 0:
            continue  # produto sem iniciativas ativas (ex: Plat. de Colaboração)
        counts = {key: int(num(r.get(col))) for key, col, *_ in STAGES}
        produtos.append({
            "produto": nome,
            "fase": str(r.get("Fase") or "").strip(),
            "sinal": str(r.get("Sinal") or "").strip(),
            "total": total,
            "counts": counts,
            "pf_estimado": num(r.get("PF estimado")),
            "pf_faturado": num(r.get("PF faturado")),
            "evolucao": num(r.get("Evolução média")),
            "proxima": str(r.get("Próxima iniciativa") or "").strip(),
            "bloqueio_ativo": str(r.get("Bloqueio ativo?") or "").strip() == "Sim",
        })
    return produtos


def kpis_from_portfolio(produtos):
    def soma(keys):
        return sum(p["counts"][k] for p in produtos for k in keys)
    return {
        "total": sum(p["total"] for p in produtos),
        "nao_iniciadas": soma(["nao_iniciada"]),
        "em_execucao": soma(EXECUCAO_KEYS),
        "em_planejamento": soma(["planejamento", "refinamento"]),
        "em_desenvolvimento": soma(["desenvolvimento", "homologacao", "treinamento"]),
        "paradas": soma(PARADA_KEYS),
        "concluidas": soma(["concluida"]),
    }


def extract_roadmap(ws):
    header = find_row_with(ws, 1, "Situação")
    rows = read_table(ws, header, key_col=3)  # key = Código
    itens = []
    for r in rows:
        if str(r.get("Situação") or "").strip() not in SITUACAO_ATIVAS:
            continue
        if str(r.get("Épico") or "").strip().upper() in EPICOS_OCULTOS_ROADMAP:
            continue
        inicio = to_date(r.get("Início"))
        if inicio is None:
            continue
        fim = to_date(r.get("Fim (previsto)"))
        equipe = []
        for col in ["PO 1", "PO 2", "Designer", "DEV 1", "DEV 2", "DEV 3"]:
            n = str(r.get(col) or "").strip()
            if n and n not in ("N/A", "A definir"):
                equipe.append(n)
        itens.append({
            "produto": str(r.get("Produto") or "").strip(),
            "epico": str(r.get("Épico") or "").strip(),
            "inicio": inicio,
            "fim": fim,
            "discovery": num(r.get("Discovery (semanas)")),
            "development": num(r.get("Development (semanas)")),
            "delivery": num(r.get("Delivery  (semanas)") or r.get("Delivery (semanas)")),
            "pf": num(r.get("PF (estimado)")),
            "evolucao": num(r.get("Evolução")),
            "equipe": equipe,
        })
    itens.sort(key=lambda x: x["inicio"])
    return itens


def extract_bloqueios(ws):
    header = find_row_with(ws, 1, "Produto")
    rows = read_table(ws, header, key_col=1)
    resolvidos = [r for r in rows if str(r.get("Status do bloqueio") or "").strip() == "Resolvido"]
    resolvidos.sort(key=lambda r: to_date(r.get("Data de resolução")) or date.min, reverse=True)
    return resolvidos


# ---------------------------------------------------------------------------
# Renderização
# ---------------------------------------------------------------------------

def render_kpis(k):
    tiles = [
        ("Total de iniciativas", k["total"], "Portfólio ativo 2026", "#3B82F6"),
        ("Não iniciadas", k["nao_iniciadas"], "Aguardando priorização", "#94A3B8"),
        ("Em execução", k["em_execucao"],
         f"{k['em_planejamento']} planej./refin. · {k['em_desenvolvimento']} dev+", "#3B82F6"),
        ("Paradas", k["paradas"], "Bloqueadas e pausadas", "#F59E0B"),
        ("Concluídas", k["concluidas"], "Entregues no ano", "#22C55E"),
    ]
    html = ""
    for label, value, sub, accent in tiles:
        html += f"""
        <div class="kpi-card" style="--accent:{accent}">
          <div class="kpi-value">{value}</div>
          <div class="kpi-label">{escape(label)}</div>
          <div class="kpi-sub">{escape(sub)}</div>
        </div>"""
    return html


def render_cards(produtos):
    color = {key: c for key, _col, c, *_ in STAGES}
    sing = {key: s for key, _col, _c, s, _p in STAGES}
    plur = {key: p for key, _col, _c, _s, p in STAGES}
    html = ""
    for p in produtos:
        sinal = SINAL_COLORS.get(p["sinal"], "#94A3B8")
        segs = "".join(
            f'<span style="flex:{p["counts"][key]};background:{color[key]}"></span>'
            for key, *_ in STAGES if p["counts"][key] > 0
        )
        itens = [(key, p["counts"][key]) for key, *_ in STAGES if p["counts"][key] > 0]
        legenda = (f'<span><i style="background:#94A3B8"></i>'
                   f'{p["total"]} {"iniciativa" if p["total"] == 1 else "iniciativas"}</span>')
        for key, n in itens:
            rot = sing[key] if n == 1 else plur[key]
            legenda += f'<span><i style="background:{color[key]}"></i>{n} {escape(rot)}</span>'

        pf = ""
        if p["pf_estimado"] > 0:
            pf = f'<span>PF {fmt_pf0(p["pf_faturado"])}/{fmt_pf0(p["pf_estimado"])}</span>'
        proxima = (f'<div class="prod-next">Próxima: {escape(p["proxima"])}</div>'
                   if p["proxima"] else "")
        tag = '<div class="prod-tag">Bloqueio ativo</div>' if p["bloqueio_ativo"] else ""
        blocked = " blocked" if p["bloqueio_ativo"] else ""
        html += f"""
        <div class="prod-card{blocked}">
          <div class="prod-head">
            <div><h3>{escape(p["produto"])}</h3>
              <div class="prod-sub">{escape(p["fase"])} · {p["total"]} iniciativas</div></div>
            <span class="sinal-dot" style="background:{sinal}"></span>
          </div>
          <div class="seg-bar">{segs}</div>
          <div class="prod-legend">{legenda}</div>
          <div class="prod-foot"><span>Evolução {fmt_pct(p["evolucao"])}</span>{pf}</div>
          {proxima}{tag}
        </div>"""
    return html


def _avatar(name):
    parts = name.split()
    ini = (parts[0][0] + (parts[1][0] if len(parts) > 1 else "")).upper()
    cor = AVATAR_PALETTE[sum(ord(c) for c in name) % len(AVATAR_PALETTE)]
    return f'<span class="rm-av" style="background:{cor}" title="{escape(name)}">{escape(ini)}</span>'


def _phase_segments(item):
    segs = []
    cursor = item["inicio"]
    for weeks, key in [(item["discovery"], "discovery"),
                       (item["development"], "development"),
                       (item["delivery"], "delivery")]:
        if weeks and weeks > 0:
            end = cursor + timedelta(weeks=weeks)
            segs.append((cursor, end, key))
            cursor = end
    if not segs:
        end = item["fim"] or (item["inicio"] + timedelta(days=14))
        segs.append((item["inicio"], end, "plain"))
    return segs


def render_roadmap(itens):
    # Cabeçalho de meses (12 colunas proporcionais aos dias)
    months_html = ""
    for m, d in zip(MONTHS, MONTH_DAYS):
        months_html += f'<div class="rm-month" style="flex:{d}">{m}</div>'

    # Marcador "estamos aqui" no começo de agosto/2026
    marker = (date(2026, 8, 1) - GANTT_START).days / GANTT_DAYS * 100
    marker_head = f'<div class="rm-now" style="left:{marker:.2f}%"><span>estamos aqui</span></div>'
    overlay = (f'<div class="rm-now-overlay"><div class="rm-now-spacer"></div>'
               f'<div class="rm-now-track"><div class="rm-now-line" style="left:{marker:.2f}%"></div></div></div>')

    rows_html = ""
    for it in itens:
        avatars = "".join(_avatar(n) for n in it["equipe"][:6])
        bars = ""
        for start, end, key in _phase_segments(it):
            left = max((start - GANTT_START).days / GANTT_DAYS * 100, 0)
            right = min((end - GANTT_START).days / GANTT_DAYS * 100, 100)
            width = right - left
            if width <= 0:
                continue
            bg, fg, phase = PHASE_COLORS[key]
            # Barras só com cor — a fase é indicada pela legenda; o nome do épico
            # fica na caixa à esquerda.
            bars += (f'<div class="rm-bar" style="left:{left:.2f}%;width:{width:.2f}%;'
                     f'background:{bg}"></div>')
        fim_txt = it["fim"].strftime("%d/%m/%Y") if it["fim"] else "—"
        tip = (f'{it["produto"]} · {it["epico"]} | início {it["inicio"].strftime("%d/%m/%Y")}'
               f' → fim {fim_txt} | PF estimado {fmt_pf(it["pf"])}')
        rows_html += f"""
        <div class="rm-row" title="{escape(tip)}">
          <div class="rm-side">
            <span class="rm-prod">{escape(it["produto"])}</span>
            <span class="rm-avatars">{avatars}</span>
            <span class="rm-epic"><b>{escape(it["epico"])}</b><i>{fmt_pct(it["evolucao"])}</i></span>
          </div>
          <div class="rm-timeline">{bars}</div>
        </div>"""

    legend = "".join(
        f'<span class="legend-item"><span class="legend-dot" style="background:{PHASE_COLORS[k][0]}"></span>{PHASE_COLORS[k][2]}</span>'
        for k in ["discovery", "development", "delivery"]
    )
    return f"""
    <div class="rm-legend">{legend}</div>
    <div class="rm-scroll"><div class="rm-inner">
      <div class="rm-head"><div class="rm-side"></div><div class="rm-timeline rm-months">{months_html}{marker_head}</div></div>
      <div class="rm-rows">{overlay}{rows_html}</div>
    </div></div>"""


def render_bloqueios(bloqueios):
    if not bloqueios:
        return '<p class="empty-state">Nenhum bloqueio registrado.</p>'
    html = ""
    for b in bloqueios:
        impacto = str(b.get("Impacto") or "").strip()
        cor = IMPACT_COLORS.get(impacto, "#64748B")
        desde = to_date(b.get("Bloqueado desde"))
        resolvido = to_date(b.get("Data de resolução"))
        desde_txt = desde.strftime("%d/%m/%Y") if desde else "—"
        resolvido_txt = resolvido.strftime("%d/%m/%Y") if resolvido else "—"
        html += f"""
        <div class="hist-item">
          <div class="hist-main">
            <strong>{escape(str(b.get("Produto")))} — {escape(str(b.get("Épico")))}</strong>
            <span class="hist-badge" style="background:{cor}1A;color:{cor}">{escape(impacto)}</span>
          </div>
          <p class="hist-desc">{escape(str(b.get("Descrição do bloqueio") or ""))}</p>
          <div class="hist-meta">
            <span>🔓 {escape(str(b.get("Quem destrava") or "—"))}</span>
            <span>⏱ {int(num(b.get("Dias parado")))} dias parado</span>
            <span>{desde_txt} → <span class="resolvido-tag">Resolvido {resolvido_txt}</span></span>
          </div>
        </div>"""
    return html


def render_financeiro(fin):
    def tile(label, value, sub=""):
        sub_html = f'<div class="fin-sub">{sub}</div>' if sub else ""
        return (f'<div class="fin-tile"><div class="fin-label">{label}</div>'
                f'<div class="fin-value">{value}</div>{sub_html}</div>')

    tiles = (
        tile("Contratado", f"{fmt_pf0(fin['contratado'])} PF", "Total do contrato anual")
        + tile("Faturado", f"{fmt_pf0(fin['faturado_total'])} PF", f"{fmt_pct(fin['faturado_pct'])} do contratado")
        + tile("Saldo a faturar", f"{fmt_pf0(fin['saldo'])} PF", "Disponível no contrato")
    )

    linhas = ""
    for nome, d in fin["linhas"]:
        cls = " fin-total" if nome == "TOTAL" else ""
        linhas += (f'<tr class="{cls.strip()}"><td>{escape(nome)}</td>'
                   f'<td>{fmt_pf0(d["previsto"])}</td><td>{fmt_pf0(d["faturado"])}</td>'
                   f'<td>{fmt_pct(d["pct"])}</td></tr>')

    vig = fin["vig_fim"].strftime("%d/%m/%Y") if fin["vig_fim"] else "—"
    ritmo = (f"Ritmo necessário: <strong>{fmt_pf0(fin['media_mes'])} PF/mês</strong> "
             f"· {fin['meses']} meses restantes (vigência até {vig}) · "
             f"item mais antigo parado há {fin['item_antigo']} dias")

    # Janelas de faturamento
    sit_color = {"aberta": "#22C55E", "perdida": "#EF4444", "fora da vigência": "#94A3B8"}
    janelas_rows = ""
    for j in fin["janelas"]:
        cor = sit_color.get(j["situacao"].lower(), "#64748B")
        inic = "—" if j["iniciativas"] in (None, "") else int(num(j["iniciativas"]))
        pf = "—" if j["pf"] in (None, "") else f'{fmt_pf0(num(j["pf"]))} PF'
        janelas_rows += (
            f'<tr><td>{fmt_dm(j["janela"])}</td><td>{fmt_dm(j["enviar_ate"])}</td>'
            f'<td><span class="sit-badge" style="background:{cor}1A;color:{cor}">{escape(j["situacao"].capitalize())}</span></td>'
            f'<td>{inic}</td><td>{pf}</td></tr>'
        )

    # Proposta a negociar (valores em PF; "Cobertura" é uma razão, exibida em %)
    proposta_rows = ""
    for item in fin["proposta"]:
        destaque = any(t in item["label"].lower() for t in ["subtotal", "faturável", "cobertura"])
        cls = ' class="fin-total"' if destaque else ""
        if item["valor"] < 2:  # razão (ex.: 1,006 = 101%)
            valor = fmt_pct(item["valor"])
        else:
            valor = f'{fmt_pf0(item["valor"])} PF'
        proposta_rows += f'<tr{cls}><td>{escape(item["label"])}</td><td>{valor}</td></tr>'
    prop_extra = ("Falta planejar outras iniciativas despriorizadas do PSGE, ED. Básica, EPI, "
                  "ESR Monitor e Plat. Colaboração não estão mensuradas em 2026.")
    prop_nota_html = ""
    if fin["prop_nota"]:
        prop_nota_html += f'<p class="fin-hint">{escape(fin["prop_nota"])}</p>'
    prop_nota_html += f'<p class="fin-hint">{escape(prop_extra)}</p>'

    return f"""
    <div class="fin-cards">{tiles}</div>
    <div class="fin-table-wrap">
      <table class="fin-table">
        <thead><tr><th>Frente</th><th>Previsto (PF)</th><th>Faturado (PF)</th><th>% faturado</th></tr></thead>
        <tbody>{linhas}</tbody>
      </table>
    </div>
    <p class="fin-ritmo">{ritmo}</p>

    <h3 class="fin-sub-title">Janelas de faturamento</h3>
    <p class="fin-hint">Prazos mensais de envio ao SESI. Total faturável dentro do ciclo: <strong>{fmt_pf0(fin['total_faturavel'])} PF</strong>.<br>
    Estimativas considerando a esteira completa de desenvolvimento — ou seja, o faturamento pode correr algumas semanas antes do fim do delivery, ocorrendo a redistribuição dos valores mensais.</p>
    <div class="fin-table-wrap">
      <table class="fin-table fin-janelas">
        <thead><tr><th>Fatura em</th><th>Enviar até</th><th>Situação</th><th>Iniciativas</th><th>PF</th></tr></thead>
        <tbody>{janelas_rows}</tbody>
      </table>
    </div>

    <h3 class="fin-sub-title">Proposta a negociar para o próximo ciclo</h3>
    <div class="fin-table-wrap">
      <table class="fin-table fin-proposta">
        <tbody>{proposta_rows}</tbody>
      </table>
    </div>
    {prop_nota_html}"""


PAGE_TEMPLATE = """<!doctype html>
<html lang="pt-BR">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Portfólio Digital — Produtos Digitais</title>
<style>
  :root {{
    --navy: {navy};
    --orange: {orange};
    --bg: #F7F8FA;
    --card-bg: #FFFFFF;
    --text: #1E293B;
    --muted: #64748B;
    --border: #E2E8F0;
  }}
  * {{ box-sizing: border-box; }}
  body {{
    margin: 0;
    font-family: system-ui, -apple-system, "Segoe UI", Roboto, sans-serif;
    background: var(--bg);
    color: var(--text);
  }}
  header.top {{
    max-width: 1240px; margin: 0 auto; padding: 36px 24px 20px;
    border-bottom: 1px solid var(--border);
  }}
  header.top .brandrow {{ display: flex; align-items: center; gap: 10px; margin-bottom: 12px; }}
  header.top .logo {{ height: 40px; width: auto; display: block; }}
  header.top .brandtext {{ font-size: 0.9rem; font-weight: 700; color: var(--navy); }}
  header.top h1 {{ margin: 0; font-size: 2.1rem; font-weight: 800; color: var(--navy); letter-spacing: -0.01em; }}
  header.top .subtitle {{ margin: 10px 0 0; color: var(--muted); font-size: 0.95rem; }}

  /* Menu de abas */
  .tabs {{ max-width: 1240px; margin: 0 auto; padding: 18px 24px 0; display: flex; gap: 8px; flex-wrap: wrap; }}
  .tab {{ padding: 10px 18px; border: 1px solid var(--border); border-radius: 999px; background: #fff;
    color: var(--muted); font-size: 0.9rem; font-weight: 600; cursor: pointer; font-family: inherit; }}
  .tab:hover {{ border-color: #CBD5E1; color: var(--navy); }}
  .tab.active {{ background: var(--navy); color: #fff; border-color: var(--navy); }}
  .panel.hidden {{ display: none; }}

  /* Produtos internos — em construção */
  .construcao {{ text-align: center; padding: 64px 24px; }}
  .construcao .logo-big {{ height: 84px; opacity: 0.92; margin-bottom: 22px; }}
  .construcao h2 {{ font-size: 1.5rem; color: var(--navy); text-transform: none; letter-spacing: normal; margin: 0 0 10px; }}
  .construcao p {{ color: var(--muted); font-size: 1rem; margin: 6px 0; }}
  .construcao p.sub {{ max-width: 470px; margin: 10px auto 0; font-size: 0.9rem; }}
  .construcao .selo {{ display: inline-block; margin-top: 8px; background: #FEF3E7; color: var(--orange);
    font-size: 0.8rem; font-weight: 700; padding: 6px 14px; border-radius: 999px; }}

  main {{ max-width: 1240px; margin: 0 auto; padding: 28px 24px 80px; }}
  section {{ margin-bottom: 44px; }}
  section h2 {{
    font-size: 0.75rem; font-weight: 700; text-transform: uppercase;
    letter-spacing: 0.08em; color: #94A3B8; margin: 0 0 16px;
  }}

  /* KPIs */
  .kpi-row {{ display: grid; grid-template-columns: repeat(5, 1fr); gap: 16px; }}
  .kpi-card {{ position: relative; background: var(--card-bg); border: 1px solid var(--border);
    border-radius: 12px; padding: 22px 20px 18px; }}
  .kpi-card::before {{ content: ""; position: absolute; top: 0; left: 16px; right: 16px;
    height: 3px; border-radius: 0 0 4px 4px; background: var(--accent); }}
  .kpi-value {{ font-size: 2.6rem; font-weight: 800; color: #0F172A; line-height: 1; }}
  .kpi-label {{ color: #334155; font-size: 0.95rem; font-weight: 600; margin-top: 10px; }}
  .kpi-sub {{ color: #94A3B8; font-size: 0.8rem; margin-top: 4px; }}

  /* Cards por produto */
  .cards-grid {{ display: grid; grid-template-columns: repeat(4, 1fr); gap: 16px; }}
  .prod-card {{ position: relative; background: var(--card-bg); border: 1px solid var(--border);
    border-left: 4px solid #3B82F6; border-radius: 12px; padding: 18px 20px 20px; }}
  .prod-card.blocked {{ border-left-color: #EF4444; border-color: #FECDD3; background: #FEF6F6; }}
  .prod-head {{ display: flex; justify-content: space-between; align-items: flex-start; gap: 10px; }}
  .prod-head h3 {{ margin: 0; font-size: 1.05rem; font-weight: 700; color: var(--navy); }}
  .prod-sub {{ color: var(--muted); font-size: 0.82rem; margin-top: 3px; }}
  .sinal-dot {{ width: 11px; height: 11px; border-radius: 50%; flex-shrink: 0; margin-top: 5px; }}
  .seg-bar {{ display: flex; gap: 2px; height: 7px; margin: 15px 0 11px; }}
  .seg-bar span {{ border-radius: 2px; min-width: 3px; }}
  .prod-legend {{ display: flex; flex-wrap: wrap; gap: 5px 12px; font-size: 0.76rem; color: #475569; }}
  .prod-legend span {{ display: inline-flex; align-items: center; }}
  .prod-legend i {{ width: 7px; height: 7px; border-radius: 50%; display: inline-block; margin-right: 5px; }}
  .prod-foot {{ display: flex; gap: 14px; margin-top: 12px; padding-top: 10px;
    border-top: 1px solid var(--border); font-size: 0.78rem; color: var(--muted); }}
  .prod-next {{ font-size: 0.76rem; color: #94A3B8; margin-top: 6px; }}
  .prod-tag {{ display: inline-block; margin-top: 10px; background: #FEE2E2; color: #DC2626;
    font-size: 0.72rem; font-weight: 600; padding: 4px 10px; border-radius: 6px; }}

  /* Roadmap (Gantt) */
  .rm-legend {{ display: flex; gap: 16px; flex-wrap: wrap; margin-bottom: 12px; font-size: 0.78rem; color: var(--muted); }}
  .legend-item {{ display: flex; align-items: center; gap: 6px; }}
  .legend-dot {{ width: 11px; height: 11px; border-radius: 3px; display: inline-block; }}
  .rm-scroll {{ overflow-x: auto; background: var(--card-bg); border: 1px solid var(--border); border-radius: 12px; }}
  .rm-inner {{ min-width: 1180px; }}
  .rm-side {{ flex: 0 0 340px; display: flex; align-items: center; gap: 8px; padding: 0 12px; }}
  .rm-timeline {{ position: relative; flex: 1; }}
  .rm-head {{ display: flex; align-items: stretch; padding: 12px 0; border-bottom: 1px solid var(--border); }}
  .rm-months {{ display: flex; gap: 6px; padding-right: 12px; }}
  .rm-month {{ background: #1E293B; color: #fff; font-size: 0.72rem; font-weight: 700;
    letter-spacing: 0.05em; text-align: center; padding: 7px 0; border-radius: 6px; }}
  .rm-rows {{ padding: 6px 0; position: relative; }}
  .rm-row {{ display: flex; align-items: center; min-height: 38px; }}
  .rm-row:nth-child(even) {{ background: #FAFBFC; }}
  .rm-prod {{ flex: 0 0 62px; font-size: 0.66rem; font-weight: 700; color: #475569;
    background: #EEF2F6; border-radius: 5px; padding: 4px 2px; text-align: center;
    overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }}
  .rm-avatars {{ flex: 0 0 90px; display: flex; }}
  .rm-av {{ width: 22px; height: 22px; border-radius: 50%; color: #fff; font-size: 0.6rem;
    font-weight: 700; display: flex; align-items: center; justify-content: center;
    border: 2px solid #fff; margin-left: -6px; }}
  .rm-av:first-child {{ margin-left: 0; }}
  .rm-epic {{ flex: 1; background: #FBF3D3; border-radius: 6px; padding: 5px 10px;
    display: flex; flex-direction: column; overflow: hidden; }}
  .rm-epic b {{ font-size: 0.76rem; color: #4A3F1A; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }}
  .rm-epic i {{ font-size: 0.68rem; color: #A08A3A; font-style: normal; }}
  .rm-timeline {{ height: 34px; margin-right: 12px; }}
  .rm-rows .rm-timeline {{
    background-image: repeating-linear-gradient(to right, transparent 0, transparent calc(100%/12 - 1px), #EEF1F5 calc(100%/12 - 1px), #EEF1F5 calc(100%/12));
  }}
  .rm-bar {{ position: absolute; top: 5px; height: 24px; border-radius: 5px; padding: 0 8px;
    display: flex; flex-direction: column; justify-content: center; overflow: hidden; }}
  .rm-bar b {{ font-size: 0.68rem; font-weight: 700; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }}
  .rm-bar small {{ font-size: 0.6rem; opacity: 0.85; line-height: 1; }}
  /* Marcador "estamos aqui" (linha contínua na data de hoje) */
  .rm-now-overlay {{ position: absolute; top: 0; bottom: 0; left: 0; right: 0;
    display: flex; pointer-events: none; z-index: 3; }}
  .rm-now-spacer {{ flex: 0 0 340px; }}
  .rm-now-track {{ position: relative; flex: 1; margin-right: 12px; }}
  .rm-now-line {{ position: absolute; top: 0; bottom: 0; width: 0; border-left: 1.6px dashed {orange}; }}
  .rm-now {{ position: absolute; top: 0; bottom: 0; width: 0; border-left: 1.6px dashed {orange}; }}
  .rm-now span {{ position: absolute; bottom: -9px; left: 50%; transform: translateX(-50%);
    background: {orange}; color: #fff; font-size: 0.56rem; font-weight: 700;
    padding: 1px 6px; border-radius: 4px; white-space: nowrap; }}

  /* Histórico de bloqueios */
  .hist-item {{ background: var(--card-bg); border: 1px solid var(--border);
    border-left: 4px solid #94A3B8; border-radius: 8px; padding: 14px 16px; margin-bottom: 10px; }}
  .hist-main {{ display: flex; justify-content: space-between; align-items: center; gap: 10px; }}
  .hist-main strong {{ color: var(--navy); }}
  .hist-badge {{ font-size: 0.72rem; font-weight: 600; padding: 3px 9px; border-radius: 999px; }}
  .hist-desc {{ margin: 7px 0; color: var(--text); font-size: 0.86rem; }}
  .hist-meta {{ display: flex; gap: 18px; flex-wrap: wrap; align-items: center; font-size: 0.78rem; color: var(--muted); }}
  .resolvido-tag {{ background: #DCFCE7; color: #16A34A; font-size: 0.72rem; font-weight: 600;
    padding: 3px 10px; border-radius: 999px; }}
  .empty-state {{ color: var(--muted); font-style: italic; }}

  /* Financeiro */
  .fin-cards {{ display: grid; grid-template-columns: repeat(3, 1fr); gap: 16px; margin-bottom: 20px; }}
  .fin-tile {{ background: var(--card-bg); border: 1px solid var(--border); border-radius: 12px;
    padding: 20px; border-top: 3px solid var(--orange); }}
  .fin-label {{ font-size: 0.78rem; color: var(--muted); text-transform: uppercase; letter-spacing: 0.04em; }}
  .fin-value {{ font-size: 1.8rem; font-weight: 800; color: var(--navy); margin-top: 6px; }}
  .fin-sub {{ font-size: 0.78rem; color: #94A3B8; margin-top: 4px; }}
  .fin-table-wrap {{ overflow-x: auto; }}
  .fin-table {{ width: 100%; border-collapse: collapse; background: var(--card-bg);
    border: 1px solid var(--border); border-radius: 12px; overflow: hidden; }}
  .fin-table th, .fin-table td {{ padding: 12px 16px; text-align: right; font-size: 0.88rem;
    border-bottom: 1px solid var(--border); }}
  .fin-table th:first-child, .fin-table td:first-child {{ text-align: left; }}
  .fin-table th {{ font-size: 0.72rem; text-transform: uppercase; letter-spacing: 0.04em;
    color: var(--muted); font-weight: 700; background: #F8FAFC; }}
  .fin-table tr.fin-total td {{ font-weight: 700; color: var(--navy); border-bottom: none; background: #F8FAFC; }}
  .fin-ritmo {{ margin-top: 14px; font-size: 0.85rem; color: var(--muted); }}
  .fin-sub-title {{ margin: 28px 0 4px; font-size: 1rem; color: var(--navy); }}
  .fin-hint {{ margin: 0 0 12px; font-size: 0.85rem; color: var(--muted); }}
  .sit-badge {{ font-size: 0.72rem; font-weight: 600; padding: 3px 10px; border-radius: 999px; }}
  .fin-janelas td:last-child {{ font-weight: 600; color: var(--navy); }}

  footer {{ text-align: center; color: var(--muted); font-size: 0.75rem; padding: 24px;
    display: flex; align-items: center; justify-content: center; gap: 8px; flex-wrap: wrap; }}
  footer .foot-logo {{ height: 20px; opacity: 0.7; }}

  @media (max-width: 1080px) {{
    .kpi-row {{ grid-template-columns: repeat(3, 1fr); }}
    .cards-grid {{ grid-template-columns: repeat(2, 1fr); }}
    .fin-cards {{ grid-template-columns: 1fr; }}
  }}
  @media (max-width: 720px) {{
    .kpi-row {{ grid-template-columns: repeat(2, 1fr); }}
    .cards-grid {{ grid-template-columns: 1fr; }}
    header.top h1 {{ font-size: 1.6rem; }}
  }}
</style>
</head>
<body>
<header class="top">
  <div class="brandrow">
    <img class="logo" src="{logo}" alt="Grupo Fênix Educação">
    <span class="brandtext">Grupo Fênix Educação</span>
  </div>
  <h1>Produtos Digitais - 2026</h1>
  <p class="subtitle">Atualização: {updated_at} · {num_produtos} produtos · {num_iniciativas} iniciativas estimadas</p>
</header>

<nav class="tabs">
  <button class="tab active" data-tab="roadmap">Roadmap geral</button>
  <button class="tab" data-tab="pf">Contrato SESI</button>
  <button class="tab" data-tab="internos">Produtos internos</button>
</nav>

<main>
  <div class="panel" id="panel-roadmap">
    <section>
      <h2>Visão geral — Portfólio</h2>
      <div class="kpi-row">{kpis_html}</div>
    </section>
    <section>
      <h2>Visão por produto</h2>
      <div class="cards-grid">{cards_html}</div>
    </section>
    <section>
      <h2>Roadmap — Iniciativas 2026</h2>
      {roadmap_html}
    </section>
    <section>
      <h2>Histórico de bloqueios</h2>
      {bloqueios_html}
    </section>
  </div>

  <div class="panel hidden" id="panel-pf">
    <section>
      <h2>Pontos de Função</h2>
      {financeiro_html}
    </section>
  </div>

  <div class="panel hidden" id="panel-internos">
    <div class="construcao">
      <img class="logo-big" src="{logo}" alt="">
      <h2>Produtos internos</h2>
      <span class="selo">🚧 Em construção</span>
      <p class="sub">Em breve, o acompanhamento dos produtos internos da Fênix (LeituraTech, Timeline, FNXCore e mais) — com roadmap, status e métricas próprias.</p>
    </div>
  </div>
</main>

<footer>
  <img class="foot-logo" src="{logo}" alt="">
  <span>Grupo Fênix Educação · gerado automaticamente a partir da planilha de gestão do portfólio.</span>
</footer>
<script>{tab_script}</script>
</body>
</html>
"""


# ---------------------------------------------------------------------------
# Proteção por senha (criptografia AES-GCM client-side)
# ---------------------------------------------------------------------------

GATE_TEMPLATE = """<!doctype html>
<html lang="pt-BR">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Portfólio Digital — Acesso</title>
<style>
  * { box-sizing: border-box; }
  body { margin: 0; min-height: 100vh; display: flex; align-items: center; justify-content: center;
    font-family: system-ui, -apple-system, "Segoe UI", Roboto, sans-serif; background: #0D2137; color: #E2E8F0; }
  .gate { width: 100%; max-width: 380px; padding: 40px 32px; text-align: center; }
  .brand { font-size: 0.85rem; font-weight: 700; color: #F97316; letter-spacing: 0.04em; }
  h1 { margin: 12px 0 6px; font-size: 1.6rem; color: #fff; }
  p.sub { margin: 0 0 24px; color: #94A3B8; font-size: 0.9rem; }
  input { width: 100%; padding: 13px 14px; border-radius: 10px; border: 1px solid #24405E;
    background: #0A1A2B; color: #fff; font-size: 1rem; outline: none; }
  input:focus { border-color: #F97316; }
  button { width: 100%; margin-top: 12px; padding: 13px; border: 0; border-radius: 10px;
    background: #F97316; color: #fff; font-size: 1rem; font-weight: 700; cursor: pointer; }
  button:hover { background: #EA6A0C; }
  .err { color: #FCA5A5; font-size: 0.85rem; min-height: 20px; margin-top: 12px; }
  .gate-logo { height: 72px; margin-bottom: 10px; }
</style>
</head>
<body>
<div class="gate">
  <img class="gate-logo" src="__LOGO__" alt="Grupo Fênix">
  <div class="brand">Grupo Fênix Educação</div>
  <h1>Portfólio Digital</h1>
  <p class="sub">Painel protegido. Digite a senha de acesso.</p>
  <form id="f">
    <input type="password" id="pw" autofocus placeholder="Senha" autocomplete="current-password">
    <button type="submit">Entrar</button>
  </form>
  <p class="err" id="err"></p>
</div>
<script>
  var D = { salt: "__SALT__", iv: "__IV__", ct: "__CT__", iter: __ITER__ };
  function b64(s){ return Uint8Array.from(atob(s), function(c){ return c.charCodeAt(0); }); }
  async function decrypt(pw){
    var enc = new TextEncoder();
    var mat = await crypto.subtle.importKey("raw", enc.encode(pw), "PBKDF2", false, ["deriveKey"]);
    var key = await crypto.subtle.deriveKey(
      { name:"PBKDF2", salt:b64(D.salt), iterations:D.iter, hash:"SHA-256" },
      mat, { name:"AES-GCM", length:256 }, false, ["decrypt"]);
    var out = await crypto.subtle.decrypt({ name:"AES-GCM", iv:b64(D.iv) }, key, b64(D.ct));
    return new TextDecoder().decode(out);
  }
  document.getElementById("f").addEventListener("submit", async function(e){
    e.preventDefault();
    var err = document.getElementById("err"); err.textContent = "Verificando…";
    try {
      var html = await decrypt(document.getElementById("pw").value);
      document.open(); document.write(html); document.close();
    } catch (_) { err.textContent = "Senha incorreta."; }
  });
</script>
</body>
</html>
"""


def encrypt_page(inner_html, password, logo):
    """Empacota o dashboard numa página com senha (AES-256-GCM + PBKDF2)."""
    from cryptography.hazmat.primitives.ciphers.aead import AESGCM

    salt = os.urandom(16)
    iv = os.urandom(12)
    key = pbkdf2_hmac("sha256", password.encode("utf-8"), salt, PBKDF2_ITER, dklen=32)
    ct = AESGCM(key).encrypt(iv, inner_html.encode("utf-8"), None)  # ct||tag
    b = lambda raw: base64.b64encode(raw).decode("ascii")
    return (GATE_TEMPLATE
            .replace("__LOGO__", logo)
            .replace("__SALT__", b(salt))
            .replace("__IV__", b(iv))
            .replace("__CT__", b(ct))
            .replace("__ITER__", str(PBKDF2_ITER)))


def read_password():
    if PASSWORD_FILE.exists():
        pw = PASSWORD_FILE.read_text(encoding="utf-8").strip()
        return pw or None
    return None


def build():
    wb = load_workbook(PORTFOLIO_FILE, read_only=True, data_only=True)
    fin = extract_financeiro(wb["Contrato"])
    produtos = extract_portfolio(wb["Portfólio"])
    kpis = kpis_from_portfolio(produtos)
    roadmap = extract_roadmap(wb["Iniciativas"])
    bloqueios = extract_bloqueios(wb["Bloqueios"])
    wb.close()

    logo = logo_data_uri()
    html = PAGE_TEMPLATE.format(
        navy=NAVY, orange=ORANGE,
        logo=logo,
        tab_script=TAB_SCRIPT,
        updated_at=datetime.now().strftime("%d/%m/%Y"),
        num_produtos=len(produtos),
        num_iniciativas=len(roadmap),
        kpis_html=render_kpis(kpis),
        cards_html=render_cards(produtos),
        roadmap_html=render_roadmap(roadmap),
        bloqueios_html=render_bloqueios(bloqueios),
        financeiro_html=render_financeiro(fin),
    )

    password = read_password()
    if password:
        OUTPUT_FILE.write_text(encrypt_page(html, password, logo), encoding="utf-8")
        protecao = "🔒 protegido por senha (AES-256)"
    else:
        OUTPUT_FILE.write_text(html, encoding="utf-8")
        protecao = ("⚠️  SEM SENHA — publicado sem proteção! "
                    f"Crie {PASSWORD_FILE.name} com a senha para proteger.")

    print(f"✅ {OUTPUT_FILE.name} gerado — {protecao}")
    print(f"   KPIs: {kpis}")
    print(f"   Produtos (cards): {len(produtos)}")
    print(f"   Iniciativas no roadmap: {len(roadmap)}")
    print(f"   Bloqueios (histórico): {len(bloqueios)}")
    print(f"   Financeiro: contratado {fin['contratado']}, faturado {fin['faturado_total']}")


# ---------------------------------------------------------------------------
# Publicação no GitHub
# ---------------------------------------------------------------------------

def is_git_repo():
    return (BASE_DIR / ".git").is_dir()


def has_remote():
    result = subprocess.run(["git", "remote"], cwd=BASE_DIR, capture_output=True, text=True)
    return "origin" in result.stdout.split()


def publish():
    if not is_git_repo() or not has_remote():
        print("ℹ️  Repositório git ainda não configurado — HTML gerado, mas não publicado.")
        return
    subprocess.run(["git", "add", "index.html"], cwd=BASE_DIR, check=True)
    if subprocess.run(["git", "diff", "--cached", "--quiet"], cwd=BASE_DIR).returncode == 0:
        print("ℹ️  Nenhuma mudança no dashboard desde o último commit.")
        return
    msg = f"Atualiza dashboard — {datetime.now().strftime('%d/%m/%Y %H:%M')}"
    subprocess.run(["git", "commit", "-m", msg], cwd=BASE_DIR, check=True)
    push = subprocess.run(["git", "push", "origin", "main"], cwd=BASE_DIR, capture_output=True, text=True)
    if push.returncode != 0:
        print("⚠️  Commit feito localmente, mas o push falhou:")
        print(push.stderr.strip())
        print("   Rode `git push origin main` manualmente para publicar.")
    else:
        print("🚀 Publicado no GitHub Pages.")


if __name__ == "__main__":
    build()
    publish()
