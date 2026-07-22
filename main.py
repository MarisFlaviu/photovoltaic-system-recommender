import math
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import List

import pandas as pd
import streamlit as st


DATA_DIR = Path(__file__).resolve().parent


st.set_page_config(
    page_title="Estimare costuri panouri fotovoltaice",
    page_icon=":sunny:",
    layout="wide",
    initial_sidebar_state="expanded",
)


CSS = """
<style>
    :root {
        --green: #0f766e;
        --dark: #0f172a;
        --muted: #64748b;
        --line: #e2e8f0;
    }
    .stApp {
        background:
            radial-gradient(circle at top left, rgba(15,118,110,.13), transparent 28rem),
            linear-gradient(180deg, #fff 0%, #f8fafc 45%, #eef2f7 100%);
        color: var(--dark);
    }
    .block-container {
        padding-top: 2rem;
        padding-bottom: 3rem;
        max-width: 1320px;
    }
    h1 {
        font-size: 2.45rem !important;
        line-height: 1.08 !important;
        color: var(--dark);
    }
    h2, h3 {
        color: var(--dark);
        letter-spacing: 0;
    }
    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #0f172a 0%, #0f766e 100%);
    }
    [data-testid="stSidebar"] * {
        color: #fff;
    }
    [data-testid="stSidebar"] input,
    [data-testid="stSidebar"] textarea {
        color: #0f172a !important;
    }
    .hero {
        padding: 2rem 2.2rem;
        border: 1px solid rgba(15,118,110,.18);
        border-radius: 18px;
        background: linear-gradient(135deg, rgba(255,255,255,.96), rgba(240,253,250,.92));
        box-shadow: 0 20px 45px rgba(15,23,42,.08);
        margin-bottom: 1.35rem;
    }
    .tag {
        display: inline-flex;
        padding: .35rem .7rem;
        border-radius: 999px;
        background: rgba(15,118,110,.12);
        color: #115e59;
        font-weight: 800;
        font-size: .82rem;
        text-transform: uppercase;
        letter-spacing: .08em;
        margin-bottom: .8rem;
    }
    .hero p {
        color: #475569;
        font-size: 1.05rem;
        line-height: 1.65;
        max-width: 880px;
    }
    .metric {
        border: 1px solid var(--line);
        border-radius: 14px;
        padding: 1rem 1.1rem;
        background: rgba(255,255,255,.94);
        box-shadow: 0 12px 28px rgba(15,23,42,.06);
        min-height: 118px;
    }
    .metric-label {
        color: var(--muted);
        font-size: .82rem;
        font-weight: 800;
        text-transform: uppercase;
        letter-spacing: .06em;
        margin-bottom: .4rem;
    }
    .metric-value {
        color: var(--dark);
        font-size: 1.68rem;
        font-weight: 850;
        line-height: 1.15;
    }
    .metric-help {
        color: var(--muted);
        font-size: .9rem;
        margin-top: .45rem;
    }
    .ok, .warn {
        border-radius: 18px;
        padding: 1.35rem 1.45rem;
        background: rgba(255,255,255,.95);
        box-shadow: 0 18px 38px rgba(15,23,42,.08);
    }
    .ok {
        border: 1px solid rgba(22,163,74,.25);
        background: linear-gradient(135deg, rgba(240,253,244,.97), rgba(255,255,255,.94));
    }
    .warn {
        border: 1px solid rgba(245,158,11,.30);
        background: linear-gradient(135deg, rgba(255,251,235,.98), rgba(255,255,255,.94));
    }
    .badge {
        display: inline-flex;
        border-radius: 999px;
        padding: .25rem .62rem;
        font-size: .78rem;
        font-weight: 800;
        background: rgba(15,118,110,.12);
        color: #115e59;
        margin-right: .35rem;
        margin-bottom: .35rem;
    }
    .muted {
        color: #64748b;
        font-size: .92rem;
        line-height: 1.55;
    }
    .footer {
        margin-top: 2rem;
        padding: 1rem 1.1rem;
        border-radius: 12px;
        background: #0f172a;
        color: #e2e8f0;
        font-size: .9rem;
    }
    div[data-testid="stDataFrame"],
    div[data-testid="stExpander"] {
        border: 1px solid #e2e8f0;
        border-radius: 12px;
        overflow: hidden;
    }
    .stButton > button {
        border-radius: 10px;
        border: 0;
        background: linear-gradient(135deg, #0f766e, #115e59);
        color: white;
        font-weight: 800;
        min-height: 2.75rem;
        box-shadow: 0 10px 20px rgba(15,118,110,.18);
    }
    .stButton > button:hover {
        border: 0;
        color: white;
        background: linear-gradient(135deg, #115e59, #0f766e);
    }
</style>
"""


@dataclass
class Panel:
    nume: str
    putere: float
    eficienta: float
    suprafata: float
    pret: float
    tehnologie: str
    descriere: str


@dataclass
class Baterie:
    nume: str
    capacitate: float
    pret: float
    tehnologie: str
    descriere: str


@dataclass
class Consumator:
    nume: str
    putere: float
    ore: float

    @property
    def consum(self) -> float:
        return self.putere * self.ore / 1000


@dataclass
class Varianta:
    panel: Panel
    baterie: Baterie
    nr_panouri: int
    nr_baterii: int
    productie: float
    consum: float
    suprafata: float
    cost_panouri: float
    cost_baterii: float
    cost_montaj: float
    cost_total: float
    buget_ramas: float
    acoperire: float
    scor: float


def citeste_baza_date(nume: str, sheet: str) -> pd.DataFrame:
    csv_path = DATA_DIR / f"{nume}.csv"
    xlsx_path = DATA_DIR / f"{nume}.xlsx"
    workbook_path = DATA_DIR / "baze_date.xlsx"
    if workbook_path.exists():
        return pd.read_excel(workbook_path, sheet_name=sheet)
    if csv_path.exists():
        return pd.read_csv(csv_path)
    if xlsx_path.exists():
        return pd.read_excel(xlsx_path)
    st.error(
        f"Lipseste baza de date pentru {nume}. Pune fisierul {nume}.csv "
        f"sau {nume}.xlsx in folderul proiectului."
    )
    st.stop()


def verifica_coloane(df: pd.DataFrame, coloane: List[str], nume: str) -> None:
    lipsa = [col for col in coloane if col not in df.columns]
    if lipsa:
        st.error(f"Fisierul {nume} nu contine coloanele necesare: {', '.join(lipsa)}")
        st.stop()


def panouri_db() -> List[Panel]:
    df = citeste_baza_date("panouri", "Panouri")
    verifica_coloane(df, ["nume", "putere", "eficienta", "suprafata", "pret", "tehnologie", "descriere"], "panouri")
    return [
        Panel(str(r["nume"]), float(r["putere"]), float(r["eficienta"]),
              float(r["suprafata"]), float(r["pret"]), str(r["tehnologie"]), str(r["descriere"]))
        for _, r in df.iterrows()
    ]


def baterii_db() -> List[Baterie]:
    df = citeste_baza_date("baterii", "Baterii")
    verifica_coloane(df, ["nume", "capacitate", "pret", "tehnologie", "descriere"], "baterii")
    return [
        Baterie(str(r["nume"]), float(r["capacitate"]), float(r["pret"]),
                str(r["tehnologie"]), str(r["descriere"]))
        for _, r in df.iterrows()
    ]


def lei(x: float) -> str:
    return f"{x:,.0f} lei".replace(",", ".")


def kwh(x: float) -> str:
    return f"{x:.2f} kWh"


def m2(x: float) -> str:
    return f"{x:.2f} m2"


def total_consum(consumatori: List[Consumator]) -> float:
    return sum(c.consum for c in consumatori)


def consumatori_predefiniti() -> pd.DataFrame:
    return pd.DataFrame([
        {"Include": False, "Consumator": "Cuptor electric", "Putere (W)": 2000.0, "Ore/zi": 1.0},
        {"Include": False, "Consumator": "Cuptor cu microunde", "Putere (W)": 1200.0, "Ore/zi": 0.3},
        {"Include": False, "Consumator": "Plita electrica", "Putere (W)": 2000.0, "Ore/zi": 1.0},
        {"Include": False, "Consumator": "Aspirator", "Putere (W)": 800.0, "Ore/zi": 0.3},
        {"Include": False, "Consumator": "Masina de spalat", "Putere (W)": 800.0, "Ore/zi": 1.0},
        {"Include": False, "Consumator": "Uscator de par", "Putere (W)": 1600.0, "Ore/zi": 0.2},
        {"Include": False, "Consumator": "Masina de spalat vase", "Putere (W)": 1800.0, "Ore/zi": 1.0},
        {"Include": False, "Consumator": "Pompa de apa", "Putere (W)": 750.0, "Ore/zi": 1.0},
        {"Include": False, "Consumator": "Frigider", "Putere (W)": 150.0, "Ore/zi": 8.0},
        {"Include": False, "Consumator": "Incalzitor de baie", "Putere (W)": 1500.0, "Ore/zi": 0.5},
        {"Include": False, "Consumator": "Boiler electric", "Putere (W)": 2000.0, "Ore/zi": 2.0},
        {"Include": False, "Consumator": "Uscator de rufe", "Putere (W)": 2500.0, "Ore/zi": 1.0},
        {"Include": False, "Consumator": "Fier de calcat", "Putere (W)": 2000.0, "Ore/zi": 0.5},
        {"Include": False, "Consumator": "Calorifer electric", "Putere (W)": 2000.0, "Ore/zi": 2.0},
        {"Include": False, "Consumator": "Aer conditionat", "Putere (W)": 1200.0, "Ore/zi": 4.0},
    ])


def df_consumatori(consumatori: List[Consumator]) -> pd.DataFrame:
    return pd.DataFrame([{
        "Consumator": c.nume,
        "Putere (W)": c.putere,
        "Ore/zi": c.ore,
        "Consum (kWh/zi)": c.consum,
    } for c in consumatori])


def df_panouri(panouri: List[Panel]) -> pd.DataFrame:
    return pd.DataFrame([{
        "Panou": p.nume,
        "Putere (W)": p.putere,
        "Eficienta (%)": p.eficienta,
        "Suprafata (m2)": p.suprafata,
        "Pret (lei)": p.pret,
        "Tehnologie": p.tehnologie,
    } for p in panouri])


def df_baterii(baterii: List[Baterie]) -> pd.DataFrame:
    return pd.DataFrame([{
        "Baterie": b.nume,
        "Capacitate (kWh)": b.capacitate,
        "Pret (lei)": b.pret,
        "Tehnologie": b.tehnologie,
    } for b in baterii])


def df_variante(variante: List[Varianta]) -> pd.DataFrame:
    return pd.DataFrame([{
        "Loc": i,
        "Panou": v.panel.nume,
        "Baterie": v.baterie.nume,
        "Nr panouri": v.nr_panouri,
        "Nr baterii": v.nr_baterii,
        "Energie produsa (kWh/zi)": v.productie,
        "Consum (kWh/zi)": v.consum,
        "Acoperire consum (%)": v.acoperire,
        "Suprafata folosita (m2)": v.suprafata,
        "Cost panouri (lei)": v.cost_panouri,
        "Cost baterii (lei)": v.cost_baterii,
        "Cost montaj (lei)": v.cost_montaj,
        "Cost total (lei)": v.cost_total,
        "Buget ramas (lei)": v.buget_ramas,
    } for i, v in enumerate(variante, 1)])


def salveaza_recomandari(variante: List[Varianta], consumatori: List[Consumator]) -> Path:
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M")
    output_path = DATA_DIR / f"recomandari_{timestamp}.xlsx"
    recomandari_df = df_variante(variante[:5])
    consumatori_df = df_consumatori(consumatori)
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        recomandari_df.to_excel(writer, index=False, sheet_name="Recomandari")
        consumatori_df.to_excel(writer, index=False, sheet_name="Consumatori")
        for sheet_name in writer.book.sheetnames:
            ws = writer.book[sheet_name]
            for cell in ws[1]:
                cell.font = cell.font.copy(bold=True, color="FFFFFF")
                cell.fill = cell.fill.copy(fill_type="solid", fgColor="0F766E")
            for column in ws.columns:
                letter = column[0].column_letter
                width = max(len("" if cell.value is None else str(cell.value)) for cell in column) + 2
                ws.column_dimensions[letter].width = min(max(width, 14), 42)
    return output_path


def calculeaza_variante(
    consumatori: List[Consumator],
    panouri: List[Panel],
    baterii: List[Baterie],
    buget: float,
    acoperis: float,
    ore_soare: float,
    montaj: float,
) -> List[Varianta]:
    consum = total_consum(consumatori)
    tinta = consum
    variante = []
    if consum <= 0 or buget <= 0 or acoperis <= 0 or ore_soare <= 0:
        return variante
    for p in panouri:
        productie_panou = p.putere / 1000 * ore_soare
        minim_panouri = max(1, math.ceil(tinta / productie_panou))
        maxim_panouri = math.floor(acoperis / p.suprafata)
        for nr_p in range(minim_panouri, maxim_panouri + 1):
            productie = nr_p * productie_panou
            suprafata = nr_p * p.suprafata
            for b in baterii:
                minim_bat = max(1, math.ceil(consum / b.capacitate))
                for nr_b in range(minim_bat, minim_bat + 3):
                    c_pan = nr_p * p.pret
                    c_bat = nr_b * b.pret
                    c_montaj = (c_pan + c_bat) * montaj / 100
                    total = c_pan + c_bat + c_montaj
                    if total > buget:
                        continue
                    ramas = buget - total
                    acoperire = productie / consum * 100
                    scor_energie = 1 - abs(productie - tinta) / max(tinta, 1)
                    scor_buget = 1 - ramas / max(buget, 1)
                    scor_spatiu = suprafata / max(acoperis, 1)
                    scor_baterie = 120 / max(b.pret / b.capacitate, 1)
                    scor = scor_energie * .43 + scor_buget * .26 + scor_spatiu * .08 + scor_baterie * .23
                    variante.append(Varianta(
                        p, b, nr_p, nr_b, productie, consum, suprafata,
                        c_pan, c_bat, c_montaj, total, ramas, acoperire, scor
                    ))
    return sorted(variante, key=lambda x: x.scor, reverse=True)


def metric(label: str, value: str, helper: str = "") -> None:
    st.markdown(f"""
    <div class="metric">
        <div class="metric-label">{label}</div>
        <div class="metric-value">{value}</div>
        <div class="metric-help">{helper}</div>
    </div>
    """, unsafe_allow_html=True)


def hero() -> None:
    st.markdown("""
    <div class="hero">
        <div class="tag">Configurator energetic</div>
        <h1>Optimizarea investiției în panouri fotovoltaice</h1>
        <p>
            Introdu consumatorii electrici, bugetul si suprafata acoperisului.
            Aplicatia compara panouri solare si baterii, include montajul si
            alege varianta care acopera consumul zilnic cat mai eficient.
        </p>
    </div>
    """, unsafe_allow_html=True)


def sidebar_inputs():
    st.sidebar.title("Date pentru dimensionarea sistemului")
    st.sidebar.caption("Completeaza ipotezele principale pentru dimensionare.")
    buget = st.sidebar.number_input("Buget disponibil (lei)", 1000.0, 500000.0, 25000.0, 500.0)
    acoperis = st.sidebar.number_input("Suprafata acoperisului (m2)", 1.0, 500.0, 35.0, 1.0)
    ore_soare = st.sidebar.slider("Ore soare echivalent / zi", 2.0, 6.5, 4.0, .1)
    montaj = st.sidebar.slider("Cost montaj (% din componente)", 0.0, 50.0, 20.0, 1.0)
    st.sidebar.markdown("---")
    return buget, acoperis, ore_soare, montaj


def editor_consumatori() -> List[Consumator]:
    st.subheader("Introdu consumatorii")
    st.caption(
        "Consumatorii sunt predefiniti dupa lista din chestionar. "
        "Bifeaza aparatele folosite si modifica valorile medii daca este nevoie."
    )
    rezultat = []
    for i, row in consumatori_predefiniti().iterrows():
        nume = str(row.get("Consumator", "")).strip()
        with st.container(border=True):
            include_col, titlu = st.columns([0.5, 8])
            with include_col:
                include = st.checkbox(
                    "Include in calcul",
                    value=bool(row.get("Include", False)),
                    key=f"include_{i}",
                    help="Include acest consumator in calcul",
                    label_visibility="collapsed",
                )
            with titlu:
                st.markdown(f"**Consumator {i + 1}: {nume}**")
            c1, c2, c3 = st.columns([2, 1, 1])
            with c1:
                nume = st.text_input("Nume consumator", value=nume, key=f"nume_{i}")
            with c2:
                putere = st.number_input(
                    "Putere (W)",
                    min_value=0.0,
                    max_value=100000.0,
                    value=float(row.get("Putere (W)", 0) or 0),
                    step=10.0,
                    key=f"putere_{i}",
                )
            with c3:
                ore = st.number_input(
                    "Ore/zi",
                    min_value=0.0,
                    max_value=24.0,
                    value=float(row.get("Ore/zi", 0) or 0),
                    step=0.1,
                    key=f"ore_{i}",
                )
        if include and nume.strip() and putere > 0 and ore > 0:
            rezultat.append(Consumator(nume, putere, ore))
    st.markdown("#### Consumatori adaugati de utilizator")
    nr_extra = st.number_input(
        "Cati consumatori noi vrei sa adaugi?",
        min_value=0,
        max_value=5,
        value=0,
        step=1,
    )
    for i in range(int(nr_extra)):
        with st.container(border=True):
            st.markdown(f"**Consumator nou {i + 1}**")
            c1, c2, c3 = st.columns([2, 1, 1])
            with c1:
                nume = st.text_input(
                    "Nume consumator nou",
                    value="",
                    placeholder="Ex: Televizor",
                    key=f"extra_nume_{i}",
                )
            with c2:
                putere = st.number_input(
                    "Putere noua (W)",
                    min_value=0.0,
                    max_value=100000.0,
                    value=0.0,
                    step=10.0,
                    key=f"extra_putere_{i}",
                )
            with c3:
                ore = st.number_input(
                    "Ore/zi nou",
                    min_value=0.0,
                    max_value=24.0,
                    value=0.0,
                    step=0.1,
                    key=f"extra_ore_{i}",
                )
        if nume.strip() and putere > 0 and ore > 0:
            rezultat.append(Consumator(nume.strip(), putere, ore))
    return rezultat


def afisare_best(best: Varianta) -> None:
    st.markdown(f"""
    <div class="ok">
        <h3>Recomandare optima</h3>
        <p class="muted">
            Varianta selectata se incadreaza in buget si in suprafata disponibila.
        </p>
        <span class="badge">{best.panel.nume}</span>
        <span class="badge">{best.baterie.nume}</span>
    </div>
    """, unsafe_allow_html=True)
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        metric("Panouri", str(best.nr_panouri), f"{best.panel.putere:.0f} W / bucata")
    with c2:
        metric("Baterii", str(best.nr_baterii), f"{best.baterie.capacitate:.2f} kWh / bucata")
    with c3:
        metric("Productie", kwh(best.productie), "estimare zilnica")
    with c4:
        metric("Cost total", lei(best.cost_total), "componente + montaj")


def afisare_fara_solutie() -> None:
    st.markdown("""
    <div class="warn">
        <h3>Nu exista o varianta eligibila</h3>
        <p class="muted">
            Nicio combinatie panou + baterie nu se incadreaza simultan in buget
            si in suprafata. Mareste bugetul, suprafata sau redu consumul.
        </p>
    </div>
    """, unsafe_allow_html=True)


def analiza(best: Varianta, acoperis: float) -> None:
    st.subheader("Analiza tehnica")
    c1, c2, c3 = st.columns(3)
    with c1:
        metric("Acoperire consum", f"{best.acoperire:.1f}%", "productie / consum")
    with c2:
        metric("Suprafata folosita", m2(best.suprafata), f"din {m2(acoperis)} disponibili")
    with c3:
        metric("Stocare instalata", kwh(best.nr_baterii * best.baterie.capacitate))
    st.subheader("Detaliere costuri")
    cost_df = pd.DataFrame([
        {"Categorie": "Panouri", "Cost lei": best.cost_panouri},
        {"Categorie": "Baterii", "Cost lei": best.cost_baterii},
        {"Categorie": "Montaj", "Cost lei": best.cost_montaj},
    ])
    cost_df["Valoare"] = cost_df["Cost lei"].apply(lambda x: f"{x:,.0f} lei".replace(",", "."))
    st.vega_lite_chart(
        cost_df,
        {
            "height": 360,
            "layer": [
                {
                    "mark": {"type": "bar", "color": "#0f766e"},
                    "encoding": {
                        "x": {"field": "Categorie", "type": "nominal", "title": "Categorie"},
                        "y": {"field": "Cost lei", "type": "quantitative", "title": "Cost (lei)"},
                        "tooltip": [
                            {"field": "Categorie", "type": "nominal"},
                            {"field": "Valoare", "type": "nominal", "title": "Cost"},
                        ],
                    },
                },
                {
                    "mark": {"type": "text", "dy": -8, "fontWeight": "bold", "color": "#0f172a"},
                    "encoding": {
                        "x": {"field": "Categorie", "type": "nominal"},
                        "y": {"field": "Cost lei", "type": "quantitative"},
                        "text": {"field": "Valoare", "type": "nominal"},
                    },
                },
            ],
        },
        use_container_width=True,
    )


def afisare_variante(variante: List[Varianta], top: int) -> pd.DataFrame:
    st.subheader("Soluții fotovoltaice alternative")
    df = df_variante(variante)
    if df.empty:
        st.info("Nu exista variante de afisat.")
    else:
        st.dataframe(df.head(top), use_container_width=True, hide_index=True)
    return df


def baze_date(panouri: List[Panel], baterii: List[Baterie]) -> None:
    with st.expander("Baza de date panouri solare"):
        st.dataframe(df_panouri(panouri), use_container_width=True, hide_index=True)
        for p in panouri:
            st.markdown(
                f"**{p.nume}**: {p.descriere} Putere `{p.putere:.0f} W`, "
                f"eficienta `{p.eficienta:.1f}%`, suprafata `{p.suprafata:.2f} m2`, pret `{lei(p.pret)}`."
            )
    with st.expander("Baza de date baterii"):
        st.dataframe(df_baterii(baterii), use_container_width=True, hide_index=True)
        for b in baterii:
            st.markdown(
                f"**{b.nume}**: {b.descriere} Capacitate `{b.capacitate:.2f} kWh`, "
                f"tehnologie `{b.tehnologie}`, pret `{lei(b.pret)}`."
            )


def footer() -> None:
    st.markdown("""
    <div class="footer">
        Aplicatia foloseste valori orientative pentru dimensionarea initiala.
        Pentru implementare reala sunt necesare verificari suplimentare:
        orientarea acoperisului, umbrirea, randamentul invertorului,
        structura de prindere, protectiile electrice si avizele necesare.
    </div>
    """, unsafe_allow_html=True)


def main() -> None:
    st.markdown(CSS, unsafe_allow_html=True)
    panouri = panouri_db()
    baterii = baterii_db()
    buget, acoperis, ore_soare, montaj = sidebar_inputs()
    hero()
    consumatori = editor_consumatori()
    consum = total_consum(consumatori)

    st.markdown("### Rezultate privind consumul zilei și bugetul   ")
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        metric("Consum zilnic", kwh(consum), "calculat din consumatori")
    with c2:
        metric("Buget", lei(buget), "limita maxima introdusa")
    with c3:
        metric("Acoperis", m2(acoperis), "suprafata disponibila")
    with c4:
        metric("Ore soare", f"{ore_soare:.1f} h/zi", "echivalent productie")
    st.markdown("---")

    genereaza = st.button(
        "Genereaza recomandarea",
        type="primary",
        use_container_width=True,
    )

    if genereaza and consum <= 0:
        st.warning("Completeaza cel putin un consumator valid inainte de generarea recomandarii.")

    if genereaza and consum > 0:
        variante = calculeaza_variante(
            consumatori,
            panouri,
            baterii,
            buget,
            acoperis,
            ore_soare,
            montaj,
        )
        best = variante[0] if variante else None
        if best:
            afisare_best(best)
            analiza(best, acoperis)
            output_path = salveaza_recomandari(variante, consumatori)
            st.caption(f"Fisier generat: {output_path.name}")
        else:
            afisare_fara_solutie()
        afisare_variante(variante, 5)

    if not genereaza:
        st.info("Dupa ce completezi datele, apasa butonul pentru a genera recomandarea.")

    baze_date(panouri, baterii)
    footer()


if __name__ == "__main__":
    main()
