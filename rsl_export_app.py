# -*- coding: utf-8 -*-
"""RSL Export — schlanke Web-App (Streamlit Community Cloud / iPhone).

Lädt live die S&P-500-Kurse via yfinance, berechnet RSL (Tagesschluss,
SMA 130/200 HT) wie RSL_Update.py, filtert Kursartefakte (Split/Spin-off-
Sprünge) heraus und zeigt das Top-20-Ranking mit Excel-Download.
Vollständig eigenständig — keine lokalen Pfade, daher cloud-deploybar.
"""
import datetime, io, warnings
warnings.filterwarnings("ignore")
import pandas as pd
import numpy as np
import streamlit as st
import yfinance as yf
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

JUMP_MAX = 0.40  # > 40 % Ein-Tages-Sprung in letzten 20 HT => Artefakt

st.set_page_config(page_title="RSL Export", page_icon="📈", layout="centered")

st.markdown(
    "<h1 style='margin-bottom:0'>📈 RSL Export</h1>"
    "<p style='color:#6e7781;margin-top:4px'>B5-Buffer · Top 20 nach RSL (26W) · S&amp;P 500 · Live via yfinance</p>",
    unsafe_allow_html=True,
)


@st.cache_data(show_spinner=False)
def lade_universum() -> pd.DataFrame:
    """S&P-500-Universum + Sektoren aus mitgelieferter CSV."""
    return pd.read_csv("sp500_universe.csv")


@st.cache_data(ttl=60 * 30, show_spinner=True)
def berechne_ranking():
    """Lädt Kurse, berechnet RSL_26W/40W, filtert Artefakte. Cache 30 Min."""
    uni = lade_universum()
    sek = {r.Symbol: (r.Sektor, r.Untersektor) for r in uni.itertuples()}
    symbole = [s.replace(".", "-") for s in uni["Symbol"].tolist()]

    try:
        fx = float(
            yf.download("EURUSD=X", period="5d", progress=False, auto_adjust=True)
            ["Close"].dropna().iloc[-1]
        )
    except Exception:
        fx = 1.17

    closes = {}
    BATCH = 120
    for i in range(0, len(symbole), BATCH):
        part = symbole[i:i + BATCH]
        raw = yf.download(part, period="350d", interval="1d",
                          auto_adjust=True, progress=False, threads=True)
        cl = raw["Close"]
        if isinstance(cl, pd.Series):
            cl = cl.to_frame(part[0])
        for s in cl.columns:
            closes[s] = cl[s]
    px = pd.DataFrame(closes).sort_index()

    rows, artefakte = [], []
    for s in px.columns:
        ser = px[s].dropna()
        if len(ser) < 130:
            continue
        if float(ser.iloc[-20:].pct_change().abs().max()) > JUMP_MAX:
            artefakte.append(s.replace("-", "."))
            continue
        price = ser.iloc[-1]
        sma40 = ser.rolling(200, min_periods=200).mean().iloc[-1]
        sma26 = ser.rolling(130, min_periods=130).mean().iloc[-1]
        if not np.isfinite(sma26):
            continue
        rsl40 = price / sma40 if np.isfinite(sma40) else np.nan
        rsl26 = price / sma26
        orig = s.replace("-", ".")
        sk, sub = sek.get(s, sek.get(orig, ("", "")))
        rows.append(dict(
            Symbol=orig, Sektor=sk, Branche=sub,
            Preis_USD=round(float(price), 2),
            Preis_EUR=round(float(price) / fx, 2),
            RSL_26W=round(float(rsl26), 4),
            RSL_40W=round(float(rsl40), 4) if np.isfinite(rsl40) else np.nan,
        ))
    df = pd.DataFrame(rows)
    df = (df[(df["RSL_26W"] > 0.01) & (df["RSL_26W"] < 5.0)]
          .sort_values("RSL_26W", ascending=False).reset_index(drop=True))
    df.insert(0, "Rang", df.index + 1)
    return df, fx, artefakte


def baue_excel(df: pd.DataFrame, fx: float) -> bytes:
    heute = datetime.date.today().strftime("%Y-%m-%d")
    HEAD = PatternFill("solid", fgColor="1F2A37")
    HF = Font(bold=True, color="FFFFFF", size=11)
    GR = Font(color="1A7F37", bold=True); RD = Font(color="B42318")
    THIN = Side(style="thin", color="D0D7DE")
    BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
    CEN = Alignment(horizontal="center")

    def sheet(ws, data, title):
        ws.sheet_view.showGridLines = False
        ws["A1"] = title
        ws["A1"].font = Font(bold=True, size=13, color="1F2A37")
        ws["A2"] = (f"Datenstand {heute} · RSL_26W = Kurs / SMA(130 HT) · "
                    f"EUR/USD {fx:.4f} · yfinance · Artefakt-Filter >{int(JUMP_MAX*100)}%/Tag")
        ws["A2"].font = Font(size=9, color="6E7781")
        for j, h in enumerate(["Rang", "Ticker", "Branche", "Sektor",
                               "RSL 26W", "Kurs USD", "Kurs EUR"], 1):
            c = ws.cell(4, j, h); c.fill = HEAD; c.font = HF
            c.alignment = CEN; c.border = BORDER
        for k, row in enumerate(data.itertuples(), start=5):
            vals = [row.Rang, row.Symbol, row.Branche, row.Sektor,
                    row.RSL_26W, row.Preis_USD, row.Preis_EUR]
            for j, v in enumerate(vals, 1):
                c = ws.cell(k, j, v); c.border = BORDER
                if j in (1, 5):
                    c.alignment = CEN
                if j == 5:
                    c.font = GR if (isinstance(v, (int, float)) and v > 1) else RD
                    c.number_format = "0.000"
                if j in (6, 7):
                    c.number_format = '#,##0.00'
        for j, w in enumerate([6, 9, 30, 22, 10, 12, 12], 1):
            ws.column_dimensions[chr(64 + j)].width = w
        ws.freeze_panes = "A5"

    wb = Workbook()
    ws1 = wb.active; ws1.title = "Top 20"
    sheet(ws1, df.head(20), "RSL Top 20 — B5-Buffer (SMA 26W)")
    sheet(wb.create_sheet("Gesamtranking"), df, "RSL Gesamtranking S&P 500 (SMA 26W)")
    buf = io.BytesIO(); wb.save(buf); return buf.getvalue()


# ── UI ───────────────────────────────────────────────────────────────────────
if st.button("🔄 Daten jetzt aktualisieren", use_container_width=True, type="primary"):
    berechne_ranking.clear()

with st.spinner("Lade S&P-500-Kurse und berechne RSL …"):
    df, fx, artefakte = berechne_ranking()

heute = datetime.date.today().strftime("%d.%m.%Y")
st.caption(f"Datenstand {heute} · EUR/USD {fx:.4f} · {len(df)} Titel · "
           f"RSL_26W = Kurs / SMA(130 Handelstage)")

top = df.head(20)[["Rang", "Symbol", "Sektor", "RSL_26W", "Preis_USD", "Preis_EUR"]].copy()
top.columns = ["Rang", "Ticker", "Sektor", "RSL 26W", "Kurs USD", "Kurs EUR"]
st.subheader("Top 20")
st.dataframe(top, hide_index=True, use_container_width=True,
             column_config={
                 "RSL 26W": st.column_config.NumberColumn(format="%.3f"),
                 "Kurs USD": st.column_config.NumberColumn(format="%.2f"),
                 "Kurs EUR": st.column_config.NumberColumn(format="%.2f"),
             })

c1, c2 = st.columns(2)
with c1:
    st.download_button(
        "⬇️ Top 20 + Ranking als Excel", baue_excel(df, fx),
        file_name=f"RSL_Top20_{datetime.date.today()}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True)
with c2:
    st.download_button(
        "⬇️ Gesamtranking als CSV",
        df.drop(columns=["Preis_EUR"]).to_csv(index=False).encode("utf-8"),
        file_name=f"rsl_export_{datetime.date.today()}.csv",
        mime="text/csv", use_container_width=True)

with st.expander(f"Gesamtranking ({len(df)} Titel) anzeigen"):
    voll = df[["Rang", "Symbol", "Sektor", "Branche", "RSL_26W", "Preis_USD", "Preis_EUR"]]
    st.dataframe(voll, hide_index=True, use_container_width=True)

if artefakte:
    st.caption("Aussortierte Kursartefakte (Split/Spin-off-Sprung > 40 %/Tag): "
               + ", ".join(artefakte))

st.caption("Hinweis: Stichtagsbetrachtung ohne Depotkenntnis. Buffer-/Sektor-Regeln "
           "und Schock-Filter siehe Strategie-Doku. Keine Anlageberatung.")
