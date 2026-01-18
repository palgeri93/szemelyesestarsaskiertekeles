import io
import re
import zipfile
from typing import List, Tuple, Optional

import pandas as pd
import streamlit as st
import plotly.express as px

MAX_POINTS = 70
APP_TITLE = "Személyes és társas kompetenciák mérése"
FOOTER_TEXT = "Készítette: Sulyok István Ált. Iskola és AMI 2026. Pálfi Gergő"


def safe_filename(s: str) -> str:
    s = str(s).strip()
    s = re.sub(r"\s+", "_", s)
    s = re.sub(r"[^0-9A-Za-zÁÉÍÓÖŐÚÜŰáéíóöőúüű_\-\.]", "", s)
    return s or "nev_nelkul"


def build_color_map(period_order: List[str]) -> dict:
    palette = ["#1f77b4", "#ff7f0e", "#2ca02c", "#d62728", "#9467bd", "#8c564b"]
    return {p: palette[i % len(palette)] for i, p in enumerate(period_order)}


def read_two_periods(xlsx_bytes: bytes) -> Tuple[pd.DataFrame, Optional[pd.DataFrame], List[str], List[str]]:
    xls = pd.ExcelFile(io.BytesIO(xlsx_bytes))
    sheets = xls.sheet_names
    if len(sheets) < 1:
        raise ValueError("Az Excel fájlban nincs munkalap.")

    period_labels = [sheets[0]]
    df1 = pd.read_excel(xls, sheet_name=sheets[0])

    df2 = None
    if len(sheets) >= 2:
        period_labels.append(sheets[1])
        df2 = pd.read_excel(xls, sheet_name=sheets[1])

    if df1.shape[1] < 3:
        raise ValueError("Az első munkalapon legalább 3 oszlop kell: Név, Osztály, és mérési területek.")

    name_col = df1.columns[0]
    class_col = df1.columns[1]
    areas = list(df1.columns[2:])

    df1 = df1[[name_col, class_col] + areas].copy()
    df1.rename(columns={name_col: "Név", class_col: "Osztály"}, inplace=True)

    if df2 is not None:
        if df2.shape[1] < 3:
            raise ValueError("A második munkalapon legalább 3 oszlop kell: Név, Osztály, és mérési területek.")
        name_col2 = df2.columns[0]
        class_col2 = df2.columns[1]
        areas2 = list(df2.columns[2:])

        common_areas = [a for a in areas if a in areas2]
        if not common_areas:
            raise ValueError("A két munkalap terület-oszlopai nem egyeznek (nincs közös terület).")

        areas = common_areas
        df2 = df2[[name_col2, class_col2] + areas].copy()
        df2.rename(columns={name_col2: "Név", class_col2: "Osztály"}, inplace=True)

    return df1, df2, areas, period_labels


def to_long_percent(df: pd.DataFrame, period: str, areas: List[str]) -> pd.DataFrame:
    out = df.copy()
    for a in areas:
        out[a] = pd.to_numeric(out[a], errors="coerce")
    long = out.melt(
        id_vars=["Név", "Osztály"],
        value_vars=areas,
        var_name="Terület",
        value_name="Pont",
    )
    long["Időszak"] = period
    long["Százalék"] = (long["Pont"] / MAX_POINTS) * 100
    return long


def avg_view(long_df: pd.DataFrame, level: str) -> pd.DataFrame:
    if level == "Összes":
        g = long_df.groupby(["Időszak", "Terület"], as_index=False)["Százalék"].mean()
    else:
        g = long_df.groupby(["Időszak", "Osztály", "Terület"], as_index=False)["Százalék"].mean()
    return g


def make_bar_figure(data, area_order, period_order, title: str):
    color_map = build_color_map(period_order)
    fig = px.bar(
        data,
        x="Terület",
        y="Százalék",
        color="Időszak",
        category_orders={"Terület": area_order, "Időszak": period_order},
        color_discrete_map=color_map,
        barmode="group",
        range_y=[0, 110],
        labels={"Százalék": "Százalék (%)"},
        title=title,
    )
    fig.update_traces(texttemplate="%{y:.1f}%", textposition="outside", cliponaxis=False)
    fig.update_layout(
        legend_title_text="Időszak",
        xaxis_title="Mérési terület",
        xaxis_tickfont_size=16,
        yaxis_tickfont_size=14,
        title_font_size=20,
        margin=dict(t=90, b=120),
    )
    return fig


def make_radar_figure(data, area_order, period_order, title: str):
    color_map = build_color_map(period_order)
    fig = px.line_polar(
        data,
        r="Százalék",
        theta="Terület",
        color="Időszak",
        category_orders={"Terület": area_order, "Időszak": period_order},
        color_discrete_map=color_map,
        line_close=True,
        range_r=[0, 100],
        labels={"Százalék": "Százalék (%)"},
        title=title,
    )
    fig.update_layout(
        legend_title_text="Időszak",
        title_font_size=20,
        margin=dict(t=90, b=120),
    )
    return fig


def wrap_html(title: str, footer: str, plotly_html_div: str) -> str:
    # Teljes HTML oldal. Innen a böngészőből: Ctrl+P → Mentés PDF-be.
    return f"""<!doctype html>
<html lang="hu">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>{title}</title>
  <style>
    body {{ font-family: Arial, sans-serif; margin: 24px; }}
    h1 {{ font-size: 20px; margin: 0 0 16px 0; }}
    .footer {{ margin-top: 18px; font-size: 12px; color: #666; }}
    @media print {{
      body {{ margin: 10mm; }}
      .footer {{ position: fixed; bottom: 8mm; left: 10mm; right: 10mm; }}
    }}
  </style>
</head>
<body>
  <h1>{title}</h1>
  {plotly_html_div}
  <div class="footer">{footer}</div>
</body>
</html>
"""


def figures_to_zip_html(fig_items) -> bytes:
    """
    fig_items: [(filename.html, fig, title), ...]
    Kaleido/Chrome nélkül működik: Plotly HTML export.
    """
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for fname, fig, title in fig_items:
            # include_plotlyjs='cdn' kisebb fájl, de internet kell a megnyitáshoz.
            # Ha offline kell, állítsd 'include_plotlyjs=True'-re (nagyobb ZIP).
            div = fig.to_html(full_html=False, include_plotlyjs="cdn")
            html = wrap_html(title=title, footer=FOOTER_TEXT, plotly_html_div=div)
            zf.writestr(fname, html.encode("utf-8"))

        # kis README
        zf.writestr(
            "README.txt",
            (
                "Megnyitás: nyisd meg a HTML fájlokat böngészőben.\n"
                "PDF: Ctrl+P (Nyomtatás) -> Mentés PDF-be.\n"
                "Megjegyzés: ha nincs internet, és nem tölt be a grafikon, szólj, átállítom offline Plotly JS-re.\n"
            ).encode("utf-8"),
        )
    return buf.getvalue()


def main():
    st.set_page_config(page_title=APP_TITLE, layout="wide")
    st.title(APP_TITLE)

    uploaded = st.file_uploader("Excel fájl feltöltése (.xlsx)", type=["xlsx"])
    if not uploaded:
        st.stop()

    try:
        df1, df2, areas, periods = read_two_periods(uploaded.getvalue())
    except Exception as e:
        st.error(f"Hiba a beolvasásnál: {e}")
        st.stop()

    long1 = to_long_percent(df1, periods[0], areas)
    long_all = long1.copy()
    if df2 is not None:
        long2 = to_long_percent(df2, periods[1], areas)
        long_all = pd.concat([long1, long2], ignore_index=True)

    area_order = areas[:]
    period_order = periods[:]

    left, _ = st.columns([1, 2])
    with left:
        view_mode = st.radio("Nézet", ["Egy tanuló", "Osztályátlag", "Összes átlag"], index=0)
        classes = sorted([c for c in long_all["Osztály"].dropna().unique().tolist()])
        selected_class = st.selectbox("Osztály szűrés", ["(mind)"] + classes, index=0)
        show_table = st.checkbox("Táblázat mutatása (%)", value=True)

    filtered = long_all.copy()
    if selected_class != "(mind)":
        filtered = filtered[filtered["Osztály"] == selected_class]

    if view_mode == "Egy tanuló":
        names = sorted(filtered["Név"].dropna().unique().tolist())
        selected_name = st.selectbox("Tanuló", names)
        data = filtered[filtered["Név"] == selected_name].copy()
        chart_title = f"Személyes és társas kompetencia : {selected_name}"
    elif view_mode == "Osztályátlag":
        data = avg_view(filtered, "Osztály")
        if selected_class == "(mind)":
            cls = st.selectbox("Melyik osztály átlagát nézzük?", sorted(data["Osztály"].unique().tolist()))
            data = data[data["Osztály"] == cls].copy()
            chart_title = f"Személyes és társas kompetencia : {cls} átlag"
        else:
            chart_title = f"Személyes és társas kompetencia : {selected_class} átlag"
    else:
        data = avg_view(filtered, "Összes")
        chart_title = "Személyes és társas kompetencia : Átlag"

    st.subheader("Területenkénti összehasonlítás (%)")
    st.plotly_chart(make_bar_figure(data, area_order, period_order, chart_title), use_container_width=True)
    st.caption(FOOTER_TEXT)

    st.subheader("Radar diagram (%)")
    st.plotly_chart(make_radar_figure(data, area_order, period_order, chart_title), use_container_width=True)
    st.caption(FOOTER_TEXT)

    if show_table:
        st.subheader("Százalékos táblázat (%)")
        pivot = (
            data.pivot_table(index="Terület", columns="Időszak", values="Százalék", aggfunc="mean")
            .reindex(area_order)
            .reindex(columns=period_order)
        )
        st.dataframe(pivot.style.format("{:.1f}%"), use_container_width=True)

    # -------- ZIP HTML --------
    st.divider()
    st.subheader("Tömeges letöltés (ZIP) – HTML riportok")
    st.write("Letöltés után nyisd meg a HTML fájlokat böngészőben. PDF-hez: Ctrl+P → Mentés PDF-be.")

    export_scope = st.radio(
        "Mit csomagoljunk a ZIP-be?",
        ["Minden tanuló (aktuális osztály-szűréssel)", "Minden tanuló (minden osztály)"],
        index=0,
        key="zip_scope",
    )

    if export_scope == "Minden tanuló (minden osztály)":
        export_df = long_all
        zip_name = "diagramok_html_minden_osztaly.zip"
    else:
        export_df = filtered
        zip_name = "diagramok_html_szurt.zip"

    st.session_state.setdefault("zip_ready", False)
    st.session_state.setdefault("zip_bytes", b"")
    st.session_state.setdefault("zip_filename", zip_name)

    if st.session_state["zip_filename"] != zip_name:
        st.session_state["zip_filename"] = zip_name
        st.session_state["zip_ready"] = False
        st.session_state["zip_bytes"] = b""

    if st.button("ZIP elkészítése", type="primary"):
        st.session_state["zip_ready"] = False
        st.session_state["zip_bytes"] = b""

        export_names = (
            export_df[["Név", "Osztály"]]
            .dropna()
            .drop_duplicates()
            .sort_values(["Osztály", "Név"])
            .to_records(index=False)
        )

        total = len(export_names)
        progress = st.progress(0)
        status = st.empty()

        try:
            items = []
            for i, (nev, osztaly) in enumerate(export_names, start=1):
                status.write(f"Feldolgozás: {osztaly} – {nev} ({i}/{total})")
                one = export_df[(export_df["Név"] == nev) & (export_df["Osztály"] == osztaly)].copy()
                if one.empty:
                    progress.progress(min(i / max(total, 1), 1.0))
                    continue

                title = f"Személyes és társas kompetencia : {nev}"
                bar = make_bar_figure(one, area_order, period_order, title)
                radar = make_radar_figure(one, area_order, period_order, title)

                base = f"{safe_filename(osztaly)}__{safe_filename(nev)}"
                items.append((f"{base}__oszlopdiagram.html", bar, title))
                items.append((f"{base}__radar.html", radar, title))

                progress.progress(min(i / max(total, 1), 1.0))

            status.write("HTML fájlok és ZIP csomagolás…")
            zip_bytes = figures_to_zip_html(items)

            st.session_state["zip_bytes"] = zip_bytes
            st.session_state["zip_ready"] = True
            st.session_state["zip_filename"] = zip_name
            st.session_state["zip_size"] = len(zip_bytes)
            st.rerun()

        except Exception as e:
            st.session_state["zip_ready"] = False
            st.session_state["zip_bytes"] = b""
            st.error(f"Nem sikerült a ZIP készítés.\n\nRészletek: {e}")

    if st.session_state.get("zip_ready") and st.session_state.get("zip_bytes"):
        size_mb = st.session_state.get("zip_size", len(st.session_state["zip_bytes"])) / (1024 * 1024)
        st.success(f"ZIP elkészült! Méret: {size_mb:.2f} MB")
        st.download_button(
            "Letöltés: HTML riportok ZIP",
            data=st.session_state["zip_bytes"],
            file_name=st.session_state.get("zip_filename", "diagramok_html.zip"),
            mime="application/zip",
        )


if __name__ == "__main__":
    main()
