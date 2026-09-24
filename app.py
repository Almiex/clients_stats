# -*- coding: utf-8 -*-
"""
Дашборд «Приёмы» для Streamlit.
Запуск:  streamlit run dashboard_priem.py
Загрузите Excel-отчёт (как priem1.xlsx) через sidebar — дашборд построится автоматически.
"""
import io
import re
import altair as alt
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Дашборд приёмов", layout="wide")

# ---------- Логика ----------

def parse_report(file_bytes: bytes) -> tuple[pd.DataFrame, dict, str | None]:
    """Парсит отчёт: метаданные (период, клиника) + таблица услуг."""
    raw = pd.read_excel(io.BytesIO(file_bytes), sheet_name=0, header=None)

    meta, clinic = {}, None
    for v in raw[0].dropna():
        s = str(v)
        if s.startswith("С:"):
            meta["date_from"] = s.split(":", 1)[1].strip()
        elif s.startswith("ПО:"):
            meta["date_to"] = s.split(":", 1)[1].strip()
        elif s.startswith("Клиника:"):
            clinic = s.split(":", 1)[1].strip()

    hdr_mask = raw[0].astype(str).str.startswith("USLCODE")
    if not hdr_mask.any():
        raise ValueError("Не найдена строка заголовка (USLCODE). Проверьте формат файла.")
    hdr = raw.index[hdr_mask][0]

    df = raw.iloc[hdr + 1:].copy()
    df.columns = ["uslcode", "usluga", "specialnost", "kolichestvo", "cena", "summa"]
    df = df.dropna(subset=["uslcode"])
    df["usluga"] = df["usluga"].astype(str).str.strip()
    df["specialnost"] = df["specialnost"].astype(str).str.strip().str.upper()
    # чистка мусорных значений специальности из МИС (напр. "В Северном")
    df["specialnost"] = df["specialnost"].replace(
        to_replace=r"(?i).*(СЕВЕРНОМ|КЛИНИКА).*",
        value="ДРУГОЕ", regex=True)
    for c in ["kolichestvo", "cena", "summa"]:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)
    return df, meta, clinic


def classify(name: str) -> str:
    n = str(name).lower()
    if "первичн" in n:
        return "Первичный"
    if "повторн" in n:
        return "Повторный"
    return "Прочее"


def agg_by_specialty(df: pd.DataFrame) -> pd.DataFrame:
    piv = (
        df.pivot_table(index="specialnost", columns="tip",
                       values="kolichestvo", aggfunc="sum", fill_value=0)
        .reset_index()
    )
    for col in ["Первичный", "Повторный", "Прочее"]:
        if col not in piv.columns:
            piv[col] = 0
    piv["Всего"] = piv[["Первичный", "Повторный", "Прочее"]].sum(axis=1)
    piv["% первичных"] = piv["Первичный"] / piv["Всего"] * 100
    piv["% повторных"] = piv["Повторный"] / piv["Всего"] * 100
    piv["% прочих"] = piv["Прочее"] / piv["Всего"] * 100
    revenue = df.groupby("specialnost")["summa"].sum().rename("Выручка")
    piv = piv.merge(revenue, on="specialnost")
    # выручка по типам приёма
    rev_tip = (
        df.pivot_table(index="specialnost", columns="tip",
                       values="summa", aggfunc="sum", fill_value=0)
        .reset_index()
    )
    for col in ["Первичный", "Повторный", "Прочее"]:
        if col not in rev_tip.columns:
            rev_tip[col] = 0
    rev_tip = rev_tip.rename(columns={
        "Первичный": "Выручка первичных",
        "Повторный": "Выручка повторных",
        "Прочее": "Выручка прочих"})
    piv = piv.merge(rev_tip, on="specialnost")
    piv["% выр. первичных"] = piv["Выручка первичных"] / piv["Выручка"].replace(0, pd.NA) * 100
    piv["% выр. повторных"] = piv["Выручка повторных"] / piv["Выручка"].replace(0, pd.NA) * 100
    piv["% выр. прочих"] = piv["Выручка прочих"] / piv["Выручка"].replace(0, pd.NA) * 100
    piv[["% выр. первичных", "% выр. повторных", "% выр. прочих"]] =         piv[["% выр. первичных", "% выр. повторных", "% выр. прочих"]].fillna(0)
    # "БЕЗ СПЕЦИАЛЬНОСТИ" — в конец, чтобы не мешало основному списку
    piv["_bezz"] = (piv["specialnost"] == "ДРУГОЕ").astype(int)
    piv = piv.sort_values(["_bezz", "Выручка"], ascending=[True, False])
    return piv.drop(columns="_bezz").reset_index(drop=True)

# ---------- Интерфейс ----------

st.title("🏥 Дашборд: выполненные приёмы")

uploaded = st.file_uploader(
    "Загрузите отчёт (.xlsx)",
    type=["xlsx"],
    help="Ожидаемый формат: отчёт «Количество выполненных услуг "
         "на сумму по убыванию» с колонками USLCODE…кол-во*цена. "
         "Период и клиника подставятся из шапки файла.")

if uploaded is None:
    st.info("Загрузите Excel-отчёт выше, чтобы построить дашборд.")
    st.stop()

try:
    df, meta, clinic = parse_report(uploaded.getvalue())
except Exception as e:
    st.error(f"Не удалось разобрать файл: {e}")
    st.stop()

df["tip"] = df["usluga"].apply(classify)
spec = agg_by_specialty(df)

# ---------- Шапка: период и KPI ----------
period = f"{meta.get('date_from', '—')} — {meta.get('date_to', '—')}"
clinic_txt = clinic if clinic else "—"
st.subheader(f"📅 Период: {period}    🏥 Клиника: {clinic_txt}")

k1, k2, k3, k4, k5 = st.columns(5)
k1.metric("Позиций услуг", len(df))
k2.metric("Всего приёмов (шт.)", int(df["kolichestvo"].sum()))
k3.metric("Выручка", f"{df['summa'].sum():,.0f} ₽".replace(",", " "))
perv_all = df.loc[df["tip"] == "Первичный", "kolichestvo"].sum()
povt_all = df.loc[df["tip"] == "Повторный", "kolichestvo"].sum()
base = perv_all + povt_all
k4.metric("% первичных (всего)", f"{perv_all / base * 100:.1f} %" if base else "—")
k5.metric("% повторных (всего)", f"{povt_all / base * 100:.1f} %" if base else "—")

st.divider()

# ---------- Доля первичных/повторных по специальностям ----------
st.subheader("Доля первичных и повторных приёмов по специальностям")

SORTS = {
    "Выручка (больше → меньше)": ("Выручка", False),
    "Количество приёмов (больше → меньше)": ("Всего", False),
    "% первичных (больше → меньше)": ("% первичных", False),
    "% повторных (больше → меньше)": ("% повторных", False),
    "% прочих (больше → меньше)": ("% прочих", False),
    "Алфавит А → Я": ("specialnost", True),
    "Алфавит Я → А": ("specialnost", False),
}

c_slider, c_sort = st.columns([1, 1])
with c_slider:
    min_visits = st.slider("Скрыть специальности с количеством приёмов меньше:", 0, 200, 0, 10)
with c_sort:
    sort_sel = st.selectbox("Сортировка:", list(SORTS), index=0)

spec_view = spec[spec["Всего"] >= min_visits]
sort_col, sort_asc = SORTS[sort_sel]
spec_view = spec_view.sort_values(sort_col, ascending=sort_asc, kind="stable")
# "ДРУГОЕ" всегда в конце, при любой сортировке
spec_view = pd.concat([
    spec_view[spec_view["specialnost"] != "ДРУГОЕ"],
    spec_view[spec_view["specialnost"] == "ДРУГОЕ"],
]).reset_index(drop=True)

TIP_INFO = {
    "% первичных": "приёмы со словом «первичный» в названии",
    "% повторных": "приёмы со словом «повторный» в названии",
    "% прочих": "ЭВН, перевязки, онлайн-консультации и др. — не относятся к первичным/повторным",
}
st.caption("  •  ".join(f"**{k}** — {v}" for k, v in TIP_INFO.items()))

st.markdown("**По количеству приёмов, %**")
long = spec_view.melt(
    id_vars=["specialnost", "Выручка", "Всего"],
    value_vars=["% первичных", "% повторных", "% прочих"],
    var_name="Тип", value_name="Доля")

chart = (
    alt.Chart(long)
    .mark_bar()
    .encode(
        x=alt.X("specialnost:N", title=None, sort=None,
                axis=alt.Axis(labelAngle=-45, labelLimit=180, labelOverlap=False)),
        y=alt.Y("Доля:Q", stack="zero",
                axis=alt.Axis(title="Доля, %")),
        color=alt.Color(
            "Тип:N",
            scale=alt.Scale(
                domain=["% первичных", "% повторных", "% прочих"],
                range=["#2e86de", "#f39c12", "#b2bec3"]),
            legend=alt.Legend(title=None, orient="bottom")),
        tooltip=[
            alt.Tooltip("specialnost:N", title="Специальность"),
            alt.Tooltip("Тип:N", title="Тип"),
            alt.Tooltip("Доля:Q", title="Доля, %", format=".1f"),
            alt.Tooltip("Всего:Q", title="Приёмов всего", format=",.0f"),
            alt.Tooltip("Выручка:Q", title="Выручка специальности, ₽", format=",.0f"),
        ],
    )
    .properties(height=420)
)
st.altair_chart(chart, use_container_width=True)

st.markdown("**Выручка, ₽**")
rows = []
for _, r in spec_view.iterrows():
    for pct_col, rev_col in zip(
            ["% выр. первичных", "% выр. повторных", "% выр. прочих"],
            ["Выручка первичных", "Выручка повторных", "Выручка прочих"]):
        rows.append({
            "specialnost": r["specialnost"],
            "Тип": pct_col.replace("выр. ", ""),
            "Доля": r[pct_col],
            "Выручка сегмента": r[rev_col],
            "Выручка специальности": r["Выручка"],
            "Приёмов всего": r["Всего"],
        })
long_rev = pd.DataFrame(rows)

chart_rev = (
    alt.Chart(long_rev)
    .mark_bar()
    .encode(
        x=alt.X("specialnost:N", title=None, sort=None,
                axis=alt.Axis(labelAngle=-45, labelLimit=180, labelOverlap=False)),
        y=alt.Y("Выручка сегмента:Q", stack="zero", axis=alt.Axis(title="Выручка, ₽")),
        color=alt.Color(
            "Тип:N",
            scale=alt.Scale(
                domain=["% первичных", "% повторных", "% прочих"],
                range=["#2e86de", "#f39c12", "#b2bec3"]),
            legend=None),
        tooltip=[
            alt.Tooltip("specialnost:N", title="Специальность"),
            alt.Tooltip("Тип:N", title="Тип"),
            alt.Tooltip("Выручка сегмента:Q", title="Выручка сегмента, ₽", format=",.0f"),
            alt.Tooltip("Выручка специальности:Q", title="Выручка специальности, ₽", format=",.0f"),
            alt.Tooltip("Приёмов всего:Q", title="Приёмов всего", format=",.0f"),
        ],
    )
    .properties(height=420)
)
st.altair_chart(chart_rev, use_container_width=True)

left, right = st.columns([3, 2])
with left:
    st.markdown("**Детализация по специальностям**")
    show = spec_view[["specialnost", "Всего", "Первичный", "Повторный", "Прочее",
                      "% первичных", "% повторных",
                      "% выр. первичных", "% выр. повторных", "% выр. прочих",
                      "Выручка"]].copy()
    show.columns = ["Специальность", "Всего", "Первичные", "Повторные", "Прочие",
                    "% первичных", "% повторных",
                    "% выр. первичных", "% выр. повторных", "% выр. прочих",
                    "Выручка"]
    for c in ["% первичных", "% повторных", "% выр. первичных", "% выр. повторных", "% выр. прочих"]:
        show[c] = show[c].map("{:.1f} %".format)
    st.dataframe(show, use_container_width=True, hide_index=True)

with right:
    st.markdown("**Общая структура (по количеству)**")
    overall = df.groupby("tip")["kolichestvo"].sum()
    st.bar_chart(overall)
    st.dataframe(overall.rename("Кол-во").reset_index().rename(columns={"tip": "Тип"}),
                 use_container_width=True, hide_index=True)

st.divider()

# ---------- Динамика по цене / топ услуг ----------
st.subheader("Топ-15 услуг по выручке")
top = df.nlargest(15, "summa")[["usluga", "specialnost", "kolichestvo", "cena", "summa"]]
top.columns = ["Услуга", "Специальность", "Кол-во", "Цена", "Сумма"]
st.dataframe(top, use_container_width=True, hide_index=True)

st.divider()

# ---------- Топ первичных / повторных ----------
st.subheader("Топ первичных и повторных услуг")
t1, t2 = st.columns(2)

def top_table(tip_name):
    t = (df[df["tip"] == tip_name]
         .nlargest(10, "kolichestvo")[["usluga", "specialnost", "kolichestvo", "summa"]]
         .copy())
    t.columns = ["Услуга", "Специальность", "Кол-во", "Сумма"]
    return t

with t1:
    st.markdown("**Топ-10 первичных (по количеству)**")
    st.dataframe(top_table("Первичный"), use_container_width=True, hide_index=True)
with t2:
    st.markdown("**Топ-10 повторных (по количеству)**")
    st.dataframe(top_table("Повторный"), use_container_width=True, hide_index=True)

st.divider()

# ---------- Выгрузка ----------
csv = spec.to_csv(index=False).encode("utf-8-sig")
st.download_button("⬇️ Скачать сводку по специальностям (CSV)", csv,
                   file_name="svodka_specialnosti.csv", mime="text/csv")

with st.expander("Исходные данные отчёта"):
    st.dataframe(df, use_container_width=True, hide_index=True)
