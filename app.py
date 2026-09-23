# -*- coding: utf-8 -*-
"""
Дашборд «Приёмы» для Streamlit.
Запуск:  streamlit run dashboard_priem.py
Загрузите Excel-отчёт (как priem1.xlsx) через sidebar — дашборд построится автоматически.
"""
import io
import re
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
        value="БЕЗ СПЕЦИАЛЬНОСТИ", regex=True)
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
    # "БЕЗ СПЕЦИАЛЬНОСТИ" — в конец, чтобы не мешало основному списку
    piv["_bezz"] = (piv["specialnost"] == "БЕЗ СПЕЦИАЛЬНОСТИ").astype(int)
    piv = piv.sort_values(["_bezz", "Выручка"], ascending=[True, False])
    return piv.drop(columns="_bezz").reset_index(drop=True)

# ---------- Интерфейс ----------

st.title("🏥 Дашборд: выполненные приёмы")

uploaded = st.sidebar.file_uploader("Загрузите отчёт (.xlsx)", type=["xlsx"])
st.sidebar.caption("Ожидаемый формат: отчёт «Количество выполненных услуг "
                   "на сумму по убыванию» с колонками USLCODE…кол-во*цена.")

if uploaded is None:
    st.info("Загрузите Excel-отчёт в панели слева, чтобы построить дашборд.")
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
st.subheader(f"📅 Период: {period}")
if clinic:
    st.caption(f"Клиника: {clinic}")

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

min_visits = st.slider("Скрыть специальности с количеством приёмов меньше:", 0, 200, 0, 10)
spec_view = spec[spec["Всего"] >= min_visits]

chart_df = spec_view.set_index("specialnost")[["% первичных", "% повторных", "% прочих"]]
st.bar_chart(chart_df, stack=True, color=["#2e86de", "#f39c12", "#b2bec3"])

left, right = st.columns([3, 2])
with left:
    st.markdown("**Детализация по специальностям**")
    show = spec_view[["specialnost", "Всего", "Первичный", "Повторный", "Прочее",
                      "% первичных", "% повторных", "Выручка"]].copy()
    show.columns = ["Специальность", "Всего", "Первичные", "Повторные", "Прочие",
                    "% первичных", "% повторных", "Выручка"]
    for c in ["% первичных", "% повторных"]:
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
c1, c2 = st.columns(2)
with c1:
    st.subheader("Топ-15 услуг по выручке")
    top = df.nlargest(15, "summa")[["usluga", "specialnost", "kolichestvo", "cena", "summa"]]
    top.columns = ["Услуга", "Специальность", "Кол-во", "Цена", "Сумма"]
    st.dataframe(top, use_container_width=True, hide_index=True)
with c2:
    st.subheader("Распределение по специальностям (выручка)")
    st.bar_chart(spec.set_index("specialnost")["Выручка"])

st.divider()

# ---------- Выгрузка ----------
csv = spec.to_csv(index=False).encode("utf-8-sig")
st.download_button("⬇️ Скачать сводку по специальностям (CSV)", csv,
                   file_name="svodka_specialnosti.csv", mime="text/csv")

with st.expander("Исходные данные отчёта"):
    st.dataframe(df, use_container_width=True, hide_index=True)
