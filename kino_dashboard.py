import re
import os
import glob
import math
from itertools import combinations

import pandas as pd
import plotly.graph_objects as go
import streamlit as st


st.set_page_config(
    page_title="KinoOlega: статистика и динамика",
    page_icon="🎬",
    layout="wide"
)


# ============================================================
# ПАСПОРТ ВОПРОСОВ (Лето#2026)
# (раунд, тема, тип, номиналы)
# ============================================================
THEMES_META = [
    (1, "Хокку ✏️", "text", [100, 200, 300, 400, 500]),
    (1, "Намек на страну 📸", "image", [100, 200, 300, 400, 500]),
    (1, "СССР и Россия ✏️", "text", [100, 200, 300, 400, 500]),
    (1, "По губам, по глазам, по голосу 📸", "image", [100, 200, 300, 400, 500]),
    (2, "Продавцы 📸", "image", [200, 400, 600, 800, 1000]),
    (2, "... против ... ✏️", "text", [200, 400, 600, 800, 1000]),
    (2, "Дело было в баре 📸", "image", [150, 300, 450, 600, 750]),
    (2, "Геометрические ответы ✏️", "text", [200, 400, 600, 800, 1000]),
    (3, "Ответы на букву Д ✏️", "text", [300, 600, 900, 1200, 1500]),
    (3, "Синонимы 📸", "image", [300, 600, 900, 1200, 1500]),
    (3, "Поп... ✏️", "text", [300, 600, 900, 1200, 1500]),
    (3, "Мне нужна твоя одежда 📸", "image", [200, 400, 600, 800, 1000]),
    ("ФИНАЛ", 'Про "бабушку"', "text", [1500]),
    ("ФИНАЛ", "Про ловкость рук", "text", [2000]),
    ("ФИНАЛ", "Про доминирование", "text", [2500]),
]

SPECIAL_QUESTIONS = {61, 62, 63}
SPECIAL_WEIGHT_MULTIPLIER = 3.0


def build_question_passport() -> pd.DataFrame:
    rows = []
    q = 0
    for rnd, theme, qtype, prices in THEMES_META:
        for price in prices:
            q += 1
            rows.append({
                "q_num": q,
                "q_col": f"Q{q}",
                "round": rnd,
                "theme": theme,
                "type": qtype,
                "type_ru": "Текстовый ✏️" if qtype == "text" else "Картинный 📸",
                "price": price,
                "is_final": rnd == "ФИНАЛ",
            })
    return pd.DataFrame(rows)


QUESTION_PASSPORT = build_question_passport()


# ============================================================
# ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
# ============================================================
def get_q_columns(df: pd.DataFrame) -> list:
    pattern = re.compile(r"^\s*Q\s*(\d+)\s*$", re.IGNORECASE)
    found = []
    for col in df.columns:
        m = pattern.match(str(col))
        if m:
            found.append((int(m.group(1)), col))
    found.sort(key=lambda x: x[0])
    return [c for _, c in found]


def pct(part, whole):
    return round(100.0 * part / whole, 1) if whole else 0.0


def streak_stats(values):
    streaks = []
    cur = 0
    for v in values:
        if v > 0:
            cur += 1
        else:
            if cur > 0:
                streaks.append(cur)
            cur = 0
    if cur > 0:
        streaks.append(cur)
    return {
        "Лучшая серия": max(streaks) if streaks else 0,
        "Всего серий": len(streaks),
        "Серий 3+": sum(1 for s in streaks if s >= 3),
        "Серий 5+": sum(1 for s in streaks if s >= 5),
    }


def recalc_for_group(group_df, passport_df, q_cols):
    """Пересчитывает взвешенные очки для группы игроков, как будто они играли в одной игре."""
    total_players = len(group_df)
    if total_players == 0:
        return pd.DataFrame(), {}

    q_numeric = group_df[q_cols].astype(float)
    correct_counts = (q_numeric > 0).sum(axis=0)

    base_weights = {}
    for q_col in q_cols:
        m = re.search(r'\d+', q_col)
        q_num = int(m.group()) if m else 0
        base = 1 - (correct_counts[q_col] / total_players) if total_players > 0 else 0
        if q_num in SPECIAL_QUESTIONS:
            base *= SPECIAL_WEIGHT_MULTIPLIER
        base_weights[q_col] = base

    nominals = passport_df.set_index('q_col')['price'].to_dict()

    results = []
    for _, row in group_df.iterrows():
        total_pts = 0.0
        total_w_pts = 0.0
        correct_count = 0
        for q_col in q_cols:
            points = float(row[q_col])
            if points > 0:
                correct_count += 1
                total_pts += points
                nominal = nominals.get(q_col, points)
                ratio = points / nominal if nominal > 0 else 1.0
                total_w_pts += base_weights[q_col] * ratio
        f_ochki = (total_pts + 3 * total_w_pts * 1000) / 4
        results.append({
            'Имя игрока': row['Имя игрока'],
            'Игра': row.get('_game_label', ''),
            'Очки': round(total_pts, 1),
            'Ответы': correct_count,
            'wОчки': round(total_w_pts, 2),
            'fОчки': round(f_ochki, 2),
        })
    return pd.DataFrame(results), base_weights


def find_local_excel_files() -> list:
    """Ищет xlsx/xlsm в папке скрипта и в подпапке data."""
    base_dir = os.path.dirname(os.path.abspath(__file__))
    data_dir = os.path.join(base_dir, "data")
    found = []
    for folder in (base_dir, data_dir):
        if os.path.isdir(folder):
            found.extend(glob.glob(os.path.join(folder, "*.xlsx")))
            found.extend(glob.glob(os.path.join(folder, "*.xlsm")))
    found = [p for p in found if not os.path.basename(p).startswith("~$")]
    return sorted(set(found))


# ============================================================
# ЗАГОЛОВОК И ИСТОЧНИК ДАННЫХ
# ============================================================
st.title("🎬 KinoOlega: статистика и динамика")
st.caption("Обычные очки из Q1...Qn. Фильтры по типу/темам, статистика, виртуальные игры.")

local_files = find_local_excel_files()
local_names = [os.path.basename(p) for p in local_files]

chosen_name = None
uploaded_file = None

with st.sidebar:
    st.header("📁 Источник данных")
    if local_files:
        chosen_name = st.selectbox("Excel-файл из папки проекта", local_names)
        st.caption(
            "Файл найден рядом со скриптом (или в папке `data`) и подхватывается "
            "автоматически. Обновишь файл — все увидят новые игры."
        )
    else:
        uploaded_file = st.file_uploader(
            "Загрузите Excel-файл KinoOlega",
            type=["xlsx", "xlsm"]
        )
        st.caption("Локальные xlsx не найдены, поэтому включён ручной режим.")

xls = None

if chosen_name:
    path = local_files[local_names.index(chosen_name)]
    st.success(f"📂 Данные загружены автоматически: **{chosen_name}**")
    try:
        xls = pd.ExcelFile(path)
    except Exception as e:
        st.error(f"Не удалось открыть файл {chosen_name}: {e}")
        st.stop()
elif uploaded_file is not None:
    uploaded_file.seek(0)
    try:
        xls = pd.ExcelFile(uploaded_file)
    except Exception as e:
        st.error(f"Не удалось открыть Excel-файл: {e}")
        st.stop()
else:
    st.info("Положите xlsx рядом со скриптом (или в папку data) либо загрузите его вручную.")
    st.stop()

sheet_names = xls.sheet_names
default_sheet = "Лист2" if "Лист2" in sheet_names else sheet_names[0]

with st.sidebar:
    sheet_name = st.selectbox("Лист Excel", sheet_names, index=sheet_names.index(default_sheet))

try:
    df = pd.read_excel(xls, sheet_name=sheet_name)
except Exception as e:
    st.error(f"Не удалось прочитать лист '{sheet_name}': {e}")
    st.stop()

if df.empty:
    st.error("Выбранный лист пустой.")
    st.stop()

q_cols = get_q_columns(df)

if "Имя игрока" not in df.columns or not q_cols:
    st.error("На листе нет колонки 'Имя игрока' и/или колонок Q1...Qn.")
    st.stop()

# Подготовка данных
df = df.dropna(subset=["Имя игрока"]).copy()
df["Имя игрока"] = df["Имя игрока"].astype(str).str.strip()

if "Дата игры" in df.columns and "Игра" in df.columns:
    df["_game_label"] = df["Дата игры"].astype(str) + " · " + df["Игра"].astype(str)
elif "Игра" in df.columns:
    df["_game_label"] = df["Игра"].astype(str)
else:
    df["_game_label"] = "Игра " + (df.index + 1).astype(str)

for col in q_cols:
    df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

for col in ["Очки", "Ответы"]:
    if col in df.columns:
        df[col] = pd.to_numeric(df[col], errors="coerce")

game_options = df["_game_label"].unique().tolist()

# ============================================================
# ГЛОБАЛЬНАЯ СЛОЖНОСТЬ ВОПРОСОВ
# ============================================================
total_rows_all = len(df)
global_q_stats = {}
for q_col in q_cols:
    correct_all = int((df[q_col] > 0).sum())
    rate = correct_all / total_rows_all if total_rows_all else 0.0
    global_q_stats[q_col] = {"correct": correct_all, "total": total_rows_all, "rate": rate}

# ============================================================
# ФИЛЬТРЫ
# ============================================================
with st.sidebar:
    st.divider()
    st.header("🎯 Срез вопросов")
    type_filter = st.radio("Тип вопроса", ["Все", "Текстовые ✏️", "Картинные 📸"], index=0)
    available_themes = QUESTION_PASSPORT["theme"].unique().tolist()
    selected_themes = st.multiselect("Темы", available_themes, default=available_themes)

    st.divider()
    st.header("📈 Статистика")
    rare_fail_pct = st.slider(
        "Редкий ответ: если НЕ ответили ≥ %",
        min_value=50, max_value=100, value=90, step=5,
        help="Если 90% игроков не ответили, а этот игрок ответил — ответ редкий."
    )

    st.divider()
    st.header("📊 График")
    chart_type = st.radio(
        "Тип отображения",
        ["Накопленные очки", "Очки за каждый вопрос", "Итоги по игрокам"],
        index=0
    )
    show_markers = st.checkbox("Показывать точки на линиях", value=True)
    bar_mode = st.selectbox("Режим столбиков", ["group", "overlay", "relative"], index=0)
    top_n = st.slider("Топ-N игроков (если никто не выбран)", 1, 20, 3)

    st.divider()
    st.header("🎮 Игры и игроки")
    selected_games = st.multiselect("Игры", game_options, default=game_options[:1])

if selected_games:
    filtered_df = df[df["_game_label"].isin(selected_games)].copy()
else:
    filtered_df = df.copy()

if filtered_df.empty:
    st.warning("Нет данных для выбранных игр.")
    st.stop()

# Фильтр вопросов по типу и теме
q_passport_filtered = QUESTION_PASSPORT.copy()
if type_filter == "Текстовые ✏️":
    q_passport_filtered = q_passport_filtered[q_passport_filtered["type"] == "text"]
elif type_filter == "Картинные 📸":
    q_passport_filtered = q_passport_filtered[q_passport_filtered["type"] == "image"]
if selected_themes:
    q_passport_filtered = q_passport_filtered[q_passport_filtered["theme"].isin(selected_themes)]

filtered_q_cols = q_passport_filtered["q_col"].tolist()
filtered_q_numbers = q_passport_filtered["q_num"].tolist()

if not filtered_q_cols:
    st.warning("По выбранным фильтрам вопросов не осталось.")
    st.stop()

q_meta_lookup = {
    row["q_num"]: {"theme": row["theme"], "price": row["price"], "type_ru": row["type_ru"]}
    for _, row in q_passport_filtered.iterrows()
}

# Выбор игроков
tmp_sums = filtered_df[filtered_q_cols].sum(axis=1)
top_players = (
    filtered_df.assign(_sum=tmp_sums)
    .sort_values("_sum", ascending=False)["Имя игрока"]
    .drop_duplicates().head(top_n).tolist()
)
player_options = filtered_df["Имя игрока"].unique().tolist()

with st.sidebar:
    selected_players_raw = st.multiselect("Игроки", player_options)

selected_players = [p for p in selected_players_raw if p in player_options]
if not selected_players:
    selected_players = top_players
    st.caption(f"Игроки не выбраны. Показаны топ-{len(selected_players)}.")

plot_df = filtered_df[filtered_df["Имя игрока"].isin(selected_players)].copy()
if plot_df.empty:
    st.warning("Нет данных для выбранных игроков.")
    st.stop()

plot_df = plot_df.sort_values(["_game_label", "Имя игрока"]).copy()
multiple_games = plot_df["_game_label"].nunique() > 1


def make_row_label(row):
    if multiple_games:
        return f"{row['Имя игрока']} · {row['_game_label']}"
    return str(row["Имя игрока"])


# ============================================================
# ДЛИННЫЙ ФОРМАТ ДЛЯ СТАТИСТИКИ
# ============================================================
long_df = plot_df.melt(
    id_vars=["Имя игрока", "_game_label"],
    value_vars=filtered_q_cols,
    var_name="q_col", value_name="points"
)
long_df["points"] = pd.to_numeric(long_df["points"], errors="coerce").fillna(0)
long_df = long_df.merge(
    QUESTION_PASSPORT[["q_col", "theme", "type_ru", "round", "price"]],
    on="q_col", how="left"
)
long_df["is_correct"] = long_df["points"] > 0
long_df["is_partial"] = (long_df["points"] > 0) & (long_df["points"] < long_df["price"])
long_df["global_rate"] = long_df["q_col"].map(lambda c: global_q_stats[c]["rate"])

rare_rate_threshold = (100 - rare_fail_pct) / 100.0
long_df["is_rare"] = long_df["is_correct"] & (long_df["global_rate"] <= rare_rate_threshold)

unique_qs = {c for c, s in global_q_stats.items() if s["correct"] == 1}
long_df["is_unique"] = long_df["is_correct"] & long_df["q_col"].isin(unique_qs)

per_game_q = long_df.groupby(["_game_label", "q_col"])["is_correct"].sum().reset_index(name="correct_in_game")
long_df = long_df.merge(per_game_q, on=["_game_label", "q_col"], how="left")
long_df["is_solo_in_game"] = long_df["is_correct"] & (long_df["correct_in_game"] == 1)


# ============================================================
# МЕТРИКИ СВЕРХУ
# ============================================================
m1, m2, m3, m4, m5 = st.columns(5)
m1.metric("Игр в срезе", len(selected_games) if selected_games else df["_game_label"].nunique())
m2.metric("Игроков", len(selected_players))
m3.metric("Вопросов в срезе", f"{len(filtered_q_cols)} / {len(q_cols)}")
m4.metric("Макс. очков в срезе", f"{q_passport_filtered['price'].sum():.0f}")
m5.metric("Фильтр", type_filter)

st.info(
    "Вся статистика ниже считается по текущему срезу: выбранные игры + тип вопроса + темы. "
    "Редкость ответа оценивается по всем играм в файле."
)


# ============================================================
# ВКЛАДКИ
# ============================================================
tab_charts, tab_overall, tab_h2h, tab_rare, tab_streaks, tab_themes, tab_difficulty, tab_virtual = st.tabs([
    "📊 Графики",
    "🏆 Общая статистика",
    "⚔️ Друг против друга",
    "💎 Редкие ответы",
    "🔥 Серии",
    "🎭 Темы и типы",
    "❓ Сложность вопросов",
    "🎲 Виртуальные игры",
])


# ============================================================
# ВКЛАДКА: ГРАФИКИ
# ============================================================
with tab_charts:
    if chart_type == "Накопленные очки":
        fig = go.Figure()
        mode = "lines+markers" if show_markers else "lines"
        for _, row in plot_df.iterrows():
            values = [float(row[c]) for c in filtered_q_cols]
            cumulative = pd.Series(values).cumsum().tolist()
            label = make_row_label(row)
            hover_texts = []
            for q_num, val in zip(filtered_q_numbers, values):
                meta = q_meta_lookup[q_num]
                hover_texts.append(
                    f"<b>{label}</b><br>Вопрос: Q{q_num}<br>Тема: {meta['theme']}<br>"
                    f"Тип: {meta['type_ru']}<br>Номинал: {meta['price']}<br>Очки за вопрос: {val:.0f}"
                )
            fig.add_trace(go.Scatter(
                x=filtered_q_numbers, y=cumulative, mode=mode,
                name=label, text=hover_texts, hoverinfo="text"
            ))
        fig.update_layout(
            title=f"Накопленные очки ({type_filter}, {len(filtered_q_cols)} вопр.)",
            xaxis_title="Номер вопроса", yaxis_title="Накопленные очки",
            hovermode="closest", legend_title="Игрок / игра", height=650
        )
        if len(filtered_q_numbers) <= 80:
            fig.update_xaxes(dtick=1)
        st.plotly_chart(fig, use_container_width=True)

    elif chart_type == "Очки за каждый вопрос":
        fig = go.Figure()
        for _, row in plot_df.iterrows():
            values = [float(row[c]) for c in filtered_q_cols]
            label = make_row_label(row)
            hover_texts = []
            for q_num, val in zip(filtered_q_numbers, values):
                meta = q_meta_lookup[q_num]
                hover_texts.append(
                    f"<b>{label}</b><br>Вопрос: Q{q_num}<br>Тема: {meta['theme']}<br>"
                    f"Тип: {meta['type_ru']}<br>Номинал: {meta['price']}<br>Очки: {val:.0f}"
                )
            fig.add_trace(go.Bar(
                x=filtered_q_numbers, y=values, name=label,
                text=hover_texts, hoverinfo="text"
            ))
        fig.update_layout(
            title=f"Очки за каждый вопрос ({type_filter})",
            xaxis_title="Номер вопроса", yaxis_title="Очки",
            barmode=bar_mode, legend_title="Игрок / игра", height=650
        )
        if len(filtered_q_numbers) <= 80:
            fig.update_xaxes(dtick=1)
        st.plotly_chart(fig, use_container_width=True)

    else:
        summary_rows = []
        for _, row in plot_df.iterrows():
            q_total = float(row[filtered_q_cols].sum())
            correct_count = int((row[filtered_q_cols] > 0).sum())
            summary_rows.append({
                "Игрок": row["Имя игрока"], "Игра": row["_game_label"],
                "Сумма по Q (срез)": q_total, "Правильных в срезе": correct_count,
                "Очки из файла (все)": row.get("Очки", None),
            })
        summary_df = pd.DataFrame(summary_rows)
        summary_df["Подпись"] = (
            summary_df["Игрок"] + " · " + summary_df["Игра"] if multiple_games else summary_df["Игрок"]
        )
        fig = go.Figure()
        fig.add_trace(go.Bar(
            x=summary_df["Подпись"], y=summary_df["Сумма по Q (срез)"],
            name="Сумма по срезу"
        ))
        if summary_df["Очки из файла (все)"].notna().any():
            fig.add_trace(go.Bar(
                x=summary_df["Подпись"], y=summary_df["Очки из файла (все)"],
                name="Очки из файла (все вопросы)"
            ))
        fig.update_layout(
            title=f"Итоги по срезу ({type_filter})", xaxis_title="Игрок",
            yaxis_title="Очки", barmode="group", height=550
        )
        st.plotly_chart(fig, use_container_width=True)
        st.dataframe(summary_df, use_container_width=True, hide_index=True)


# ============================================================
# ВКЛАДКА: ОБЩАЯ СТАТИСТИКА
# ============================================================
with tab_overall:
    st.subheader("Сводные показатели по игрокам")

    nominal_per_question = q_passport_filtered["price"].sum()

    overall_rows = []
    for player in selected_players:
        p_df = long_df[long_df["Имя игрока"] == player]
        games = p_df["_game_label"].nunique()
        questions = len(p_df)
        correct = int(p_df["is_correct"].sum())
        total_points = float(p_df["points"].sum())
        partial = int(p_df["is_partial"].sum())
        rare = int(p_df["is_rare"].sum())
        unique = int(p_df["is_unique"].sum())
        solo = int(p_df["is_solo_in_game"].sum())
        max_nominal = games * nominal_per_question
        conversion = pct(total_points, max_nominal)
        avg_per_q = round(total_points / questions, 1) if questions else 0

        best_streak = 0
        for _, g_row in plot_df[plot_df["Имя игрока"] == player].iterrows():
            vals = [g_row[c] for c in filtered_q_cols]
            best_streak = max(best_streak, streak_stats(vals)["Лучшая серия"])

        overall_rows.append({
            "Игрок": player,
            "Игр": games,
            "Вопросов": questions,
            "Верных": correct,
            "Неверных/пропуск": questions - correct,
            "% ответов": pct(correct, questions),
            "Очки (срез)": round(total_points, 0),
            "Среднее/вопрос": avg_per_q,
            "Конверсия номинала, %": conversion,
            "Частичных": partial,
            "Редких": rare,
            "Уникальных": unique,
            "Соло в игре": solo,
            "Лучшая серия": best_streak,
        })

    overall_df = pd.DataFrame(overall_rows).sort_values("Очки (срез)", ascending=False)
    st.dataframe(overall_df, use_container_width=True, hide_index=True)


# ============================================================
# ВКЛАДКА: ДРУГ ПРОТИВ ДРУГА
# ============================================================
with tab_h2h:
    st.subheader("⚔️ Кто кого переиграл")
    st.caption(
        "Число вопросов в срезе, где игрок по строке ответил верно, "
        "а игрок по столбцу — нет (в одной игре)."
    )

    pivot = long_df.pivot_table(
        index=["_game_label", "q_col"], columns="Имя игрока",
        values="is_correct", aggfunc="max"
    ).fillna(False)

    h2h_matrix = {}
    for a in selected_players:
        h2h_matrix[a] = {}
        for b in selected_players:
            if a == b:
                h2h_matrix[a][b] = None
            elif a in pivot.columns and b in pivot.columns:
                h2h_matrix[a][b] = int(((pivot[a] == True) & (pivot[b] == False)).sum())
            else:
                h2h_matrix[a][b] = 0

    h2h_df = pd.DataFrame(h2h_matrix).T
    h2h_df.index.name = "Игрок"
    st.dataframe(h2h_df.fillna("—"), use_container_width=True)

    st.markdown("---")
    st.subheader("📊 Процент доминирования")
    st.caption(
        "Доля побед игрока по строке над игроком по столбцу. "
        "50% — игра на равных, >50% — игрок сверху сильнее."
    )

    dom_matrix = {}
    for a in selected_players:
        dom_matrix[a] = {}
        for b in selected_players:
            if a == b:
                dom_matrix[a][b] = None
            else:
                a_vs_b = h2h_matrix[a][b]
                b_vs_a = h2h_matrix[b][a]
                total = a_vs_b + b_vs_a
                if total == 0:
                    dom_matrix[a][b] = 50.0
                else:
                    dom_matrix[a][b] = round(100.0 * a_vs_b / total, 1)

    dom_df = pd.DataFrame(dom_matrix).T
    dom_df.index.name = "Игрок"
    st.dataframe(dom_df.fillna("—").astype(str).replace({"None": "—"}), use_container_width=True)

    st.markdown("---")
    st.subheader("👑 Итоговый рейтинг доминирования")
    st.caption("Средний процент доминирования игрока над всеми соперниками в срезе.")

    dom_scores = []
    for a in selected_players:
        vals = [dom_matrix[a][b] for b in selected_players if b != a and dom_matrix[a][b] is not None]
        avg_dom = round(sum(vals) / len(vals), 1) if vals else 0
        dom_scores.append({"Игрок": a, "Средний % доминирования": avg_dom})

    dom_scores_df = pd.DataFrame(dom_scores).sort_values("Средний % доминирования", ascending=False)
    st.dataframe(dom_scores_df, use_container_width=True, hide_index=True)

    if len(dom_scores_df) > 1:
        fig_dom = go.Figure(go.Bar(
            x=dom_scores_df["Средний % доминирования"],
            y=dom_scores_df["Игрок"],
            orientation="h",
            marker_color=["#2ecc71" if v > 50 else "#e74c3c" for v in dom_scores_df["Средний % доминирования"]]
        ))
        fig_dom.update_layout(
            title="Средний процент доминирования (>50% — доминирует)",
            xaxis_title="% доминирования", yaxis_title="Игрок", height=400
        )
        st.plotly_chart(fig_dom, use_container_width=True)


# ============================================================
# ВКЛАДКА: РЕДКИЕ ОТВЕТЫ
# ============================================================
with tab_rare:
    st.subheader(f"Редкие ответы (не ответили ≥ {rare_fail_pct}% игроков)")

    rare_instances = long_df[long_df["is_rare"]].copy()
    rare_instances["Ответили всего"] = rare_instances["q_col"].map(lambda c: global_q_stats[c]["correct"])
    rare_instances["% ответили"] = (rare_instances["global_rate"] * 100).round(1)

    show_rare = rare_instances[[
        "Имя игрока", "_game_label", "q_col", "theme", "type_ru",
        "price", "points", "Ответили всего", "% ответили"
    ]].rename(columns={
        "_game_label": "Игра", "q_col": "Вопрос", "theme": "Тема",
        "type_ru": "Тип", "price": "Номинал", "points": "Взято очков"
    }).sort_values(["Имя игрока", "Вопрос"])

    if show_rare.empty:
        st.info("В текущем срезе редких ответов нет. Попробуйте снизить порог.")
    else:
        st.dataframe(show_rare, use_container_width=True, hide_index=True)

    st.markdown("---")
    st.subheader("Уникальные ответы (игрок — единственный, кто взял вопрос)")
    uniq_instances = long_df[long_df["is_unique"]][
        ["Имя игрока", "_game_label", "q_col", "theme", "price", "points"]
    ].rename(columns={
        "_game_label": "Игра", "q_col": "Вопрос", "theme": "Тема",
        "price": "Номинал", "points": "Взято очков"
    }).sort_values(["Имя игрока", "Вопрос"])
    if uniq_instances.empty:
        st.info("Уникальных ответов в срезе нет.")
    else:
        st.dataframe(uniq_instances, use_container_width=True, hide_index=True)


# ============================================================
# ВКЛАДКА: СЕРИИ
# ============================================================
with tab_streaks:
    st.subheader("Серии верных ответов подряд")

    streak_records = []
    for _, row in plot_df.iterrows():
        vals = [row[c] for c in filtered_q_cols]
        s = streak_stats(vals)
        streak_records.append({"Имя игрока": row["Имя игрока"], "Игра": row["_game_label"], **s})
    streak_df = pd.DataFrame(streak_records)

    st.markdown("**По каждой игре:**")
    st.dataframe(streak_df, use_container_width=True, hide_index=True)

    st.markdown("**Лучшие показатели по игрокам:**")
    best = streak_df.groupby("Имя игрока").agg(
        **{
            "Макс. серия": ("Лучшая серия", "max"),
            "Серий всего": ("Всего серий", "sum"),
            "Серий 3+ (суммарно)": ("Серий 3+", "sum"),
            "Серий 5+ (суммарно)": ("Серий 5+", "sum"),
        }
    ).reset_index().sort_values("Макс. серия", ascending=False)
    st.dataframe(best, use_container_width=True, hide_index=True)


# ============================================================
# ВКЛАДКА: ТЕМЫ И ТИПЫ
# ============================================================
with tab_themes:
    st.subheader("Успешность по темам (% верных)")
    theme_pivot = long_df.pivot_table(
        index="Имя игрока", columns="theme", values="is_correct", aggfunc="mean"
    ) * 100
    st.dataframe(theme_pivot.round(1), use_container_width=True)

    st.markdown("---")
    st.subheader("Текстовые vs Картинные")
    type_pivot = long_df.pivot_table(
        index="Имя игрока", columns="type_ru", values="is_correct", aggfunc="mean"
    ) * 100
    st.dataframe(type_pivot.round(1), use_container_width=True)

    st.markdown("---")
    st.subheader("Очки по раундам")
    round_points = long_df.pivot_table(
        index="Имя игрока", columns="round", values="points", aggfunc="sum"
    ).fillna(0).astype(int)
    st.dataframe(round_points, use_container_width=True)


# ============================================================
# ВКЛАДКА: СЛОЖНОСТЬ ВОПРОСОВ
# ============================================================
with tab_difficulty:
    st.subheader("Сложность вопросов в срезе (по всем играм в файле)")
    st.caption("Считается по всем игрокам и играм в файле, не только по выбранным.")

    diff_rows = []
    for q_col in filtered_q_cols:
        meta = QUESTION_PASSPORT[QUESTION_PASSPORT["q_col"] == q_col].iloc[0]
        s = global_q_stats[q_col]
        rate = s["rate"]
        if rate == 0:
            tier = "💀 Мёртвый"
        elif rate < 0.25:
            tier = "🔴 Жёсткий"
        elif rate < 0.60:
            tier = "🟡 Средний"
        else:
            tier = "🟢 Лёгкий"
        diff_rows.append({
            "Вопрос": q_col, "Тема": meta["theme"], "Тип": meta["type_ru"],
            "Раунд": meta["round"], "Номинал": meta["price"],
            "Ответили": s["correct"], "Всего": s["total"],
            "% ответили": round(100 * rate, 1), "Уровень": tier,
        })
    diff_df = pd.DataFrame(diff_rows)
    st.dataframe(diff_df, use_container_width=True, hide_index=True)

    st.markdown("---")
    col_a, col_b = st.columns(2)
    with col_a:
        st.markdown("**Самые сложные:**")
        st.dataframe(
            diff_df.sort_values("% ответили").head(10)[["Вопрос", "Тема", "Номинал", "% ответили", "Уровень"]],
            use_container_width=True, hide_index=True
        )
    with col_b:
        st.markdown("**Самые лёгкие:**")
        st.dataframe(
            diff_df.sort_values("% ответили", ascending=False).head(10)[["Вопрос", "Тема", "Номинал", "% ответили", "Уровень"]],
            use_container_width=True, hide_index=True
        )


# ============================================================
# ВКЛАДКА: ВИРТУАЛЬНЫЕ ИГРЫ
# ============================================================
with tab_virtual:
    st.subheader("🎲 Виртуальные игры: сборные составы")
    st.caption(
        "Формируйте составы из игроков разных игр. Взвешенные очки пересчитываются "
        "по тому же принципу, что и для всех игр, но базовый вес каждого вопроса "
        "считается только по выбранным игрокам."
    )

    vmode = st.radio(
        "Режим",
        ["Ручной выбор состава", "Перебор всех комбинаций"],
        horizontal=True
    )

    all_pg = df[["Имя игрока", "_game_label"]].drop_duplicates()
    all_pg["label"] = all_pg["Имя игрока"].astype(str) + " — " + all_pg["_game_label"].astype(str)
    all_options = all_pg["label"].tolist()

    def select_rows_by_labels(labels):
        if not labels:
            return pd.DataFrame()
        row_labels = df["Имя игрока"].astype(str) + " — " + df["_game_label"].astype(str)
        return df[row_labels.isin(labels)].copy()

    if vmode == "Ручной выбор состава":
        chosen = st.multiselect(
            "Игроки в составе виртуальной игры",
            all_options,
            default=all_options[:3]
        )

        use_slice = st.checkbox(
            "Использовать текущий срез вопросов (иначе все вопросы)",
            value=False
        )
        active_q = filtered_q_cols if use_slice else q_cols

        if chosen:
            group_df = select_rows_by_labels(chosen)
            result_df, base_weights = recalc_for_group(group_df, QUESTION_PASSPORT, active_q)

            st.markdown("#### Результаты виртуальной игры")
            st.dataframe(
                result_df[['Имя игрока', 'Игра', 'Очки', 'Ответы', 'wОчки', 'fОчки']]
                .sort_values('fОчки', ascending=False),
                use_container_width=True, hide_index=True
            )

            if not result_df.empty:
                fig_v = go.Figure(go.Bar(
                    x=result_df['fОчки'],
                    y=result_df['Имя игрока'] + ' (' + result_df['Игра'] + ')',
                    orientation='h',
                    marker_color='#3498db'
                ))
                fig_v.update_layout(
                    title="fОчки игроков в виртуальной игре",
                    xaxis_title="fОчки", height=400
                )
                st.plotly_chart(fig_v, use_container_width=True)

            with st.expander("Базовые веса вопросов в этой виртуальной игре"):
                bw_rows = []
                for q_col, w in base_weights.items():
                    meta = QUESTION_PASSPORT[QUESTION_PASSPORT['q_col'] == q_col]
                    theme = meta['theme'].iloc[0] if not meta.empty else ''
                    qnum = meta['q_num'].iloc[0] if not meta.empty else ''
                    bw_rows.append({
                        'Вопрос': q_col, '№': qnum,
                        'Тема': theme, 'Базовый вес': round(w, 3)
                    })
                st.dataframe(pd.DataFrame(bw_rows), use_container_width=True, hide_index=True)
        else:
            st.info("Выберите хотя бы одного игрока.")

    else:
        col_a, col_b = st.columns(2)
        with col_a:
            group_size = st.slider("Размер команды", 2, 6, 3)
        with col_b:
            top_k = st.slider("Показать топ-N составов", 5, 200, 20)

        use_slice = st.checkbox(
            "Использовать текущий срез вопросов (иначе все вопросы)",
            value=False, key="combo_slice"
        )
        active_q = filtered_q_cols if use_slice else q_cols

        pool = st.multiselect(
            "Пул игроков (пусто = все игроки)",
            all_options,
            help="Ограничьте пул, если комбинаций слишком много."
        )
        pool_df = select_rows_by_labels(pool) if pool else df.copy()

        n_pool = len(pool_df)
        n_combos = math.comb(n_pool, group_size) if n_pool >= group_size else 0
        st.info(f"Пул: **{n_pool}** игроков. Комбинаций по {group_size}: **{n_combos}**.")

        MAX_COMBOS = 30000
        if n_combos > MAX_COMBOS:
            st.warning(
                f"Слишком много комбинаций ({n_combos}). Максимум {MAX_COMBOS}. "
                "Уменьшите пул или размер команды."
            )
        elif n_combos == 0:
            st.info("Недостаточно игроков для выбранного размера команды.")
        else:
            if st.button("🚀 Запустить перебор"):
                combo_results = []
                progress = st.progress(0)
                total = n_combos

                for ci, combo in enumerate(combinations(range(n_pool), group_size)):
                    subset = pool_df.iloc[list(combo)]
                    res, _ = recalc_for_group(subset, QUESTION_PASSPORT, active_q)
                    combo_results.append({
                        'Состав': ' + '.join(subset['Имя игрока'].astype(str)),
                        'Средний fОчки': round(res['fОчки'].mean(), 2),
                        'Сумма fОчки': round(res['fОчки'].sum(), 2),
                        'Среднее Очки': round(res['Очки'].mean(), 1),
                        'Средний % ответов': round(100 * res['Ответы'].mean() / len(active_q), 1),
                    })
                    if ci % 200 == 0:
                        progress.progress(int(100 * ci / total))
                progress.progress(100)

                combo_df = pd.DataFrame(combo_results).sort_values(
                    'Средний fОчки', ascending=False
                )

                st.markdown(f"#### Топ-{min(top_k, len(combo_df))} составов по среднему fОчки")
                st.dataframe(combo_df.head(top_k), use_container_width=True, hide_index=True)

                csv = combo_df.to_csv(index=False).encode('utf-8-sig')
                st.download_button(
                    "Скачать все комбинации (CSV)",
                    csv,
                    "virtual_games_combinations.csv",
                    "text/csv"
                )

                top_plot = combo_df.head(top_k)
                fig_c = go.Figure(go.Bar(
                    x=top_plot['Средний fОчки'],
                    y=top_plot['Состав'],
                    orientation='h',
                    marker_color='#2ecc71'
                ))
                fig_c.update_layout(
                    title="Топ составов по среднему fОчки",
                    xaxis_title="Средний fОчки",
                    height=min(700, 30 * top_k + 100),
                    yaxis={'categoryorder': 'total ascending'}
                )
                st.plotly_chart(fig_c, use_container_width=True)


# ============================================================
# ПАСПОРТ И ИСХОДНЫЕ ДАННЫЕ
# ============================================================
with st.expander(f"📖 Паспорт вопросов в срезе ({len(filtered_q_cols)} шт.)"):
    st.dataframe(
        q_passport_filtered[["q_num", "theme", "type_ru", "round", "price"]]
        .rename(columns={"q_num": "№", "theme": "Тема", "type_ru": "Тип",
                         "round": "Раунд", "price": "Номинал"}),
        use_container_width=True, hide_index=True
    )

with st.expander("📋 Данные игроков по выбранным вопросам"):
    display_columns = ["Имя игрока", "_game_label"]
    if "Очки" in plot_df.columns:
        display_columns.append("Очки")
    display_columns.extend(filtered_q_cols)
    st.dataframe(
        plot_df[display_columns].rename(columns={"_game_label": "Игра"}),
        use_container_width=True, hide_index=True
    )
