"""
栄養トラッカー アプリ（日付対応・棒グラフ版）
=============================================
使い方:
  1. pip install streamlit pandas plotly openpyxl
  2. 抽出結果.xlsx と同じフォルダに置く
  3. streamlit run nutrition_tracker.py
"""

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from datetime import date

# ─────────────────────────────────────────
# 1. ページ設定
# ─────────────────────────────────────────
st.set_page_config(
    page_title="栄養トラッカー",
    page_icon="🥗",
    layout="centered",
    initial_sidebar_state="collapsed",
)

# ─────────────────────────────────────────
# 2. カスタムCSS（文字・背景を見やすく）
# ─────────────────────────────────────────
st.markdown("""
<style>
/* 全体背景・文字色 */
.stApp { background-color: #1a1a2e; color: #ffffff !important; }

/* 全テキストを白に */
p, span, div, label, li { color: #ffffff !important; }

/* 見出し */
h1, h2, h3, h4 { color: #ffffff !important; }

/* キャプション */
.stCaption { color: #cccccc !important; }

/* テキスト入力・数値入力 */
.stTextInput input, .stNumberInput input {
    background-color: #2a2a4a !important;
    color: #ffffff !important;
    border: 2px solid #4a4a7a !important;
    border-radius: 10px !important;
    font-size: 16px !important;
}

/* セレクトボックス */
.stSelectbox > div > div {
    background-color: #2a2a4a !important;
    color: #ffffff !important;
    border: 2px solid #4a4a7a !important;
}

/* 日付ピッカー */
.stDateInput input {
    background-color: #2a2a4a !important;
    color: #ffffff !important;
    border: 2px solid #4a4a7a !important;
}

/* ラジオボタンのラベル */
.stRadio label { color: #ffffff !important; font-size: 15px !important; }

/* ボタン（通常） */
.stButton > button {
    background-color: #2a2a4a !important;
    color: #ffffff !important;
    border: 2px solid #4a4a7a !important;
    border-radius: 10px !important;
    font-weight: 600 !important;
    font-size: 14px !important;
}
.stButton > button:hover {
    background-color: #3a3a6a !important;
    border-color: #ff6b35 !important;
}

/* プライマリボタン */
.stButton > button[kind="primary"] {
    background-color: #ff6b35 !important;
    color: #ffffff !important;
    border: none !important;
    font-size: 16px !important;
}

/* メトリクスカード */
[data-testid="stMetric"] {
    background: #2a2a4a !important;
    border-radius: 14px !important;
    padding: 16px 20px !important;
    border: 2px solid #4a4a7a !important;
}
[data-testid="stMetricLabel"] > div { color: #dddddd !important; font-size: 14px !important; }
[data-testid="stMetricValue"] { color: #ffffff !important; font-size: 2rem !important; font-weight: 800 !important; }

/* タブ */
.stTabs [data-baseweb="tab"] { color: #cccccc !important; font-size: 15px !important; font-weight: 600 !important; }
.stTabs [aria-selected="true"] { color: #ff6b35 !important; border-bottom-color: #ff6b35 !important; }

/* データフレーム */
[data-testid="stDataFrame"] { color: #ffffff !important; }
.dataframe { color: #ffffff !important; background: #2a2a4a !important; }

/* 区切り線 */
hr { border-color: #4a4a7a !important; }

/* 成功メッセージ */
.stSuccess { background-color: #1a3a2a !important; color: #88ffaa !important; border: 1px solid #44aa66 !important; }
.stSuccess p { color: #88ffaa !important; }

/* infoメッセージ */
.stInfo { background-color: #1a2a3a !important; color: #88ccff !important; }
.stInfo p { color: #88ccff !important; }
</style>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────
# 3. 食品データの読み込み
# ─────────────────────────────────────────
@st.cache_data
def load_food_data() -> pd.DataFrame:
    df = pd.read_excel("抽出結果.xlsx")
    df = df.iloc[11:].copy()
    df.columns = ["食品名", "カロリー(kcal)", "たんぱく質(g)", "シート名"]
    df["カロリー(kcal)"] = pd.to_numeric(df["カロリー(kcal)"], errors="coerce")
    df["たんぱく質(g)"]  = pd.to_numeric(df["たんぱく質(g)"],  errors="coerce")
    df = df.dropna(subset=["食品名", "カロリー(kcal)", "たんぱく質(g)"])
    df["食品名"] = df["食品名"].astype(str).str.strip()
    return df[["食品名", "カロリー(kcal)", "たんぱく質(g)"]].reset_index(drop=True)


# ─────────────────────────────────────────
# 4. セッションステートの初期化
# ─────────────────────────────────────────
if "meal_log" not in st.session_state:
    st.session_state.meal_log = {}       # {"2026-06-11": [{...}, ...], ...}

if "selected_date" not in st.session_state:
    st.session_state.selected_date = date.today()


# ─────────────────────────────────────────
# 5. データ読み込み
# ─────────────────────────────────────────
try:
    food_df = load_food_data()
except FileNotFoundError:
    st.error("⚠️ 抽出結果.xlsx が見つかりません。同じフォルダに置いてください。")
    st.stop()


# ─────────────────────────────────────────
# 6. ヘッダー
# ─────────────────────────────────────────
st.markdown("# 🥗 栄養トラッカー")
st.caption(f"食品データ: {len(food_df):,}件（100gあたりの栄養素）")
st.divider()


# ─────────────────────────────────────────
# 7. グラフ共通レイアウト設定
# ─────────────────────────────────────────
CHART_LAYOUT = dict(
    paper_bgcolor="rgba(0,0,0,0)",
    plot_bgcolor="#1e1e3a",
    font=dict(color="#ffffff", size=13),
    margin=dict(t=20, b=20, l=10, r=10),
    yaxis=dict(
        gridcolor="#3a3a5a",
        color="#ffffff",
        tickfont=dict(color="#ffffff", size=12),
        titlefont=dict(color="#ffffff"),
    ),
    xaxis=dict(
        color="#ffffff",
        tickfont=dict(color="#ffffff", size=12),
        tickangle=-30,
    ),
    legend=dict(
        font=dict(color="#ffffff", size=12),
        bgcolor="rgba(0,0,0,0)",
    ),
)


# ─────────────────────────────────────────
# 8. タブ
# ─────────────────────────────────────────
tab_search, tab_log, tab_compare = st.tabs([
    "🔍 食品検索・追加",
    "📊 今日の記録",
    "📅 日別比較",
])


# ══════════════════════════════════════════
# タブ①: 食品検索・追加
# ══════════════════════════════════════════
with tab_search:

    target_date = st.date_input(
        "📅 記録する日付",
        value=st.session_state.selected_date,
        max_value=date.today(),
    )
    st.session_state.selected_date = target_date
    date_key = str(target_date)

    st.divider()

    query = st.text_input(
        "食品名を検索",
        placeholder="例：鶏肉、ご飯、豆腐",
        label_visibility="collapsed",
    )

    if query:
        filtered = food_df[food_df["食品名"].str.contains(query, na=False)]
    else:
        filtered = food_df.head(50)

    st.caption(f"{len(filtered):,}件 表示中")

    if filtered.empty:
        st.info("🔍 見つかりませんでした。別のキーワードで試してください。")
    else:
        selected_name = st.selectbox(
            "食品を選択",
            options=filtered["食品名"].tolist(),
            label_visibility="collapsed",
        )

        selected_row = filtered[filtered["食品名"] == selected_name].iloc[0]
        cal_per_100  = selected_row["カロリー(kcal)"]
        prot_per_100 = selected_row["たんぱく質(g)"]

        st.divider()

        c1, c2 = st.columns(2)
        c1.metric("🔥 カロリー（100g）", f"{cal_per_100} kcal")
        c2.metric("💪 たんぱく質（100g）", f"{prot_per_100} g")

        st.divider()
        st.markdown("**グラム数を入力**")

        quick_cols = st.columns(6)
        for i, g in enumerate([50, 80, 100, 150, 200, 300]):
            if quick_cols[i].button(f"{g}g", key=f"quick_{g}"):
                st.session_state["gram_input"] = float(g)

        gram = st.number_input(
            "グラム数",
            min_value=1.0,
            max_value=9999.0,
            value=st.session_state.get("gram_input", 100.0),
            step=10.0,
            format="%.0f",
            label_visibility="collapsed",
            key="gram_input",
        )

        calc_cal  = round(cal_per_100 * gram / 100)
        calc_prot = round(prot_per_100 * gram / 100, 1)

        st.markdown(f"### {gram:.0f}g あたり")
        p1, p2 = st.columns(2)
        p1.metric("🔥 カロリー", f"{calc_cal} kcal")
        p2.metric("💪 たんぱく質", f"{calc_prot} g")

        st.divider()

        if st.button("➕ 記録に追加する", type="primary", use_container_width=True):
            if date_key not in st.session_state.meal_log:
                st.session_state.meal_log[date_key] = []
            st.session_state.meal_log[date_key].append({
                "食品名":          selected_name,
                "グラム":          int(gram),
                "カロリー(kcal)":  calc_cal,
                "たんぱく質(g)":   calc_prot,
            })
            st.success(f"✅ {selected_name}  {gram:.0f}g を {target_date} に追加しました！")
            st.balloons()


# ══════════════════════════════════════════
# タブ②: 今日の記録
# ══════════════════════════════════════════
with tab_log:

    view_date = st.date_input(
        "📅 表示する日付",
        value=st.session_state.selected_date,
        key="view_date",
        max_value=date.today(),
    )
    view_key = str(view_date)
    log = st.session_state.meal_log.get(view_key, [])

    st.divider()

    if not log:
        st.markdown("""
        <div style="text-align:center; padding:40px 0;">
            <div style="font-size:48px;">🍽️</div>
            <div style="font-size:18px; color:#ffffff; margin-top:12px; font-weight:600;">この日の記録はありません</div>
            <div style="font-size:14px; color:#cccccc; margin-top:6px;">「食品検索・追加」タブから追加しましょう</div>
        </div>
        """, unsafe_allow_html=True)
    else:
        total_cal  = sum(e["カロリー(kcal)"] for e in log)
        total_prot = sum(e["たんぱく質(g)"]  for e in log)

        t1, t2 = st.columns(2)
        t1.metric("🔥 カロリー合計",   f"{total_cal:,} kcal")
        t2.metric("💪 たんぱく質合計", f"{total_prot:.1f} g")

        st.divider()

        log_df = pd.DataFrame(log)

        COLORS_CAL  = ["#ff6b35","#ff8c5a","#ffad80","#f4a26b","#e8743b","#d4512b","#c03a1c","#ff4500"]
        COLORS_PROT = ["#4ecdc4","#6dd5cd","#3ab5ac","#29a299","#5bc8c0","#8cddd6","#198f86","#00b8a9"]

        # カロリー 積み上げ棒グラフ
        st.markdown("#### 🔥 カロリー内訳")
        fig_cal = go.Figure()
        for i, row in log_df.iterrows():
            fig_cal.add_trace(go.Bar(
                name=row["食品名"],
                x=["カロリー"],
                y=[row["カロリー(kcal)"]],
                marker_color=COLORS_CAL[i % len(COLORS_CAL)],
                text=f"{row['カロリー(kcal)']}kcal",
                textposition="inside",
                textfont=dict(color="#ffffff", size=13),
            ))
        fig_cal.update_layout(
            **CHART_LAYOUT,
            barmode="stack",
            height=320,
            yaxis=dict(**CHART_LAYOUT["yaxis"], title="kcal"),
            xaxis=dict(showticklabels=False),
        )
        st.plotly_chart(fig_cal, use_container_width=True)

        # たんぱく質 積み上げ棒グラフ
        st.markdown("#### 💪 たんぱく質内訳")
        fig_prot = go.Figure()
        for i, row in log_df.iterrows():
            fig_prot.add_trace(go.Bar(
                name=row["食品名"],
                x=["たんぱく質"],
                y=[row["たんぱく質(g)"]],
                marker_color=COLORS_PROT[i % len(COLORS_PROT)],
                text=f"{row['たんぱく質(g)']}g",
                textposition="inside",
                textfont=dict(color="#ffffff", size=13),
            ))
        fig_prot.update_layout(
            **CHART_LAYOUT,
            barmode="stack",
            height=320,
            yaxis=dict(**CHART_LAYOUT["yaxis"], title="g"),
            xaxis=dict(showticklabels=False),
        )
        st.plotly_chart(fig_prot, use_container_width=True)

        st.divider()

        # 食事ログ一覧
        st.markdown("#### 📋 食事ログ")
        for i, entry in enumerate(log):
            cols = st.columns([3, 1, 1, 1, 0.5])
            cols[0].markdown(f"**{entry['食品名']}**")
            cols[1].markdown(f"**{entry['グラム']}g**")
            cols[2].markdown(f"🔥 **{entry['カロリー(kcal)']}** kcal")
            cols[3].markdown(f"💪 **{entry['たんぱく質(g)']}** g")
            if cols[4].button("✕", key=f"del_{view_key}_{i}"):
                st.session_state.meal_log[view_key].pop(i)
                st.rerun()

        st.divider()
        if st.button("🗑️ この日の記録を全て削除", type="secondary", use_container_width=True):
            st.session_state.meal_log[view_key] = []
            st.rerun()


# ══════════════════════════════════════════
# タブ③: 日別比較
# ══════════════════════════════════════════
with tab_compare:

    st.markdown("#### 📅 日別 カロリー・たんぱく質 比較")

    all_dates = sorted(st.session_state.meal_log.keys())

    if not all_dates:
        st.markdown("""
        <div style="text-align:center; padding:40px 0;">
            <div style="font-size:48px;">📅</div>
            <div style="font-size:18px; color:#ffffff; margin-top:12px; font-weight:600;">まだ記録がありません</div>
            <div style="font-size:14px; color:#cccccc; margin-top:6px;">複数日記録すると日別比較ができます</div>
        </div>
        """, unsafe_allow_html=True)
    else:
        # 日付ごとに合計を集計
        summary_rows = []
        for d in all_dates:
            entries = st.session_state.meal_log[d]
            if entries:
                summary_rows.append({
                    "日付":           d,
                    "カロリー(kcal)": sum(e["カロリー(kcal)"] for e in entries),
                    "たんぱく質(g)":  round(sum(e["たんぱく質(g)"] for e in entries), 1),
                })

        if not summary_rows:
            st.info("記録のある日がありません。")
        else:
            summary_df = pd.DataFrame(summary_rows)

            range_label = st.radio(
                "表示範囲",
                options=["直近7日", "直近30日", "全期間"],
                horizontal=True,
            )
            limit = {"直近7日": 7, "直近30日": 30, "全期間": 9999}[range_label]
            display_df = summary_df.tail(limit)

            # カロリー 日別棒グラフ
            st.markdown("#### 🔥 カロリー（日別）")
            fig_d_cal = go.Figure(go.Bar(
                x=display_df["日付"],
                y=display_df["カロリー(kcal)"],
                marker_color="#ff6b35",
                text=display_df["カロリー(kcal)"].astype(str) + " kcal",
                textposition="outside",
                textfont=dict(color="#ffffff", size=13),
            ))
            fig_d_cal.update_layout(
                **CHART_LAYOUT,
                height=320,
                yaxis=dict(**CHART_LAYOUT["yaxis"], title="kcal"),
            )
            st.plotly_chart(fig_d_cal, use_container_width=True)

            # たんぱく質 日別棒グラフ
            st.markdown("#### 💪 たんぱく質（日別）")
            fig_d_prot = go.Figure(go.Bar(
                x=display_df["日付"],
                y=display_df["たんぱく質(g)"],
                marker_color="#4ecdc4",
                text=display_df["たんぱく質(g)"].astype(str) + " g",
                textposition="outside",
                textfont=dict(color="#ffffff", size=13),
            ))
            fig_d_prot.update_layout(
                **CHART_LAYOUT,
                height=320,
                yaxis=dict(**CHART_LAYOUT["yaxis"], title="g"),
            )
            st.plotly_chart(fig_d_prot, use_container_width=True)

            st.divider()

            # 日別サマリー表
            st.markdown("#### 📋 日別サマリー")
            st.dataframe(
                display_df.set_index("日付").style.set_properties(**{
                    "color": "white",
                    "background-color": "#2a2a4a",
                    "font-size": "15px",
                }),
                use_container_width=True,
            )
