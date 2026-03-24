import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from datetime import datetime, timedelta
import os
import glob
import gc

st.set_page_config(
    page_title="米国株RSダッシュボード",
    page_icon="📈",
    layout="wide"
)

st.title("📈 米国株RS分析ダッシュボード")
st.markdown("---")

DATA_FOLDER = "data"


# =============================================
# ユーティリティ関数
# =============================================

def get_display_date(screening_date: datetime) -> datetime:
    return screening_date - timedelta(days=1)


def get_year_month_from_date(date: datetime) -> str:
    return date.strftime('%Y年%m月')


def get_available_months(all_data: list) -> list:
    months = set()
    for data in all_data:
        months.add(get_year_month_from_date(data['display_date']))
    return sorted(list(months), reverse=True)


def filter_data_by_month(all_data: list, selected_month: str) -> list:
    return [
        data for data in all_data
        if get_year_month_from_date(data['display_date']) == selected_month
    ]


# =============================================
# RS値 → 背景色（matplotlib不要）
# =============================================

def rs_to_bgcolor(val: float) -> str:
    try:
        v = float(val)
    except (TypeError, ValueError):
        return ''

    v = max(0.0, min(100.0, v))

    stops = [
        (0,   0xd7, 0x30, 0x27),
        (25,  0xf4, 0x6d, 0x43),
        (50,  0xff, 0xff, 0xbf),
        (75,  0x74, 0xc4, 0x76),
        (100, 0x1a, 0x96, 0x41),
    ]

    for i in range(len(stops) - 1):
        v0, r0, g0, b0 = stops[i]
        v1, r1, g1, b1 = stops[i + 1]
        if v0 <= v <= v1:
            t = (v - v0) / (v1 - v0)
            r = int(r0 + t * (r1 - r0))
            g = int(g0 + t * (g1 - g0))
            b = int(b0 + t * (b1 - b0))
            brightness = 0.299 * r + 0.587 * g + 0.114 * b
            fg = '#000000' if brightness > 128 else '#ffffff'
            return f'background-color: #{r:02x}{g:02x}{b:02x}; color: {fg}'

    return ''


def color_rs_col(series: pd.Series) -> list:
    return [rs_to_bgcolor(v) for v in series]


def color_diff_col(series: pd.Series) -> list:
    styles = []
    for v in series:
        try:
            val = int(v)
        except (TypeError, ValueError):
            styles.append('')
            continue
        if val > 0:
            styles.append('color: #2e7d32; font-weight: bold')
        elif val < 0:
            styles.append('color: #c62828; font-weight: bold')
        else:
            styles.append('')
    return styles


# =============================================
# データ読み込み
# =============================================

@st.cache_data(ttl=300)
def load_all_data(data_folder: str = DATA_FOLDER) -> list:
    all_data = []

    if not os.path.exists(data_folder):
        st.error(f"フォルダが見つかりません: {data_folder}")
        return all_data

    excel_files = sorted(
        glob.glob(os.path.join(data_folder, "*.xlsx")) +
        glob.glob(os.path.join(data_folder, "*.xls"))
    )

    if not excel_files:
        st.warning(f"{data_folder} にExcelファイルが見つかりません。")
        return all_data

    progress_bar = st.progress(0)
    status_ph = st.empty()

    for idx, file_path in enumerate(excel_files):
        filename = os.path.basename(file_path)
        status_ph.text(f"読み込み中: {filename}  ({idx+1}/{len(excel_files)})")

        try:
            parts = filename.split('_')
            date_str = next((p for p in parts if len(p) == 8 and p.isdigit()), None)
            date = datetime.strptime(date_str, '%Y%m%d') if date_str else datetime.now()

            with pd.ExcelFile(file_path) as excel:

                sector_rs_df = None
                if 'Screening_Results' in excel.sheet_names:
                    raw = excel.parse(
                        'Screening_Results',
                        usecols=['Sector', 'Sector_RS_Pct_CW', 'Sector_RS_Pct_EW']
                    )
                    sector_rs_df = (
                        raw.dropna(subset=['Sector'])
                           .groupby('Sector', as_index=False)
                           .agg(
                               Sector_RS_Pct_CW=('Sector_RS_Pct_CW', 'first'),
                               Sector_RS_Pct_EW=('Sector_RS_Pct_EW', 'first'),
                           )
                    )

                market_summary = None
                if 'Market_Summary' in excel.sheet_names:
                    ms = excel.parse('Market_Summary')
                    ms_dict = dict(zip(ms.iloc[:, 0], ms.iloc[:, 1]))
                    market_summary = {
                        'status': ms_dict.get('総合判定', ''),
                        'score':  ms_dict.get('スコア率', '')
                    }

            all_data.append({
                'date':           date,
                'display_date':   get_display_date(date),
                'sector_rs_df':   sector_rs_df,
                'market_summary': market_summary,
                'filename':       filename,
            })

        except Exception as e:
            st.warning(f"読み込みエラー ({filename}): {e}")

        progress_bar.progress((idx + 1) / len(excel_files))

    progress_bar.empty()
    status_ph.empty()
    gc.collect()
    return all_data


# =============================================
# ヒートマップ描画（CW / EW 共通）
# =============================================

def build_heatmap(
    month_data: list,
    value_col: str,
    title: str,
) -> go.Figure | None:

    records = []
    for dp in month_data:
        df = dp['sector_rs_df']
        if df is None or df.empty or value_col not in df.columns:
            continue
        tmp = df[['Sector', value_col]].copy()
        tmp['Date'] = dp['display_date']
        records.append(tmp)

    if not records:
        return None

    ts_df = pd.concat(records, ignore_index=True)

    # RS値のピボット（色付け用）
    pivot_val = ts_df.pivot_table(
        index='Sector',
        columns='Date',
        values=value_col,
        aggfunc='first'
    )

    # ランクのピボット（セル内テキスト用）
    pivot_rank = pivot_val.rank(axis=0, ascending=False, method='min').astype(int)

    # 最新日のランクで降順ソート
    # → Plotlyは下から上に描画するため、降順にすると画面上は上が最下位・下が1位になる
    # → さらに autorange='reversed' でY軸を反転させることで上が1位になる
    latest_col = pivot_val.columns[-1]
    sort_order = pivot_rank[latest_col].sort_values(ascending=False).index
    pivot_val  = pivot_val.loc[sort_order]
    pivot_rank = pivot_rank.loc[sort_order]

    x_labels  = [d.strftime('%m/%d') for d in pivot_val.columns]
    y_labels  = pivot_val.index.tolist()
    z_vals    = pivot_val.values   # 色はRS値
    text_vals = pivot_rank.values  # テキストはランク番号

    fig = go.Figure(data=go.Heatmap(
        z=z_vals,
        x=x_labels,
        y=y_labels,
        colorscale='RdYlGn',
        zmin=0,
        zmax=100,
        text=text_vals,
        texttemplate='%{text}',
        textfont={"size": 12},
        hoverongaps=False,
        colorbar=dict(
            title="RS%",
            tickmode='array',
            tickvals=[0, 25, 50, 75, 100],
            ticktext=['0', '25', '50', '75', '100'],
        ),
        hovertemplate=(
            '<b>セクター</b>: %{y}<br>'
            '<b>日付</b>: %{x}<br>'
            '<b>ランク</b>: %{text}<br>'
            f'<b>{value_col}</b>: ' + '%{z:.1f}<br>'
            '<extra></extra>'
        )
    ))

    n = len(y_labels)
    fig.update_layout(
        title=dict(text=title, font=dict(size=15)),
        xaxis=dict(title="日付", side='bottom', tickangle=-30),
        yaxis=dict(
            title="セクター",
            autorange='reversed',  # Y軸を反転 → 上が1位（数字が小さい順）
        ),
        height=max(400, n * 52 + 160),
        margin=dict(l=200, r=60, t=70, b=80),
        font=dict(size=11),
    )
    return fig


# =============================================
# 比較表生成
# =============================================

def build_latest_table(latest_df: pd.DataFrame) -> pd.DataFrame:
    if latest_df is None or latest_df.empty:
        return pd.DataFrame()

    df = latest_df[['Sector', 'Sector_RS_Pct_CW', 'Sector_RS_Pct_EW']].copy()
    df = df.sort_values('Sector_RS_Pct_CW', ascending=False).reset_index(drop=True)
    df.insert(0, 'CW順位', range(1, len(df) + 1))

    ew_rank = df['Sector_RS_Pct_EW'].rank(ascending=False, method='min').astype(int)
    df.insert(3, 'EW順位', ew_rank)

    df['順位差\n(EW-CW)'] = df['CW順位'] - df['EW順位']

    df.columns = ['CW順位', 'セクター', 'RS%（CW）', 'EW順位', 'RS%（EW）', '順位差\n(EW-CW)']
    return df


# =============================================
# メイン UI
# =============================================

with st.spinner("データ読み込み中..."):
    all_data = load_all_data(DATA_FOLDER)

if not all_data:
    st.error("読み込めるデータがありません。")
    st.stop()

all_data.sort(key=lambda x: x['date'])
latest = all_data[-1]

# ── マーケット状況バナー ──────────────────────────────────
if latest['market_summary']:
    ms     = latest['market_summary']
    status = str(ms.get('status', ''))
    score  = str(ms.get('score',  ''))
    color_map = {
        'Strong Positive': ("#d4edda", "#155724", "🟢"),
        'Positive':        ("#e8f5e9", "#2e7d32", "🟢"),
        'Neutral':         ("#fff3cd", "#856404", "🟡"),
        'Negative':        ("#ffebee", "#c62828", "🔴"),
        'Strong Negative': ("#f8d7da", "#721c24", "🔴"),
    }
    bg, fg, emoji = color_map.get(status, ("#f0f0f0", "#333", "⚪"))
    st.markdown(
        f"""
        <div style="padding:16px;border-radius:10px;background:{bg};
                    border-left:6px solid {fg};margin-bottom:20px;">
            <h3 style="margin:0;color:{fg};">{emoji} マーケット状況: {status}</h3>
            <p style="margin:4px 0 0;color:{fg};font-size:15px;">スコア率: {score}</p>
        </div>
        """,
        unsafe_allow_html=True
    )

st.markdown("---")

# ── 月選択 ───────────────────────────────────────────────
st.header("📊 セクター RS 推移（CW / EW）")

available_months = get_available_months(all_data)

col_sel, _ = st.columns([2, 8])
with col_sel:
    selected_month = st.selectbox(
        "表示する月を選択",
        options=available_months,
        index=0,
        key="sector_month"
    )

month_data = filter_data_by_month(all_data, selected_month)
st.caption(f"📅 {selected_month} のデータ: {len(month_data)} 日分")

# ── タブ ─────────────────────────────────────────────────
tab_cw, tab_ew, tab_compare = st.tabs([
    "📈 Cap Weight（CW）",
    "⚖️ Equal Weight（EW）",
    "🔀 CW / EW 比較表",
])

with tab_cw:
    if len(month_data) >= 2:
        fig_cw = build_heatmap(
            month_data,
            value_col='Sector_RS_Pct_CW',
            title=f"セクター RS_Pct_CW 推移 ― {selected_month}",
        )
        if fig_cw:
            st.plotly_chart(fig_cw, use_container_width=True)
        else:
            st.info("CW データが不足しています。")
    else:
        st.info("データが1日分しかありません（最低2日分必要）。")

with tab_ew:
    if len(month_data) >= 2:
        fig_ew = build_heatmap(
            month_data,
            value_col='Sector_RS_Pct_EW',
            title=f"セクター RS_Pct_EW 推移 ― {selected_month}",
        )
        if fig_ew:
            st.plotly_chart(fig_ew, use_container_width=True)
        else:
            st.info("EW データが不足しています。")
    else:
        st.info("データが1日分しかありません（最低2日分必要）。")

with tab_compare:
    st.subheader("📋 最新セクター CW / EW ランキング比較")
    st.caption("CW順位で降順ソート。EW順位はEW値に基づく独立したランクです。")

    latest_df  = latest['sector_rs_df']
    compare_df = build_latest_table(latest_df)

    if not compare_df.empty:
        styled = (
            compare_df.style
            .apply(color_rs_col, subset=['RS%（CW）'])
            .apply(color_rs_col, subset=['RS%（EW）'])
            .apply(color_diff_col, subset=['順位差\n(EW-CW)'])
            .format({
                'RS%（CW）':       '{:.0f}',
                'RS%（EW）':       '{:.0f}',
                '順位差\n(EW-CW)': '{:+d}',
            })
        )
        st.dataframe(styled, use_container_width=True, hide_index=True, height=450)

        st.markdown("""
        **順位差（EW－CW）の見方：**
        - 🟢 **プラス（緑）**: EWのほうが上位 → 中小型株が大型株より強い
        - 🔴 **マイナス（赤）**: CWのほうが上位 → 大型株が中小型株より強い
        """)
    else:
        st.info("最新ファイルに Screening_Results シートが見つかりません。")

gc.collect()
