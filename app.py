import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from datetime import datetime, timedelta
import os
import glob
import gc

# ページ設定
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
    """スクリーニング実行日の前日を表示日とする"""
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
# データ読み込み
# =============================================

@st.cache_data(ttl=300)
def load_all_data(data_folder: str = DATA_FOLDER) -> list:
    """
    Screening_Results シートから Sector と Sector_RS_Pct_CW を抽出。
    セクターごとに値は同一なので first() で代表値を取得する。
    """
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
            # ファイル名から日付を抽出（例: ..._20260323_...xlsx）
            parts = filename.split('_')
            date_str = next((p for p in parts if len(p) == 8 and p.isdigit()), None)
            date = datetime.strptime(date_str, '%Y%m%d') if date_str else datetime.now()

            with pd.ExcelFile(file_path) as excel:

                # ── Screening_Results ──────────────────────────────
                sector_rs_df = None
                if 'Screening_Results' in excel.sheet_names:
                    raw = excel.parse(
                        'Screening_Results',
                        usecols=['Sector', 'Sector_RS_Pct_CW']
                    )
                    # セクターごとに Sector_RS_Pct_CW の代表値（first）を取得
                    sector_rs_df = (
                        raw.dropna(subset=['Sector', 'Sector_RS_Pct_CW'])
                           .groupby('Sector', as_index=False)['Sector_RS_Pct_CW']
                           .first()
                    )

                # ── Market_Summary ─────────────────────────────────
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
                'sector_rs_df':   sector_rs_df,   # Sector / Sector_RS_Pct_CW
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
# ヒートマップ描画
# =============================================

def build_heatmap(month_data: list, selected_month: str) -> go.Figure | None:
    """
    各日付の Sector_RS_Pct_CW 値をそのまま使ってヒートマップを描画する。
    セル内表示 : Sector_RS_Pct_CW の実値（例: 63, 99, 81）
    セルの色   : 値が高いほど緑、低いほど赤
    """
    records = []
    for dp in month_data:
        df = dp['sector_rs_df']
        if df is None or df.empty:
            continue
        df = df.copy()
        df['Date'] = dp['display_date']
        records.append(df)

    if not records:
        return None

    ts_df = pd.concat(records, ignore_index=True)

    # ピボット: 行=Sector, 列=Date, 値=Sector_RS_Pct_CW
    pivot = ts_df.pivot_table(
        index='Sector',
        columns='Date',
        values='Sector_RS_Pct_CW',
        aggfunc='first'
    )

    # 最新日の値で降順ソート（RS値が高いセクターを上に）
    latest_col = pivot.columns[-1]
    pivot = pivot.sort_values(by=latest_col, ascending=False)

    x_labels = [d.strftime('%m/%d') for d in pivot.columns]
    y_labels  = pivot.index.tolist()
    z_vals    = pivot.values                          # 実値をそのまま色に使う
    text_vals = pivot.values.astype(float)            # セル内テキスト

    fig = go.Figure(data=go.Heatmap(
        z=z_vals,
        x=x_labels,
        y=y_labels,
        colorscale='RdYlGn',
        zmin=0,
        zmax=100,
        text=text_vals,
        texttemplate='%{text:.0f}',   # 小数なし整数表示
        textfont={"size": 12},
        hoverongaps=False,
        colorbar=dict(
            title="RS%（CW）",
            tickmode='array',
            tickvals=[0, 25, 50, 75, 100],
            ticktext=['0', '25', '50', '75', '100'],
        ),
        hovertemplate=(
            '<b>セクター</b>: %{y}<br>'
            '<b>日付</b>: %{x}<br>'
            '<b>RS_Pct_CW</b>: %{z:.1f}<br>'
            '<extra></extra>'
        )
    ))

    n_sectors = len(y_labels)
    fig.update_layout(
        title=dict(
            text=f"セクター Sector_RS_Pct_CW 推移 ― {selected_month}",
            font=dict(size=16)
        ),
        xaxis=dict(title="日付", side='bottom', tickangle=-30),
        yaxis=dict(title="セクター"),
        height=max(400, n_sectors * 52 + 160),
        margin=dict(l=200, r=80, t=80, b=80),
        font=dict(size=11),
    )
    return fig


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

# ── セクターRS_CW 推移 ────────────────────────────────────
st.header("📊 セクター RS_Pct_CW 推移")

available_months = get_available_months(all_data)

col_sel, _ = st.columns([2, 8])
with col_sel:
    selected_month = st.selectbox(
        "表示する月を選択",
        options=available_months,
        index=0,
        key="sector_cw_month"
    )

month_data = filter_data_by_month(all_data, selected_month)

if len(month_data) >= 2:
    fig = build_heatmap(month_data, selected_month)
    if fig:
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("ヒートマップの生成に必要なデータが不足しています。")
    st.caption(f"📅 {selected_month} のデータ: {len(month_data)} 日分")
else:
    st.info(f"{selected_month} のデータが1日分しかありません（最低2日分必要）。")

st.markdown("---")

# ── 最新セクターランキング表 ─────────────────────────────
st.subheader("📋 最新セクター RS_Pct_CW ランキング")

latest_df = latest['sector_rs_df']
if latest_df is not None and not latest_df.empty:
    disp = (
        latest_df
        .sort_values('Sector_RS_Pct_CW', ascending=False)
        .reset_index(drop=True)
    )
    disp.index += 1
    disp.index.name = 'Rank'
    disp.columns = ['セクター', 'RS%（CW）']
    st.dataframe(disp, use_container_width=True, height=430)
else:
    st.info("最新ファイルに Screening_Results シートが見つかりません。")

gc.collect()
