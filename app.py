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
    Screening_Results シートから
      Sector / Sector_RS_Pct_CW / Sector_RS_Pct_EW
    を抽出。セクターごとに first() で代表値を取得する。
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
            parts = filename.split('_')
            date_str = next((p for p in parts if len(p) == 8 and p.isdigit()), None)
            date = datetime.strptime(date_str, '%Y%m%d') if date_str else datetime.now()

            with pd.ExcelFile(file_path) as excel:

                # ── Screening_Results ──────────────────────────────
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
    value_col: str,          # 'Sector_RS_Pct_CW' or 'Sector_RS_Pct_EW'
    title: str,
) -> go.Figure | None:
    """
    value_col の実値（0〜100）をそのまま色スケールに使うヒートマップ。
    最新日の値で降順ソート（値が高いセクターを上段に表示）。
    """
    records = []
    for dp in month_data:
        df = dp['sector_rs_df']
        if df is None or df.empty:
            continue
        if value_col not in df.columns:
            continue
        tmp = df[['Sector', value_col]].copy()
        tmp['Date'] = dp['display_date']
        records.append(tmp)

    if not records:
        return None

    ts_df = pd.concat(records, ignore_index=True)

    pivot = ts_df.pivot_table(
        index='Sector',
        columns='Date',
        values=value_col,
        aggfunc='first'
    )

    # 最新日の値で降順ソート
    latest_col = pivot.columns[-1]
    pivot = pivot.sort_values(by=latest_col, ascending=False)

    x_labels = [d.strftime('%m/%d') for d in pivot.columns]
    y_labels  = pivot.index.tolist()
    z_vals    = pivot.values

    fig = go.Figure(data=go.Heatmap(
        z=z_vals,
        x=x_labels,
        y=y_labels,
        colorscale='RdYlGn',
        zmin=0,
        zmax=100,
        text=z_vals,
        texttemplate='%{text:.0f}',
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
            f'<b>{value_col}</b>: ' + '%{z:.1f}<br>'
            '<extra></extra>'
        )
    ))

    n = len(y_labels)
    fig.update_layout(
        title=dict(text=title, font=dict(size=15)),
        xaxis=dict(title="日付", side='bottom', tickangle=-30),
        yaxis=dict(title="セクター"),
        height=max(400, n * 52 + 160),
        margin=dict(l=200, r=60, t=70, b=80),
        font=dict(size=11),
    )
    return fig


# =============================================
# 最新ランキング表（CW / EW 並列）
# =============================================

def build_latest_table(latest_df: pd.DataFrame) -> pd.DataFrame:
    """
    CW / EW 両方の値を含む最新ランキング表を生成する。
    CW の値で降順ソートし、EW のランクを別列で付与する。
    """
    if latest_df is None or latest_df.empty:
        return pd.DataFrame()

    df = latest_df[['Sector', 'Sector_RS_Pct_CW', 'Sector_RS_Pct_EW']].copy()

    # CW ランク
    df = df.sort_values('Sector_RS_Pct_CW', ascending=False).reset_index(drop=True)
    df.insert(0, 'CW Rank', range(1, len(df) + 1))

    # EW ランク（値が大きいほど上位）
    ew_rank = df['Sector_RS_Pct_EW'].rank(ascending=False, method='min').astype(int)
    df.insert(3, 'EW Rank', ew_rank)

    df.columns = ['CW順位', 'セクター', 'RS%（CW）', 'EW順位', 'RS%（EW）']
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

# ── 月選択（CW / EW 共通） ───────────────────────────────
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

# ── CW / EW タブ ────────────────────────────────────────
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

    latest_df = latest['sector_rs_df']
    compare_df = build_latest_table(latest_df)

    if not compare_df.empty:

        # CW順位とEW順位の差分を追加（正=CWのほうが上位、負=EWのほうが上位）
        compare_df['順位差(CW-EW)'] = compare_df['EW順位'] - compare_df['CW順位']

        def color_diff(val):
            """順位差をカラーで強調"""
            if val > 0:
                # EWのほうが上位（CWに比べて小型株が強い）
                return 'color: #2e7d32; font-weight: bold'
            elif val < 0:
                # CWのほうが上位（大型株が強い）
                return 'color: #c62828; font-weight: bold'
            return ''

        styled = (
            compare_df.style
            .applymap(color_diff, subset=['順位差(CW-EW)'])
            .format({
                'RS%（CW）': '{:.0f}',
                'RS%（EW）': '{:.0f}',
                '順位差(CW-EW)': '{:+d}',
            })
            .background_gradient(
                subset=['RS%（CW）', 'RS%（EW）'],
                cmap='RdYlGn',
                vmin=0, vmax=100
            )
        )
        st.dataframe(styled, use_container_width=True, hide_index=True, height=450)

        st.markdown("""
        **順位差（CW－EW）の見方：**
        - 🟢 **プラス（緑）**: EWのほうが上位 → 中小型株が大型株より強い
        - 🔴 **マイナス（赤）**: CWのほうが上位 → 大型株が中小型株より強い
        """)
    else:
        st.info("最新ファイルに Screening_Results シートが見つかりません。")

gc.collect()
