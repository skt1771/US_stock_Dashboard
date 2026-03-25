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
# RS値 → 背景色
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
            date_str = next(
                (p for p in parts if len(p) == 8 and p.isdigit()), None
            )
            date = (
                datetime.strptime(date_str, '%Y%m%d') if date_str
                else datetime.now()
            )

            with pd.ExcelFile(file_path) as excel:

                sector_rs_df   = None
                industry_rs_df = None
                stock_df       = None

                if 'Screening_Results' in excel.sheet_names:
                    rs_cols = [
                        'Sector',   'Sector_RS_Pct_CW',   'Sector_RS_Pct_EW',
                        'Industry', 'Industry_RS_Pct_CW', 'Industry_RS_Pct_EW',
                    ]
                    stock_cols = [
                        'Symbol', 'Company Name',
                        'Sector', 'Industry',
                        'Screening_Score', 'Technical_Score', 'Fundamental_Score',
                        'RS_Score',
                        'Individual_RS_Percentile',
                        'Sector_RS_Pct_CW', 'Sector_RS_Pct_EW',
                        'Industry_RS_Pct_CW', 'Industry_RS_Pct_EW',
                        'Current_Price', 'MA21', 'MA50', 'MA150',
                        'ATR_Pct_from_MA50', 'ADR',
                        'sales_accel_3_qtrs', 'eps_accel_3_qtrs',
                        # ★ BP カラム追加
                        'BP_Stock',
                        'BP_Sector_CW', 'BP_Sector_EW',
                        'BP_Industry_CW', 'BP_Industry_EW',
                    ]

                    raw_all = excel.parse('Screening_Results')

                    avail_rs = [c for c in rs_cols if c in raw_all.columns]
                    raw = raw_all[avail_rs].copy()

                    if 'Sector' in raw.columns:
                        sector_rs_df = (
                            raw.dropna(subset=['Sector'])
                               .groupby('Sector', as_index=False)
                               .agg(
                                   Sector_RS_Pct_CW=('Sector_RS_Pct_CW', 'first'),
                                   Sector_RS_Pct_EW=('Sector_RS_Pct_EW', 'first'),
                               )
                        )

                    if 'Industry' in raw.columns:
                        industry_rs_df = (
                            raw.dropna(subset=['Industry'])
                               .groupby('Industry', as_index=False)
                               .agg(
                                   Industry_RS_Pct_CW=('Industry_RS_Pct_CW', 'first'),
                                   Industry_RS_Pct_EW=('Industry_RS_Pct_EW', 'first'),
                               )
                        )

                    avail_stock = [c for c in stock_cols if c in raw_all.columns]
                    stock_df = raw_all[avail_stock].copy()

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
                'industry_rs_df': industry_rs_df,
                'stock_df':       stock_df,
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
# セクター用ヒートマップ
# =============================================

def build_sector_heatmap(
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

    pivot_val  = ts_df.pivot_table(
        index='Sector', columns='Date', values=value_col, aggfunc='first'
    )
    pivot_rank = pivot_val.rank(axis=0, ascending=False, method='min').astype(int)

    latest_col = pivot_val.columns[-1]
    sort_order = pivot_rank[latest_col].sort_values(ascending=True).index
    pivot_val  = pivot_val.loc[sort_order]
    pivot_rank = pivot_rank.loc[sort_order]

    x_labels = [d.strftime('%m/%d') for d in pivot_val.columns]
    y_labels = pivot_val.index.tolist()
    n = len(y_labels)

    fig = go.Figure(data=go.Heatmap(
        z=pivot_val.values,
        x=x_labels,
        y=y_labels,
        colorscale='RdYlGn',
        zmin=0, zmax=100,
        text=pivot_rank.values,
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
    fig.update_layout(
        title=dict(text=title, font=dict(size=15)),
        xaxis=dict(title="日付", side='bottom', tickangle=-30),
        yaxis=dict(title="セクター", autorange='reversed'),
        height=max(400, n * 52 + 160),
        margin=dict(l=200, r=60, t=70, b=80),
        font=dict(size=11),
    )
    return fig


# =============================================
# インダストリー用ヒートマップ
# =============================================

def build_industry_heatmap(
    month_data: list,
    value_col: str,
    title: str,
    top_n: int = 30,
) -> go.Figure | None:

    records = []
    for dp in month_data:
        df = dp['industry_rs_df']
        if df is None or df.empty or value_col not in df.columns:
            continue
        tmp = df[['Industry', value_col]].copy()
        tmp['Date'] = dp['display_date']
        records.append(tmp)

    if not records:
        return None

    ts_df = pd.concat(records, ignore_index=True)

    pivot_val  = ts_df.pivot_table(
        index='Industry', columns='Date', values=value_col, aggfunc='first'
    )
    pivot_rank = pivot_val.rank(axis=0, ascending=False, method='min').astype(int)

    latest_col = pivot_val.columns[-1]
    sort_order = pivot_rank[latest_col].sort_values(ascending=True).index
    pivot_val  = pivot_val.loc[sort_order].head(top_n)
    pivot_rank = pivot_rank.loc[sort_order].head(top_n)

    x_labels = [d.strftime('%m/%d') for d in pivot_val.columns]
    y_labels = pivot_val.index.tolist()
    n = len(y_labels)

    fig = go.Figure(data=go.Heatmap(
        z=pivot_val.values,
        x=x_labels,
        y=y_labels,
        colorscale='RdYlGn',
        zmin=0, zmax=100,
        text=pivot_rank.values,
        texttemplate='%{text}',
        textfont={"size": 10},
        hoverongaps=False,
        colorbar=dict(
            title="RS%",
            tickmode='array',
            tickvals=[0, 25, 50, 75, 100],
            ticktext=['0', '25', '50', '75', '100'],
        ),
        hovertemplate=(
            '<b>インダストリー</b>: %{y}<br>'
            '<b>日付</b>: %{x}<br>'
            '<b>ランク</b>: %{text}<br>'
            f'<b>{value_col}</b>: ' + '%{z:.1f}<br>'
            '<extra></extra>'
        )
    ))
    fig.update_layout(
        title=dict(text=title, font=dict(size=15)),
        xaxis=dict(title="日付", side='bottom', tickangle=-30),
        yaxis=dict(title="インダストリー", autorange='reversed'),
        height=max(500, n * 38 + 160),
        margin=dict(l=250, r=60, t=70, b=80),
        font=dict(size=10),
    )
    return fig


# =============================================
# セクター比較表
# =============================================

def build_latest_sector_table(latest_df: pd.DataFrame) -> pd.DataFrame:
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
# インダストリー比較表
# =============================================

def build_latest_industry_table(latest_df: pd.DataFrame, top_n: int = 30) -> pd.DataFrame:
    if latest_df is None or latest_df.empty:
        return pd.DataFrame()

    df = latest_df[['Industry', 'Industry_RS_Pct_CW', 'Industry_RS_Pct_EW']].copy()
    df = df.sort_values('Industry_RS_Pct_CW', ascending=False).reset_index(drop=True)
    df.insert(0, 'CW順位', range(1, len(df) + 1))

    ew_rank = df['Industry_RS_Pct_EW'].rank(ascending=False, method='min').astype(int)
    df.insert(3, 'EW順位', ew_rank)
    df['順位差\n(EW-CW)'] = df['CW順位'] - df['EW順位']

    df.columns = ['CW順位', 'インダストリー', 'RS%（CW）', 'EW順位', 'RS%（EW）', '順位差\n(EW-CW)']
    return df.head(top_n)


# =============================================
# モメンタム銘柄スクリーニング（CW or EW 単体）
# =============================================

def render_momentum_tab(
    stock_df: pd.DataFrame,
    display_date: str,
    rs_mode: str,
    tab_key: str,
):
    sector_rs_col   = f'Sector_RS_Pct_{rs_mode}'
    industry_rs_col = f'Industry_RS_Pct_{rs_mode}'
    mode_label      = "（時価総額加重: CW）" if rs_mode == "CW" else "（等加重: EW）"

    if stock_df is None or stock_df.empty:
        st.error(
            "銘柄レベルのデータが読み込めませんでした。"
            " Screening_Results シートに必要なカラムが含まれているか確認してください。"
        )
        return

    st.caption(
        f"📅 データ日付: {display_date}　　"
        f"対象銘柄数: {len(stock_df):,} 銘柄"
    )

    with st.expander("⚙️ フィルター条件を設定する", expanded=True):
        col1, col2 = st.columns(2)

        with col1:
            st.subheader("📊 テクニカル条件")
            enable_technical = st.checkbox(
                "テクニカル条件を有効にする", value=True,
                key=f"{tab_key}_enable_tech"
            )

            if enable_technical:
                st.markdown("**ATR条件**")
                atr_min = st.number_input(
                    "ATR from MA50 最小値 (%)", value=2.0, step=0.1,
                    key=f"{tab_key}_atr_min"
                )
                atr_max = st.number_input(
                    "ATR from MA50 最大値 (%)", value=5.0, step=0.1,
                    key=f"{tab_key}_atr_max"
                )
                adr_min = st.number_input(
                    "ADR 最小値 (%)", value=4.0, step=0.5,
                    key=f"{tab_key}_adr_min"
                )
                st.markdown("---")
                st.markdown("**移動平均線条件**")
                ma21_cond = st.checkbox(
                    "株価 > MA21（21日移動平均）", value=True,
                    key=f"{tab_key}_ma21"
                )
                ma50_cond = st.checkbox(
                    "株価 > MA50（50日移動平均）", value=True,
                    key=f"{tab_key}_ma50"
                )
                ma150_cond = st.checkbox(
                    "株価 > MA150（150日移動平均）", value=True,
                    key=f"{tab_key}_ma150"
                )
                ma_order_cond = st.checkbox(
                    "MA21 > MA50 > MA150（上昇トレンド）", value=True,
                    key=f"{tab_key}_ma_order"
                )
            else:
                atr_min = atr_max = adr_min = 0.0
                ma21_cond = ma50_cond = ma150_cond = ma_order_cond = False
                st.info("テクニカル条件は無効です。")

            st.markdown("---")
            st.markdown("**価格条件**")
            price_min = st.number_input(
                "株価 最小値 ($)", value=10.0, step=1.0,
                key=f"{tab_key}_price_min"
            )

            st.markdown("---")
            st.markdown("**ファンダメンタル条件**")
            enable_fundamental = st.checkbox(
                "ファンダメンタル条件を有効にする", value=False,
                key=f"{tab_key}_enable_fund"
            )
            if enable_fundamental:
                fundamental_min = st.number_input(
                    "ファンダメンタルスコア 最小値",
                    min_value=0, max_value=10, value=5, step=1,
                    key=f"{tab_key}_fund_min"
                )
            else:
                fundamental_min = 0
                st.info("ファンダメンタル条件は無効です。")

        with col2:
            st.subheader(f"📈 RS条件 {mode_label}")
            enable_rs = st.checkbox(
                "RS条件を有効にする", value=True,
                key=f"{tab_key}_enable_rs"
            )

            if enable_rs:
                individual_rs_min = st.number_input(
                    "Individual RS Percentile 最小値",
                    value=80, step=1,
                    key=f"{tab_key}_ind_rs_min"
                )
                sector_rs_min = st.number_input(
                    f"Sector RS Pct {rs_mode} 最小値",
                    value=79, step=1,
                    key=f"{tab_key}_sec_rs_min"
                )
                industry_rs_min = st.number_input(
                    f"Industry RS Pct {rs_mode} 最小値",
                    value=80, step=1,
                    key=f"{tab_key}_ind_rs_min2"
                )
            else:
                individual_rs_min = sector_rs_min = industry_rs_min = 0
                st.info("RS条件は無効です。")

            st.markdown("---")
            st.markdown("**📋 現在の設定:**")
            lines = [
                f"テクニカル条件: {'✅ 有効' if enable_technical else '❌ 無効'}",
            ]
            if enable_technical:
                lines += [
                    f"  - ATR: {atr_min}% ~ {atr_max}%",
                    f"  - ADR: {adr_min}% 以上",
                    f"  - MA21条件: {'✅' if ma21_cond else '❌'}",
                    f"  - MA50条件: {'✅' if ma50_cond else '❌'}",
                    f"  - MA150条件: {'✅' if ma150_cond else '❌'}",
                    f"  - MA順列: {'✅' if ma_order_cond else '❌'}",
                ]
            lines.append(
                f"RS条件（{rs_mode}）: {'✅ 有効' if enable_rs else '❌ 無効'}"
            )
            if enable_rs:
                lines += [
                    f"  - 個別RS: {individual_rs_min}% 以上",
                    f"  - セクターRS {rs_mode}: {sector_rs_min}% 以上",
                    f"  - 業種RS {rs_mode}: {industry_rs_min}% 以上",
                ]
            lines.append(f"価格: ${price_min} 以上")
            lines.append(
                f"ファンダメンタル: {'✅ 有効' if enable_fundamental else '❌ 無効'}"
                + (f"  ({fundamental_min}点以上)" if enable_fundamental else "")
            )
            st.info("\n".join(lines))

    # フィルタリング実行
    filtered = stock_df.copy()

    if enable_technical:
        if 'ATR_Pct_from_MA50' in filtered.columns:
            filtered = filtered[
                (filtered['ATR_Pct_from_MA50'] >= atr_min) &
                (filtered['ATR_Pct_from_MA50'] <= atr_max)
            ]
        if 'ADR' in filtered.columns:
            filtered = filtered[filtered['ADR'] >= adr_min]
        if ma21_cond and {'MA21', 'Current_Price'}.issubset(filtered.columns):
            filtered = filtered[filtered['Current_Price'] > filtered['MA21']]
        if ma50_cond and {'MA50', 'Current_Price'}.issubset(filtered.columns):
            filtered = filtered[filtered['Current_Price'] > filtered['MA50']]
        if ma150_cond and {'MA150', 'Current_Price'}.issubset(filtered.columns):
            filtered = filtered[filtered['Current_Price'] > filtered['MA150']]
        if ma_order_cond and {'MA21', 'MA50', 'MA150'}.issubset(filtered.columns):
            filtered = filtered[
                (filtered['MA21'] > filtered['MA50']) &
                (filtered['MA50'] > filtered['MA150'])
            ]

    if enable_rs:
        if 'Individual_RS_Percentile' in filtered.columns:
            filtered = filtered[
                filtered['Individual_RS_Percentile'] >= individual_rs_min
            ]
        if sector_rs_col in filtered.columns:
            filtered = filtered[filtered[sector_rs_col] >= sector_rs_min]
        if industry_rs_col in filtered.columns:
            filtered = filtered[filtered[industry_rs_col] >= industry_rs_min]

    if 'Current_Price' in filtered.columns:
        filtered = filtered[filtered['Current_Price'] >= price_min]

    if enable_fundamental and 'Fundamental_Score' in filtered.columns:
        filtered = filtered[filtered['Fundamental_Score'] >= fundamental_min]

    # 結果表示
    st.markdown("---")
    st.subheader(f"🚀 フィルタリング結果: {len(filtered)} 銘柄")

    if len(filtered) == 0:
        st.warning("⚠️ 条件に合致する銘柄がありません。条件を緩和してください。")
        return

    display_cols_ordered = [
        'Symbol', 'Company Name', 'Sector', 'Industry',
        'Screening_Score', 'Technical_Score', 'Fundamental_Score',
        'RS_Score', 'Individual_RS_Percentile',
        sector_rs_col, industry_rs_col,
        'Current_Price', 'MA21', 'MA50', 'MA150',
        'ATR_Pct_from_MA50', 'ADR',
    ]
    display_cols = [c for c in display_cols_ordered if c in filtered.columns]

    sort_key = next(
        (c for c in ['Screening_Score', 'RS_Score', 'Individual_RS_Percentile']
         if c in filtered.columns),
        display_cols[0]
    )

    st.dataframe(
        filtered[display_cols].sort_values(sort_key, ascending=False),
        use_container_width=True,
        height=600,
        hide_index=True,
    )

    with st.expander("📊 フィルタリング結果の統計"):
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            st.metric("銘柄数", len(filtered))
        with c2:
            if 'Screening_Score' in filtered.columns:
                st.metric("平均スコア", f"{filtered['Screening_Score'].mean():.1f}")
        with c3:
            if 'Individual_RS_Percentile' in filtered.columns:
                st.metric(
                    "平均個別RS",
                    f"{filtered['Individual_RS_Percentile'].mean():.1f}%"
                )
        with c4:
            if 'ADR' in filtered.columns:
                st.metric("平均ADR", f"{filtered['ADR'].mean():.1f}%")

        if 'Sector' in filtered.columns:
            st.markdown("**セクター分布:**")
            st.bar_chart(filtered['Sector'].value_counts())

    dl1, dl2 = st.columns(2)
    with dl1:
        csv = filtered[display_cols].to_csv(index=False).encode('utf-8')
        st.download_button(
            label="📥 CSVダウンロード（全データ）",
            data=csv,
            file_name=(
                f"momentum_{rs_mode.lower()}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
            ),
            mime='text/csv',
            key=f"{tab_key}_dl_csv",
        )
    with dl2:
        if 'Symbol' in filtered.columns:
            syms = (
                filtered.sort_values(sort_key, ascending=False)['Symbol']
                .dropna().astype(str).tolist()
            )
            st.download_button(
                label="📝 Symbolリスト（TXT）",
                data=','.join(syms),
                file_name=(
                    f"momentum_symbols_{rs_mode.lower()}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
                ),
                mime='text/plain',
                key=f"{tab_key}_dl_txt",
            )

    if 'Symbol' in filtered.columns:
        with st.expander("📌 Symbolリスト表示（TradingView用）"):
            syms = (
                filtered.sort_values(sort_key, ascending=False)['Symbol']
                .dropna().astype(str).tolist()
            )
            st.markdown("**カンマ区切り（コピー用）:**")
            st.code(','.join(syms), language=None)
            st.success(f"✅ 合計 {len(syms)} 銘柄")
            if len(syms) > 10:
                st.info(f"📊 上位10銘柄: {', '.join(syms[:10])}")


# =============================================
# モメンタム銘柄スクリーニング（CW + EW 両条件）
# =============================================

def render_momentum_tab_both(
    stock_df: pd.DataFrame,
    display_date: str,
    tab_key: str,
):
    if stock_df is None or stock_df.empty:
        st.error(
            "銘柄レベルのデータが読み込めませんでした。"
            " Screening_Results シートに必要なカラムが含まれているか確認してください。"
        )
        return

    st.caption(
        f"📅 データ日付: {display_date}　　"
        f"対象銘柄数: {len(stock_df):,} 銘柄"
    )

    with st.expander("⚙️ フィルター条件を設定する", expanded=True):

        # ── テクニカル条件 ───────────────────────────────────
        st.subheader("📊 テクニカル条件")
        col_t1, col_t2 = st.columns(2)

        with col_t1:
            enable_technical = st.checkbox(
                "テクニカル条件を有効にする", value=True,
                key=f"{tab_key}_enable_tech"
            )
            if enable_technical:
                st.markdown("**ATR条件**")
                atr_min = st.number_input(
                    "ATR from MA50 最小値 (%)", value=1.5, step=0.1,
                    key=f"{tab_key}_atr_min"
                )
                atr_max = st.number_input(
                    "ATR from MA50 最大値 (%)", value=6.0, step=0.1,
                    key=f"{tab_key}_atr_max"
                )
                st.markdown("---")
                st.markdown("**ADR条件**")
                adr_min = st.number_input(
                    "ADR 最小値 (%)", value=2.5, step=0.1,
                    key=f"{tab_key}_adr_min"
                )
                adr_max = st.number_input(
                    "ADR 最大値 (%)", value=5.5, step=0.1,
                    key=f"{tab_key}_adr_max"
                )
            else:
                atr_min = atr_max = adr_min = 0.0
                adr_max = float('inf')
                st.info("テクニカル条件は無効です。")

        with col_t2:
            if enable_technical:
                st.markdown("**移動平均線条件**")
                ma21_cond = st.checkbox(
                    "株価 > MA21（21日移動平均）", value=True,
                    key=f"{tab_key}_ma21"
                )
                ma50_cond = st.checkbox(
                    "株価 > MA50（50日移動平均）", value=True,
                    key=f"{tab_key}_ma50"
                )
                ma150_cond = st.checkbox(
                    "株価 > MA150（150日移動平均）", value=True,
                    key=f"{tab_key}_ma150"
                )
                ma_order_cond = st.checkbox(
                    "MA21 > MA50 > MA150（上昇トレンド）", value=True,
                    key=f"{tab_key}_ma_order"
                )
            else:
                ma21_cond = ma50_cond = ma150_cond = ma_order_cond = False

        st.markdown("---")
        col_p1, col_p2 = st.columns(2)
        with col_p1:
            st.markdown("**価格条件**")
            price_min = st.number_input(
                "株価 最小値 ($)", value=10.0, step=1.0,
                key=f"{tab_key}_price_min"
            )
        with col_p2:
            st.markdown("**ファンダメンタル条件**")
            enable_fundamental = st.checkbox(
                "ファンダメンタル条件を有効にする", value=False,
                key=f"{tab_key}_enable_fund"
            )
            if enable_fundamental:
                fundamental_min = st.number_input(
                    "ファンダメンタルスコア 最小値",
                    min_value=0, max_value=10, value=5, step=1,
                    key=f"{tab_key}_fund_min"
                )
            else:
                fundamental_min = 0
                st.info("ファンダメンタル条件は無効です。")

        st.markdown("---")

        # ── RS条件（CW / EW 横並び） ─────────────────────────
        st.subheader("📈 RS条件（CW・EW 両方）")
        col_cw, col_ew = st.columns(2)

        with col_cw:
            st.markdown("**🔵 CW（時価総額加重）**")
            enable_rs_cw = st.checkbox(
                "CW RS条件を有効にする", value=True,
                key=f"{tab_key}_enable_rs_cw"
            )
            if enable_rs_cw:
                individual_rs_min = st.number_input(
                    "Individual RS Percentile 最小値",
                    value=80, step=1,
                    key=f"{tab_key}_ind_rs_min"
                )
                sector_rs_cw_min = st.number_input(
                    "Sector RS Pct CW 最小値",
                    value=70, step=1,
                    key=f"{tab_key}_sec_rs_cw_min"
                )
                industry_rs_cw_min = st.number_input(
                    "Industry RS Pct CW 最小値",
                    value=70, step=1,
                    key=f"{tab_key}_ind_rs_cw_min"
                )
            else:
                individual_rs_min = sector_rs_cw_min = industry_rs_cw_min = 0
                st.info("CW RS条件は無効です。")

        with col_ew:
            st.markdown("**🟠 EW（等加重）**")
            enable_rs_ew = st.checkbox(
                "EW RS条件を有効にする", value=True,
                key=f"{tab_key}_enable_rs_ew"
            )
            if enable_rs_ew:
                sector_rs_ew_min = st.number_input(
                    "Sector RS Pct EW 最小値",
                    value=70, step=1,
                    key=f"{tab_key}_sec_rs_ew_min"
                )
                industry_rs_ew_min = st.number_input(
                    "Industry RS Pct EW 最小値",
                    value=70, step=1,
                    key=f"{tab_key}_ind_rs_ew_min"
                )
            else:
                sector_rs_ew_min = industry_rs_ew_min = 0
                st.info("EW RS条件は無効です。")

        st.markdown("---")

        # ── バイプレッシャー条件 ★ 新規 ─────────────────────
        st.subheader("💹 バイプレッシャー条件")
        enable_bp = st.checkbox(
            "バイプレッシャー条件を有効にする", value=True,
            key=f"{tab_key}_enable_bp"
        )
        if enable_bp:
            col_bp1, col_bp2 = st.columns(2)
            with col_bp1:
                bp_stock_min = st.number_input(
                    "BP_Stock 最小値", value=0.6, step=0.05,
                    format="%.2f", key=f"{tab_key}_bp_stock_min"
                )
                bp_sector_cw_min = st.number_input(
                    "BP_Sector_CW 最小値", value=0.6, step=0.05,
                    format="%.2f", key=f"{tab_key}_bp_sec_cw_min"
                )
                bp_sector_ew_min = st.number_input(
                    "BP_Sector_EW 最小値", value=0.6, step=0.05,
                    format="%.2f", key=f"{tab_key}_bp_sec_ew_min"
                )
            with col_bp2:
                bp_industry_cw_min = st.number_input(
                    "BP_Industry_CW 最小値", value=0.6, step=0.05,
                    format="%.2f", key=f"{tab_key}_bp_ind_cw_min"
                )
                bp_industry_ew_min = st.number_input(
                    "BP_Industry_EW 最小値", value=0.6, step=0.05,
                    format="%.2f", key=f"{tab_key}_bp_ind_ew_min"
                )
        else:
            bp_stock_min = bp_sector_cw_min = bp_sector_ew_min = 0.0
            bp_industry_cw_min = bp_industry_ew_min = 0.0
            st.info("バイプレッシャー条件は無効です。")

        # 設定サマリー
        st.markdown("---")
        st.markdown("**📋 現在の設定:**")
        lines = [
            f"テクニカル条件: {'✅ 有効' if enable_technical else '❌ 無効'}",
        ]
        if enable_technical:
            lines += [
                f"  - ATR: {atr_min}% ~ {atr_max}%",
                f"  - ADR: {adr_min}% ~ {adr_max}%",
                f"  - MA21条件: {'✅' if ma21_cond else '❌'}",
                f"  - MA50条件: {'✅' if ma50_cond else '❌'}",
                f"  - MA150条件: {'✅' if ma150_cond else '❌'}",
                f"  - MA順列: {'✅' if ma_order_cond else '❌'}",
            ]
        lines.append(f"価格: ${price_min} 以上")
        lines.append(
            f"ファンダメンタル: {'✅ 有効' if enable_fundamental else '❌ 無効'}"
            + (f"  ({fundamental_min}点以上)" if enable_fundamental else "")
        )
        lines.append(
            f"CW RS条件: {'✅ 有効' if enable_rs_cw else '❌ 無効'}"
        )
        if enable_rs_cw:
            lines += [
                f"  - 個別RS: {individual_rs_min}% 以上",
                f"  - セクターRS CW: {sector_rs_cw_min}% 以上",
                f"  - 業種RS CW: {industry_rs_cw_min}% 以上",
            ]
        lines.append(
            f"EW RS条件: {'✅ 有効' if enable_rs_ew else '❌ 無効'}"
        )
        if enable_rs_ew:
            lines += [
                f"  - セクターRS EW: {sector_rs_ew_min}% 以上",
                f"  - 業種RS EW: {industry_rs_ew_min}% 以上",
            ]
        lines.append(
            f"バイプレッシャー条件: {'✅ 有効' if enable_bp else '❌ 無効'}"
        )
        if enable_bp:
            lines += [
                f"  - BP_Stock: {bp_stock_min:.2f} 以上",
                f"  - BP_Sector_CW: {bp_sector_cw_min:.2f} 以上",
                f"  - BP_Sector_EW: {bp_sector_ew_min:.2f} 以上",
                f"  - BP_Industry_CW: {bp_industry_cw_min:.2f} 以上",
                f"  - BP_Industry_EW: {bp_industry_ew_min:.2f} 以上",
            ]
        st.info("\n".join(lines))

    # ── フィルタリング実行 ────────────────────────────────
    filtered = stock_df.copy()

    if enable_technical:
        if 'ATR_Pct_from_MA50' in filtered.columns:
            filtered = filtered[
                (filtered['ATR_Pct_from_MA50'] >= atr_min) &
                (filtered['ATR_Pct_from_MA50'] <= atr_max)
            ]
        if 'ADR' in filtered.columns:
            filtered = filtered[
                (filtered['ADR'] >= adr_min) &
                (filtered['ADR'] <= adr_max)
            ]
        if ma21_cond and {'MA21', 'Current_Price'}.issubset(filtered.columns):
            filtered = filtered[filtered['Current_Price'] > filtered['MA21']]
        if ma50_cond and {'MA50', 'Current_Price'}.issubset(filtered.columns):
            filtered = filtered[filtered['Current_Price'] > filtered['MA50']]
        if ma150_cond and {'MA150', 'Current_Price'}.issubset(filtered.columns):
            filtered = filtered[filtered['Current_Price'] > filtered['MA150']]
        if ma_order_cond and {'MA21', 'MA50', 'MA150'}.issubset(filtered.columns):
            filtered = filtered[
                (filtered['MA21'] > filtered['MA50']) &
                (filtered['MA50'] > filtered['MA150'])
            ]

    if 'Current_Price' in filtered.columns:
        filtered = filtered[filtered['Current_Price'] >= price_min]

    if enable_fundamental and 'Fundamental_Score' in filtered.columns:
        filtered = filtered[filtered['Fundamental_Score'] >= fundamental_min]

    # CW RS条件
    if enable_rs_cw:
        if 'Individual_RS_Percentile' in filtered.columns:
            filtered = filtered[
                filtered['Individual_RS_Percentile'] >= individual_rs_min
            ]
        if 'Sector_RS_Pct_CW' in filtered.columns:
            filtered = filtered[filtered['Sector_RS_Pct_CW'] >= sector_rs_cw_min]
        if 'Industry_RS_Pct_CW' in filtered.columns:
            filtered = filtered[filtered['Industry_RS_Pct_CW'] >= industry_rs_cw_min]

    # EW RS条件
    if enable_rs_ew:
        if 'Sector_RS_Pct_EW' in filtered.columns:
            filtered = filtered[filtered['Sector_RS_Pct_EW'] >= sector_rs_ew_min]
        if 'Industry_RS_Pct_EW' in filtered.columns:
            filtered = filtered[filtered['Industry_RS_Pct_EW'] >= industry_rs_ew_min]

    # ★ バイプレッシャー条件
    if enable_bp:
        if 'BP_Stock' in filtered.columns:
            filtered = filtered[filtered['BP_Stock'] >= bp_stock_min]
        if 'BP_Sector_CW' in filtered.columns:
            filtered = filtered[filtered['BP_Sector_CW'] >= bp_sector_cw_min]
        if 'BP_Sector_EW' in filtered.columns:
            filtered = filtered[filtered['BP_Sector_EW'] >= bp_sector_ew_min]
        if 'BP_Industry_CW' in filtered.columns:
            filtered = filtered[filtered['BP_Industry_CW'] >= bp_industry_cw_min]
        if 'BP_Industry_EW' in filtered.columns:
            filtered = filtered[filtered['BP_Industry_EW'] >= bp_industry_ew_min]

    # ── 結果表示 ──────────────────────────────────────────
    st.markdown("---")
    st.subheader(f"🚀 フィルタリング結果: {len(filtered)} 銘柄")

    if len(filtered) == 0:
        st.warning("⚠️ 条件に合致する銘柄がありません。条件を緩和してください。")
        return

    display_cols_ordered = [
        'Symbol', 'Company Name', 'Sector', 'Industry',
        'Screening_Score', 'Technical_Score', 'Fundamental_Score',
        'RS_Score', 'Individual_RS_Percentile',
        'Sector_RS_Pct_CW',  'Sector_RS_Pct_EW',
        'Industry_RS_Pct_CW', 'Industry_RS_Pct_EW',
        'Current_Price', 'MA21', 'MA50', 'MA150',
        'ATR_Pct_from_MA50', 'ADR',
        # ★ BP カラムも表示
        'BP_Stock',
        'BP_Sector_CW', 'BP_Sector_EW',
        'BP_Industry_CW', 'BP_Industry_EW',
    ]
    display_cols = [c for c in display_cols_ordered if c in filtered.columns]

    sort_key = next(
        (c for c in ['Screening_Score', 'RS_Score', 'Individual_RS_Percentile']
         if c in filtered.columns),
        display_cols[0]
    )

    st.dataframe(
        filtered[display_cols].sort_values(sort_key, ascending=False),
        use_container_width=True,
        height=600,
        hide_index=True,
    )

    # ── 統計サマリー ──────────────────────────────────────
    with st.expander("📊 フィルタリング結果の統計"):
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            st.metric("銘柄数", len(filtered))
        with c2:
            if 'Screening_Score' in filtered.columns:
                st.metric("平均スコア", f"{filtered['Screening_Score'].mean():.1f}")
        with c3:
            if 'Individual_RS_Percentile' in filtered.columns:
                st.metric(
                    "平均個別RS",
                    f"{filtered['Individual_RS_Percentile'].mean():.1f}%"
                )
        with c4:
            if 'ADR' in filtered.columns:
                st.metric("平均ADR", f"{filtered['ADR'].mean():.1f}%")

        if 'Sector' in filtered.columns:
            st.markdown("**セクター分布:**")
            st.bar_chart(filtered['Sector'].value_counts())

    # ── ダウンロード ──────────────────────────────────────
    dl1, dl2 = st.columns(2)
    with dl1:
        csv = filtered[display_cols].to_csv(index=False).encode('utf-8')
        st.download_button(
            label="📥 CSVダウンロード（全データ）",
            data=csv,
            file_name=(
                f"momentum_both_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
            ),
            mime='text/csv',
            key=f"{tab_key}_dl_csv",
        )
    with dl2:
        if 'Symbol' in filtered.columns:
            syms = (
                filtered.sort_values(sort_key, ascending=False)['Symbol']
                .dropna().astype(str).tolist()
            )
            st.download_button(
                label="📝 Symbolリスト（TXT）",
                data=','.join(syms),
                file_name=(
                    f"momentum_symbols_both_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
                ),
                mime='text/plain',
                key=f"{tab_key}_dl_txt",
            )

    # ── TradingView 用コピー ──────────────────────────────
    if 'Symbol' in filtered.columns:
        with st.expander("📌 Symbolリスト表示（TradingView用）"):
            syms = (
                filtered.sort_values(sort_key, ascending=False)['Symbol']
                .dropna().astype(str).tolist()
            )
            st.markdown("**カンマ区切り（コピー用）:**")
            st.code(','.join(syms), language=None)
            st.success(f"✅ 合計 {len(syms)} 銘柄")
            if len(syms) > 10:
                st.info(f"📊 上位10銘柄: {', '.join(syms[:10])}")


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
latest_stock_df  = latest.get('stock_df')
latest_disp_date = latest['display_date'].strftime('%Y年%m月%d日')

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
st.header("📊 RS 推移ダッシュボード")

available_months = get_available_months(all_data)

col_sel, _ = st.columns([2, 8])
with col_sel:
    selected_month = st.selectbox(
        "表示する月を選択",
        options=available_months,
        index=0,
        key="rs_month"
    )

month_data = filter_data_by_month(all_data, selected_month)
st.caption(f"📅 {selected_month} のデータ: {len(month_data)} 日分")

# ── タブ定義 ─────────────────────────────────────────────
(
    tab_sec_cw,
    tab_sec_ew,
    tab_ind_cw,
    tab_ind_ew,
    tab_sec_compare,
    tab_ind_compare,
    tab_momentum_cw,
    tab_momentum_ew,
    tab_momentum_both,
) = st.tabs([
    "📈 セクター CW",
    "⚖️ セクター EW",
    "🏭 インダストリー CW",
    "🏭 インダストリー EW",
    "🔀 セクター CW/EW 比較",
    "🔀 インダストリー CW/EW 比較",
    "🚀 モメンタム銘柄 CW",
    "⚖️ モメンタム銘柄 EW",
    "🎯 モメンタム銘柄 CW＋EW",
])

# ---- セクター CW ----------------------------------------
with tab_sec_cw:
    if len(month_data) >= 2:
        fig = build_sector_heatmap(
            month_data,
            value_col='Sector_RS_Pct_CW',
            title=f"セクター RS_Pct_CW 推移 ― {selected_month}",
        )
        if fig:
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("CW データが不足しています。")
    else:
        st.info("データが1日分しかありません（最低2日分必要）。")

# ---- セクター EW ----------------------------------------
with tab_sec_ew:
    if len(month_data) >= 2:
        fig = build_sector_heatmap(
            month_data,
            value_col='Sector_RS_Pct_EW',
            title=f"セクター RS_Pct_EW 推移 ― {selected_month}",
        )
        if fig:
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("EW データが不足しています。")
    else:
        st.info("データが1日分しかありません（最低2日分必要）。")

# ---- インダストリー CW ----------------------------------
with tab_ind_cw:
    if len(month_data) >= 2:
        col_slider, _ = st.columns([2, 8])
        with col_slider:
            top_n_cw = st.slider(
                "表示するインダストリー数（上位）",
                min_value=10, max_value=50, value=30, step=5,
                key="industry_top_n_cw"
            )
        fig = build_industry_heatmap(
            month_data,
            value_col='Industry_RS_Pct_CW',
            title=f"インダストリー RS_Pct_CW 推移（上位{top_n_cw}） ― {selected_month}",
            top_n=top_n_cw,
        )
        if fig:
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("インダストリー CW データが不足しています。")
    else:
        st.info("データが1日分しかありません（最低2日分必要）。")

# ---- インダストリー EW ----------------------------------
with tab_ind_ew:
    if len(month_data) >= 2:
        col_slider, _ = st.columns([2, 8])
        with col_slider:
            top_n_ew = st.slider(
                "表示するインダストリー数（上位）",
                min_value=10, max_value=50, value=30, step=5,
                key="industry_top_n_ew"
            )
        fig = build_industry_heatmap(
            month_data,
            value_col='Industry_RS_Pct_EW',
            title=f"インダストリー RS_Pct_EW 推移（上位{top_n_ew}） ― {selected_month}",
            top_n=top_n_ew,
        )
        if fig:
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("インダストリー EW データが不足しています。")
    else:
        st.info("データが1日分しかありません（最低2日分必要）。")

# ---- セクター CW/EW 比較表 ------------------------------
with tab_sec_compare:
    st.subheader("📋 最新セクター CW / EW ランキング比較")
    st.caption("CW順位で昇順ソート。EW順位はEW値に基づく独立したランクです。")

    compare_df = build_latest_sector_table(latest['sector_rs_df'])

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

# ---- インダストリー CW/EW 比較表 ------------------------
with tab_ind_compare:
    st.subheader("📋 最新インダストリー CW / EW ランキング比較")
    st.caption("CW順位で昇順ソート。EW順位はEW値に基づく独立したランクです。")

    col_slider, _ = st.columns([2, 8])
    with col_slider:
        top_n_compare = st.slider(
            "表示するインダストリー数（上位）",
            min_value=10, max_value=146, value=30, step=5,
            key="industry_compare_top_n"
        )

    ind_compare_df = build_latest_industry_table(
        latest['industry_rs_df'],
        top_n=top_n_compare
    )

    if not ind_compare_df.empty:
        styled = (
            ind_compare_df.style
            .apply(color_rs_col, subset=['RS%（CW）'])
            .apply(color_rs_col, subset=['RS%（EW）'])
            .apply(color_diff_col, subset=['順位差\n(EW-CW)'])
            .format({
                'RS%（CW）':       '{:.0f}',
                'RS%（EW）':       '{:.0f}',
                '順位差\n(EW-CW)': '{:+d}',
            })
        )
        st.dataframe(styled, use_container_width=True, hide_index=True, height=600)
        st.markdown("""
        **順位差（EW－CW）の見方：**
        - 🟢 **プラス（緑）**: EWのほうが上位 → 中小型株が大型株より強い
        - 🔴 **マイナス（赤）**: CWのほうが上位 → 大型株が中小型株より強い
        """)
    else:
        st.info("最新ファイルに Screening_Results シートが見つかりません。")

# ---- モメンタム銘柄 CW ----------------------------------
with tab_momentum_cw:
    st.header("🚀 モメンタム銘柄スクリーニング（時価総額加重: CW）")
    render_momentum_tab(
        stock_df=latest_stock_df,
        display_date=latest_disp_date,
        rs_mode='CW',
        tab_key='mom_cw',
    )

# ---- モメンタム銘柄 EW ----------------------------------
with tab_momentum_ew:
    st.header("⚖️ モメンタム銘柄スクリーニング（等加重: EW）")
    st.info(
        "💡 **EW（Equal Weight）モード**: セクター・インダストリーのRS%を"
        "時価総額に関係なく等加重で算出した値でフィルタリングします。"
        " 中小型株が強い相場環境の発掘に有効です。"
    )
    render_momentum_tab(
        stock_df=latest_stock_df,
        display_date=latest_disp_date,
        rs_mode='EW',
        tab_key='mom_ew',
    )

# ---- モメンタム銘柄 CW＋EW ------------------------------
with tab_momentum_both:
    st.header("🎯 モメンタム銘柄スクリーニング（CW＋EW 両条件）")
    st.info(
        "💡 **CW＋EW モード**: 時価総額加重（CW）と等加重（EW）の"
        "両方のRS条件を同時に満たす銘柄を抽出します。"
        " 大型株・中小型株いずれの視点でも強いセクター・インダストリーに属する"
        "銘柄を厳選したい場合に活用してください。"
    )
    render_momentum_tab_both(
        stock_df=latest_stock_df,
        display_date=latest_disp_date,
        tab_key='mom_both',
    )

gc.collect()
