import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from datetime import datetime, timedelta
import os
import glob
import gc
import re

st.set_page_config(page_title="米国株RSダッシュボード", page_icon="📈", layout="wide")
DATA_FOLDER = "data"

# ── Utility ──────────────────────────────────────────────────────────────────
def get_display_date(d):       return d - timedelta(days=1)
def get_year_month(d):         return d.strftime('%Y年%m月')
def get_available_months(all_data):
    return sorted({get_year_month(e['display_date']) for e in all_data}, reverse=True)
def filter_by_month(all_data, m):
    return [e for e in all_data if get_year_month(e['display_date']) == m]

# ── Color helpers ─────────────────────────────────────────────────────────────
def rs_to_bgcolor(val):
    try: v = float(val)
    except: return ''
    v = max(0.0, min(100.0, v))
    stops = [(0,0xd7,0x30,0x27),(25,0xf4,0x6d,0x43),(50,0xff,0xff,0xbf),
             (75,0x74,0xc4,0x76),(100,0x1a,0x96,0x41)]
    for (v0,r0,g0,b0),(v1,r1,g1,b1) in zip(stops, stops[1:]):
        if v0 <= v <= v1:
            t=(v-v0)/(v1-v0)
            r,g,b=int(r0+t*(r1-r0)),int(g0+t*(g1-g0)),int(b0+t*(b1-b0))
            fg='#000000' if 0.299*r+0.587*g+0.114*b>128 else '#ffffff'
            return f'background-color:#{r:02x}{g:02x}{b:02x};color:{fg}'
    return ''

def color_rs_col(s):  return [rs_to_bgcolor(v) for v in s]
def color_diff_col(s):
    out=[]
    for v in s:
        try: i=int(v)
        except: out.append(''); continue
        out.append('color:#2e7d32;font-weight:bold' if i>0 else
                   ('color:#c62828;font-weight:bold' if i<0 else ''))
    return out

# ── Data loading ──────────────────────────────────────────────────────────────
@st.cache_data(ttl=300)
def load_all_data(folder=DATA_FOLDER):
    all_data = []
    if not os.path.exists(folder):
        st.error(f"フォルダが見つかりません: {folder}"); return all_data
    files = sorted(
        glob.glob(os.path.join(folder, "*.xlsx")) +
        glob.glob(os.path.join(folder, "*.xls"))
    )
    if not files:
        st.warning(f"{folder} にExcelファイルが見つかりません。"); return all_data

    progress = st.progress(0)
    ph = st.empty()

    for idx, fp in enumerate(files):
        fn = os.path.basename(fp)
        ph.text(f"読み込み中: {fn} ({idx+1}/{len(files)})")
        try:
            # ── ファイル名から YYYYMMDD を正規表現で抽出 ──────────────────
            match = re.search(r'(\d{8})', fn)
            if not match:
                st.warning(f"{fn}: ファイル名に8桁の日付が見つかりません。スキップします。")
                progress.progress((idx+1)/len(files))
                continue
            date_str  = match.group(1)                          # 例: "20260321"
            file_date = datetime.strptime(date_str, "%Y%m%d")
            display_date = get_display_date(file_date)

            xl = pd.ExcelFile(fp)

            # Sector RS
            sector_df = None
            for sname in ["Sector_RS","SectorRS","Sector RS","sector_rs"]:
                if sname in xl.sheet_names:
                    sector_df = xl.parse(sname); break

            # Industry RS
            industry_df = None
            for sname in ["Industry_RS","IndustryRS","Industry RS","industry_rs"]:
                if sname in xl.sheet_names:
                    industry_df = xl.parse(sname); break

            # Stock data
            stock_df = None
            stock_cols = [
                'Symbol','Company Name','Sector','Industry',
                'RS_Percentile_CW','Sector_RS_CW','Industry_RS_CW',
                'RS_Percentile_EW','Sector_RS_EW','Industry_RS_EW',
                'Price','MA21','MA50','MA150','ATR_from_50MA','ADR',
                'Fundamental_Score',
                'BP_Stock','BP_Sector_CW','BP_Sector_EW',
                'BP_Industry_CW','BP_Industry_EW'
            ]
            for sname in ["Screening_Results","ScreeningResults","Stock_Data","stock_data"]:
                if sname in xl.sheet_names:
                    raw   = xl.parse(sname)
                    avail = [c for c in stock_cols if c in raw.columns]
                    stock_df = raw[avail].copy(); break

            # Market summary
            market_summary = None
            for sname in ["Market_Summary","MarketSummary","Market Summary","market_summary"]:
                if sname in xl.sheet_names:
                    market_summary = xl.parse(sname); break

            all_data.append({
                'date':           file_date,
                'display_date':   display_date,
                'filename':       fn,
                'sector_df':      sector_df,
                'industry_df':    industry_df,
                'stock_df':       stock_df,
                'market_summary': market_summary,
            })

        except Exception as e:
            st.warning(f"{fn} の読み込みに失敗: {e}")

        progress.progress((idx+1)/len(files))

    progress.empty(); ph.empty(); gc.collect()
    return all_data

# ── Heatmap builders ──────────────────────────────────────────────────────────
def build_sector_heatmap(month_data, rs_col, title):
    if not month_data: return None
    frames = []
    for e in month_data:
        df = e.get('sector_df')
        if df is None or rs_col not in df.columns: continue
        tmp = df[['Sector', rs_col]].copy()
        tmp['date'] = e['display_date']
        frames.append(tmp)
    if not frames: return None
    merged = pd.concat(frames)
    pivot  = merged.pivot_table(index='Sector', columns='date', values=rs_col)
    pivot  = pivot.sort_values(by=pivot.columns[-1], ascending=False)
    dates  = [d.strftime('%m/%d') for d in pivot.columns]
    z      = pivot.values
    text   = [[f"{v:.1f}" if not pd.isna(v) else "" for v in row] for row in z]
    fig = go.Figure(go.Heatmap(
        z=z, x=dates, y=pivot.index.tolist(),
        text=text, texttemplate="%{text}",
        colorscale=[[0,'#d73027'],[0.25,'#f46d43'],[0.5,'#ffffbf'],
                    [0.75,'#74c476'],[1,'#1a9641']],
        zmin=0, zmax=100, showscale=True
    ))
    fig.update_layout(title=title, height=max(400, len(pivot)*28+100),
                      margin=dict(l=180,r=20,t=60,b=40))
    return fig

def build_industry_heatmap(month_data, rs_col, title):
    if not month_data: return None
    frames = []
    for e in month_data:
        df = e.get('industry_df')
        if df is None or rs_col not in df.columns: continue
        tmp = df[['Industry', rs_col]].copy()
        tmp['date'] = e['display_date']
        frames.append(tmp)
    if not frames: return None
    merged = pd.concat(frames)
    pivot  = merged.pivot_table(index='Industry', columns='date', values=rs_col)
    pivot  = pivot.sort_values(by=pivot.columns[-1], ascending=False)
    dates  = [d.strftime('%m/%d') for d in pivot.columns]
    z      = pivot.values
    text   = [[f"{v:.1f}" if not pd.isna(v) else "" for v in row] for row in z]
    fig = go.Figure(go.Heatmap(
        z=z, x=dates, y=pivot.index.tolist(),
        text=text, texttemplate="%{text}",
        colorscale=[[0,'#d73027'],[0.25,'#f46d43'],[0.5,'#ffffbf'],
                    [0.75,'#74c476'],[1,'#1a9641']],
        zmin=0, zmax=100, showscale=True
    ))
    fig.update_layout(title=title, height=max(500, len(pivot)*22+100),
                      margin=dict(l=220,r=20,t=60,b=40))
    return fig

# ── Comparison table builders ─────────────────────────────────────────────────
def build_latest_sector_table(all_data):
    if len(all_data) < 1: return None
    latest = all_data[-1]
    prev   = all_data[-2] if len(all_data) >= 2 else None
    df     = latest.get('sector_df')
    if df is None: return None
    rs_cols = [c for c in ['RS_CW','RS_EW','Sector_RS_CW','Sector_RS_EW'] if c in df.columns]
    if not rs_cols: return None
    result = df[['Sector'] + rs_cols].copy()
    if prev is not None:
        pdf = prev.get('sector_df')
        if pdf is not None:
            for c in rs_cols:
                if c in pdf.columns:
                    merged = result.merge(
                        pdf[['Sector', c]].rename(columns={c: c+'_prev'}),
                        on='Sector', how='left')
                    result[c+'_diff'] = (merged[c] - merged[c+'_prev']).round(1)
    return result

def build_latest_industry_table(all_data):
    if len(all_data) < 1: return None
    latest = all_data[-1]
    prev   = all_data[-2] if len(all_data) >= 2 else None
    df     = latest.get('industry_df')
    if df is None: return None
    rs_cols = [c for c in ['RS_CW','RS_EW','Industry_RS_CW','Industry_RS_EW'] if c in df.columns]
    if not rs_cols: return None
    result = df[['Industry'] + rs_cols].copy()
    if prev is not None:
        pdf = prev.get('industry_df')
        if pdf is not None:
            for c in rs_cols:
                if c in pdf.columns:
                    merged = result.merge(
                        pdf[['Industry', c]].rename(columns={c: c+'_prev'}),
                        on='Industry', how='left')
                    result[c+'_diff'] = (merged[c] - merged[c+'_prev']).round(1)
    return result

# ── Single-mode momentum tab (CW or EW) ───────────────────────────────────────
def render_momentum_tab(stock_df, display_date, rs_mode='CW', tab_key='mom_cw'):
    if stock_df is None:
        st.error("銘柄データ（Screening_Results）が見つかりません。"); return

    rs_ind  = f'RS_Percentile_{rs_mode}'
    rs_sec  = f'Sector_RS_{rs_mode}'
    rs_ind2 = f'Industry_RS_{rs_mode}'

    st.caption(f"データ基準日: {display_date}　｜　銘柄数: {len(stock_df):,}")

    with st.expander("🔧 フィルター設定", expanded=True):
        st.subheader("テクニカル条件")
        enable_technical = st.checkbox("テクニカル条件を有効にする", value=True, key=f"{tab_key}_tech")
        c1,c2,c3,c4 = st.columns(4)
        atr_min  = c1.number_input("ATR from 50MA 最小(%)", value=2.0, step=0.1, key=f"{tab_key}_atr_min")
        atr_max  = c2.number_input("ATR from 50MA 最大(%)", value=5.0, step=0.1, key=f"{tab_key}_atr_max")
        adr_min  = c3.number_input("ADR 最小(%)",           value=4.0, step=0.1, key=f"{tab_key}_adr_min")
        enable_ma = c4.checkbox("MA条件を有効にする", value=True, key=f"{tab_key}_ma")
        if enable_ma:
            cm1,cm2,cm3 = st.columns(3)
            above_ma21  = cm1.checkbox("価格 > MA21",  value=True, key=f"{tab_key}_ma21")
            above_ma50  = cm2.checkbox("価格 > MA50",  value=True, key=f"{tab_key}_ma50")
            above_ma150 = cm3.checkbox("価格 > MA150", value=True, key=f"{tab_key}_ma150")
            ma_order    = st.checkbox("MA21 > MA50 > MA150 の順序", value=False, key=f"{tab_key}_maord")
        else:
            above_ma21=above_ma50=above_ma150=ma_order=False

        st.subheader("価格条件")
        enable_price = st.checkbox("価格条件を有効にする", value=True, key=f"{tab_key}_price")
        price_min = st.number_input("最低株価($)", value=10.0, step=1.0, key=f"{tab_key}_pmin")

        st.subheader("ファンダメンタル条件")
        enable_fund = st.checkbox("ファンダメンタル条件を有効にする", value=False, key=f"{tab_key}_fund")
        fund_min = st.number_input("Fundamental Score 最小", value=5, step=1, key=f"{tab_key}_fmin")

        st.subheader(f"RS条件（{rs_mode}）")
        enable_rs = st.checkbox("RS条件を有効にする", value=True, key=f"{tab_key}_rs")
        cr1,cr2,cr3 = st.columns(3)
        individual_rs_min = cr1.number_input("Individual RS 最小(%)", value=80, step=1, key=f"{tab_key}_rs_ind")
        sector_rs_min     = cr2.number_input("Sector RS 最小(%)",     value=79, step=1, key=f"{tab_key}_rs_sec")
        industry_rs_min   = cr3.number_input("Industry RS 最小(%)",   value=80, step=1, key=f"{tab_key}_rs_ind2")

    settings = []
    if enable_technical: settings.append(f"ATR:{atr_min}~{atr_max}% | ADR≥{adr_min}%")
    if enable_price:     settings.append(f"Price≥${price_min}")
    if enable_fund:      settings.append(f"Fundamental≥{fund_min}")
    if enable_rs:        settings.append(f"RS_Ind≥{individual_rs_min}% | RS_Sec≥{sector_rs_min}% | RS_Ind2≥{industry_rs_min}%")
    st.info("現在の条件: " + ("  /  ".join(settings) if settings else "なし"))

    filtered = stock_df.copy()
    if enable_technical:
        if 'ATR_from_50MA' in filtered.columns:
            filtered = filtered[(filtered['ATR_from_50MA'] >= atr_min) &
                                (filtered['ATR_from_50MA'] <= atr_max)]
        if 'ADR' in filtered.columns:
            filtered = filtered[filtered['ADR'] >= adr_min]
        if enable_ma:
            if above_ma21  and 'Price' in filtered.columns and 'MA21'  in filtered.columns:
                filtered = filtered[filtered['Price'] > filtered['MA21']]
            if above_ma50  and 'Price' in filtered.columns and 'MA50'  in filtered.columns:
                filtered = filtered[filtered['Price'] > filtered['MA50']]
            if above_ma150 and 'Price' in filtered.columns and 'MA150' in filtered.columns:
                filtered = filtered[filtered['Price'] > filtered['MA150']]
            if ma_order and all(c in filtered.columns for c in ['MA21','MA50','MA150']):
                filtered = filtered[(filtered['MA21']>filtered['MA50']) &
                                    (filtered['MA50']>filtered['MA150'])]
    if enable_price and 'Price' in filtered.columns:
        filtered = filtered[filtered['Price'] >= price_min]
    if enable_fund and 'Fundamental_Score' in filtered.columns:
        filtered = filtered[filtered['Fundamental_Score'] >= fund_min]
    if enable_rs:
        if rs_ind  in filtered.columns: filtered = filtered[filtered[rs_ind]  >= individual_rs_min]
        if rs_sec  in filtered.columns: filtered = filtered[filtered[rs_sec]  >= sector_rs_min]
        if rs_ind2 in filtered.columns: filtered = filtered[filtered[rs_ind2] >= industry_rs_min]

    st.subheader(f"スクリーニング結果: {len(filtered)} 銘柄")
    if filtered.empty:
        st.warning("条件に合う銘柄が見つかりませんでした。"); return

    display_cols = [c for c in ['Symbol','Company Name','Sector','Industry',
                                rs_ind, rs_sec, rs_ind2,
                                'Price','ATR_from_50MA','ADR','Fundamental_Score']
                    if c in filtered.columns]
    sort_col = rs_ind if rs_ind in filtered.columns else display_cols[0]
    chosen_sort = st.selectbox("並び替え列", display_cols,
                               index=display_cols.index(sort_col) if sort_col in display_cols else 0,
                               key=f"{tab_key}_sort")
    result_df = filtered[display_cols].sort_values(chosen_sort, ascending=False)
    st.dataframe(result_df, use_container_width=True)

    mc1,mc2,mc3,mc4 = st.columns(4)
    mc1.metric("該当銘柄数", len(filtered))
    if rs_ind in filtered.columns:
        mc2.metric(f"平均 {rs_ind}", f"{filtered[rs_ind].mean():.1f}%")
    if 'ADR' in filtered.columns:
        mc3.metric("平均 ADR", f"{filtered['ADR'].mean():.1f}%")
    if 'ATR_from_50MA' in filtered.columns:
        mc4.metric("平均 ATR", f"{filtered['ATR_from_50MA'].mean():.1f}%")

    if 'Sector' in filtered.columns:
        sec_count = filtered['Sector'].value_counts().reset_index()
        sec_count.columns = ['Sector','Count']
        fig = go.Figure(go.Bar(x=sec_count['Sector'], y=sec_count['Count'],
                               marker_color='#1976d2'))
        fig.update_layout(title="セクター分布", xaxis_tickangle=-30, height=350)
        st.plotly_chart(fig, use_container_width=True)

    from datetime import datetime as _dt
    ts = _dt.now().strftime('%Y%m%d_%H%M%S')
    dl1,dl2 = st.columns(2)
    csv_bytes = result_df.to_csv(index=False).encode('utf-8-sig')
    dl1.download_button("📥 CSV ダウンロード", csv_bytes,
                        f"momentum_{rs_mode.lower()}_{ts}.csv","text/csv",
                        key=f"{tab_key}_dl_csv")
    if 'Symbol' in result_df.columns:
        syms = result_df['Symbol'].dropna().tolist()
        dl2.download_button("📋 Symbol リスト (.txt)", "\n".join(syms).encode(),
                            f"symbols_{rs_mode.lower()}_{ts}.txt","text/plain",
                            key=f"{tab_key}_dl_txt")
        st.subheader("TradingView 用 Symbol リスト")
        st.code(",".join(syms), language="text")

# ── CW+EW combined momentum tab ───────────────────────────────────────────────
def render_momentum_tab_both(stock_df, display_date, tab_key='mom_both'):
    if stock_df is None:
        st.error("銘柄データ（Screening_Results）が見つかりません。"); return

    st.caption(f"データ基準日: {display_date}　｜　銘柄数: {len(stock_df):,}")
    st.info("ℹ️ CW（時価総額加重）と EW（均等加重）両方の RS 条件でスクリーニングします。")

    with st.expander("🔧 フィルター設定", expanded=True):

        # ── Technical ──────────────────────────────────────────────────────────
        st.subheader("テクニカル条件")
        enable_technical = st.checkbox("テクニカル条件を有効にする", value=True, key=f"{tab_key}_tech")

        c1, c2, c3, c4 = st.columns(4)
        atr_min = c1.number_input("ATR from 50MA 最小(%)", value=1.5, step=0.1, key=f"{tab_key}_atr_min")
        atr_max = c2.number_input("ATR from 50MA 最大(%)", value=6.0, step=0.1, key=f"{tab_key}_atr_max")
        adr_min = c3.number_input("ADR 最小(%)",           value=2.5, step=0.1, key=f"{tab_key}_adr_min")
        adr_max = c4.number_input("ADR 最大(%)",           value=5.5, step=0.1, key=f"{tab_key}_adr_max")

        enable_ma = st.checkbox("MA条件を有効にする", value=True, key=f"{tab_key}_ma")
        if enable_ma:
            cm1, cm2, cm3 = st.columns(3)
            above_ma21  = cm1.checkbox("価格 > MA21",  value=True,  key=f"{tab_key}_ma21")
            above_ma50  = cm2.checkbox("価格 > MA50",  value=True,  key=f"{tab_key}_ma50")
            above_ma150 = cm3.checkbox("価格 > MA150", value=True,  key=f"{tab_key}_ma150")
            ma_order    = st.checkbox("MA21 > MA50 > MA150 の順序", value=False, key=f"{tab_key}_maord")
        else:
            above_ma21 = above_ma50 = above_ma150 = ma_order = False

        # ── Price ──────────────────────────────────────────────────────────────
        st.subheader("価格条件")
        enable_price = st.checkbox("価格条件を有効にする", value=True, key=f"{tab_key}_price")
        price_min = st.number_input("最低株価($)", value=10.0, step=1.0, key=f"{tab_key}_pmin")

        # ── Fundamental ────────────────────────────────────────────────────────
        st.subheader("ファンダメンタル条件")
        enable_fund = st.checkbox("ファンダメンタル条件を有効にする", value=False, key=f"{tab_key}_fund")
        fund_min = st.number_input("Fundamental Score 最小", value=5, step=1, key=f"{tab_key}_fmin")

        # ── RS (CW & EW) ───────────────────────────────────────────────────────
        st.subheader("RS条件（CW & EW）")
        enable_rs = st.checkbox("RS条件を有効にする", value=True, key=f"{tab_key}_rs")
        col_cw, col_ew = st.columns(2)

        with col_cw:
            st.markdown("**CW（時価総額加重）**")
            individual_rs_min  = st.number_input("Individual RS 最小(%)",  value=80, step=1, key=f"{tab_key}_rs_ind")
            sector_rs_cw_min   = st.number_input("Sector RS CW 最小(%)",   value=70, step=1, key=f"{tab_key}_rs_sec_cw")
            industry_rs_cw_min = st.number_input("Industry RS CW 最小(%)", value=70, step=1, key=f"{tab_key}_rs_ind_cw")

        with col_ew:
            st.markdown("**EW（均等加重）**")
            sector_rs_ew_min   = st.number_input("Sector RS EW 最小(%)",   value=70, step=1, key=f"{tab_key}_rs_sec_ew")
            industry_rs_ew_min = st.number_input("Industry RS EW 最小(%)", value=70, step=1, key=f"{tab_key}_rs_ind_ew")

        # ── Buy Pressure ───────────────────────────────────────────────────────
        st.subheader("バイプレッシャー条件")
        enable_bp = st.checkbox("バイプレッシャー条件を有効にする", value=True, key=f"{tab_key}_bp_enable")

        # (column, label, default_min, has_max, default_max)
        bp_items = [
            ("BP_Stock",       "BP_Stock",       0.60, True,  0.70),
            ("BP_Sector_CW",   "BP_Sector_CW",   0.50, False, None),
            ("BP_Sector_EW",   "BP_Sector_EW",   0.50, False, None),
            ("BP_Industry_CW", "BP_Industry_CW", 0.55, False, None),
            ("BP_Industry_EW", "BP_Industry_EW", 0.55, False, None),
        ]
        bp_settings = {}

        for col, label, def_min, has_max, def_max in bp_items:
            bp_col1, bp_col2, bp_col3 = st.columns([1, 1, 1])
            enabled = bp_col1.checkbox(
                f"{label} を有効", value=True,
                key=f"{tab_key}_bp_{col}_en",
                disabled=not enable_bp
            )
            min_val = bp_col2.number_input(
                f"{label} 最小値", value=def_min, step=0.05, format="%.2f",
                key=f"{tab_key}_bp_{col}_min",
                disabled=not enable_bp
            )
            if has_max:
                max_val = bp_col3.number_input(
                    f"{label} 最大値", value=def_max, step=0.05, format="%.2f",
                    key=f"{tab_key}_bp_{col}_max",
                    disabled=not enable_bp
                )
            else:
                max_val = None

            bp_settings[col] = (enabled and enable_bp, min_val, max_val)

    # ── Settings summary ──────────────────────────────────────────────────────
    settings = []
    if enable_technical:
        settings.append(f"ATR:{atr_min}~{atr_max}% | ADR:{adr_min}~{adr_max}%")
    if enable_price:
        settings.append(f"Price≥${price_min}")
    if enable_fund:
        settings.append(f"Fundamental≥{fund_min}")
    if enable_rs:
        settings.append(
            f"RS_Ind≥{individual_rs_min}% | "
            f"Sec_CW≥{sector_rs_cw_min}% | Ind_CW≥{industry_rs_cw_min}% | "
            f"Sec_EW≥{sector_rs_ew_min}% | Ind_EW≥{industry_rs_ew_min}%"
        )
    if enable_bp:
        bp_parts = []
        for col, (en, mn, mx) in bp_settings.items():
            if en:
                part = f"{col}≥{mn:.2f}"
                if mx is not None:
                    part += f"~{mx:.2f}"
                bp_parts.append(part)
        if bp_parts:
            settings.append("BP: " + " | ".join(bp_parts))
    st.info("現在の条件: " + ("  /  ".join(settings) if settings else "なし"))

    # ── Apply filters ─────────────────────────────────────────────────────────
    filtered = stock_df.copy()

    if enable_technical:
        if 'ATR_from_50MA' in filtered.columns:
            filtered = filtered[(filtered['ATR_from_50MA'] >= atr_min) &
                                (filtered['ATR_from_50MA'] <= atr_max)]
        if 'ADR' in filtered.columns:
            filtered = filtered[(filtered['ADR'] >= adr_min) &
                                (filtered['ADR'] <= adr_max)]
        if enable_ma:
            if above_ma21  and 'Price' in filtered.columns and 'MA21'  in filtered.columns:
                filtered = filtered[filtered['Price'] > filtered['MA21']]
            if above_ma50  and 'Price' in filtered.columns and 'MA50'  in filtered.columns:
                filtered = filtered[filtered['Price'] > filtered['MA50']]
            if above_ma150 and 'Price' in filtered.columns and 'MA150' in filtered.columns:
                filtered = filtered[filtered['Price'] > filtered['MA150']]
            if ma_order and all(c in filtered.columns for c in ['MA21','MA50','MA150']):
                filtered = filtered[(filtered['MA21'] > filtered['MA50']) &
                                    (filtered['MA50'] > filtered['MA150'])]

    if enable_price and 'Price' in filtered.columns:
        filtered = filtered[filtered['Price'] >= price_min]

    if enable_fund and 'Fundamental_Score' in filtered.columns:
        filtered = filtered[filtered['Fundamental_Score'] >= fund_min]

    if enable_rs:
        rs_thresholds = {
            'RS_Percentile_CW': individual_rs_min,
            'Sector_RS_CW':     sector_rs_cw_min,
            'Industry_RS_CW':   industry_rs_cw_min,
            'Sector_RS_EW':     sector_rs_ew_min,
            'Industry_RS_EW':   industry_rs_ew_min,
        }
        for col, thr in rs_thresholds.items():
            if col in filtered.columns:
                filtered = filtered[filtered[col] >= thr]

    for col, (en, mn, mx) in bp_settings.items():
        if en and col in filtered.columns:
            filtered = filtered[filtered[col] >= mn]
            if mx is not None:
                filtered = filtered[filtered[col] <= mx]

    # ── Results ───────────────────────────────────────────────────────────────
    st.subheader(f"スクリーニング結果: {len(filtered)} 銘柄")
    if filtered.empty:
        st.warning("条件に合う銘柄が見つかりませんでした。"); return

    display_cols = [c for c in [
        'Symbol','Company Name','Sector','Industry',
        'RS_Percentile_CW','Sector_RS_CW','Industry_RS_CW',
        'RS_Percentile_EW','Sector_RS_EW','Industry_RS_EW',
        'Price','ATR_from_50MA','ADR','Fundamental_Score',
        'BP_Stock','BP_Sector_CW','BP_Sector_EW','BP_Industry_CW','BP_Industry_EW'
    ] if c in filtered.columns]

    sort_col = 'RS_Percentile_CW' if 'RS_Percentile_CW' in display_cols else display_cols[0]
    chosen_sort = st.selectbox("並び替え列", display_cols,
                               index=display_cols.index(sort_col),
                               key=f"{tab_key}_sort")
    result_df = filtered[display_cols].sort_values(chosen_sort, ascending=False)
    st.dataframe(result_df, use_container_width=True)

    mc1, mc2, mc3, mc4 = st.columns(4)
    mc1.metric("該当銘柄数", len(filtered))
    for col, lbl in [('RS_Percentile_CW','平均 RS_CW'),('RS_Percentile_EW','平均 RS_EW')]:
        if col in filtered.columns:
            mc2.metric(lbl, f"{filtered[col].mean():.1f}%"); break
    if 'ADR' in filtered.columns:
        mc3.metric("平均 ADR", f"{filtered['ADR'].mean():.1f}%")
    if 'ATR_from_50MA' in filtered.columns:
        mc4.metric("平均 ATR", f"{filtered['ATR_from_50MA'].mean():.1f}%")

    if 'Sector' in filtered.columns:
        sec_count = filtered['Sector'].value_counts().reset_index()
        sec_count.columns = ['Sector','Count']
        fig = go.Figure(go.Bar(x=sec_count['Sector'], y=sec_count['Count'],
                               marker_color='#7b1fa2'))
        fig.update_layout(title="セクター分布", xaxis_tickangle=-30, height=350)
        st.plotly_chart(fig, use_container_width=True)

    from datetime import datetime as _dt
    ts = _dt.now().strftime('%Y%m%d_%H%M%S')
    dl1, dl2 = st.columns(2)
    csv_bytes = result_df.to_csv(index=False).encode('utf-8-sig')
    dl1.download_button("📥 CSV ダウンロード", csv_bytes,
                        f"momentum_both_{ts}.csv","text/csv",
                        key=f"{tab_key}_dl_csv")
    if 'Symbol' in result_df.columns:
        syms = result_df['Symbol'].dropna().tolist()
        dl2.download_button("📋 Symbol リスト (.txt)", "\n".join(syms).encode(),
                            f"symbols_both_{ts}.txt","text/plain",
                            key=f"{tab_key}_dl_txt")
        st.subheader("TradingView 用 Symbol リスト")
        st.code(",".join(syms), language="text")

# ── Main UI ───────────────────────────────────────────────────────────────────
with st.spinner("データ読み込み中..."):
    all_data = load_all_data(DATA_FOLDER)

if not all_data:
    st.error("読み込めるデータがありません。"); st.stop()

all_data.sort(key=lambda x: x['date'])
latest           = all_data[-1]
latest_stock_df  = latest.get('stock_df')
latest_disp_date = latest['display_date'].strftime('%Y年%m月%d日')

# Market summary banner
ms = latest.get('market_summary')
if ms is not None:
    try:
        st.markdown("### 📊 マーケットサマリー")
        cols = st.columns(len(ms.columns))
        for i, col in enumerate(ms.columns):
            val = ms.iloc[0][col]
            cols[i].metric(col, val)
    except Exception:
        pass

# Month selector
available_months = get_available_months(all_data)
selected_month   = st.selectbox("📅 表示月を選択", available_months, index=0)
month_data       = filter_by_month(all_data, selected_month)

# ── Tabs ──────────────────────────────────────────────────────────────────────
tabs = st.tabs([
    "📊 セクター RS (CW)",
    "📊 セクター RS (EW)",
    "🏭 業種 RS (CW)",
    "🏭 業種 RS (EW)",
    "📋 セクター比較",
    "📋 業種比較",
    "🚀 モメンタム銘柄 (CW)",
    "🌊 モメンタム銘柄 (EW)",
    "🎯 モメンタム銘柄 CW＋EW",
])
(tab_sec_cw, tab_sec_ew, tab_ind_cw, tab_ind_ew,
 tab_sec_cmp, tab_ind_cmp,
 tab_mom_cw, tab_mom_ew, tab_mom_both) = tabs

with tab_sec_cw:
    st.header("セクター RS ヒートマップ（CW）")
    fig = build_sector_heatmap(month_data, 'RS_CW', f"セクター RS CW — {selected_month}")
    if fig: st.plotly_chart(fig, use_container_width=True)
    else:   st.warning("データがありません。")

with tab_sec_ew:
    st.header("セクター RS ヒートマップ（EW）")
    fig = build_sector_heatmap(month_data, 'RS_EW', f"セクター RS EW — {selected_month}")
    if fig: st.plotly_chart(fig, use_container_width=True)
    else:   st.warning("データがありません。")

with tab_ind_cw:
    st.header("業種 RS ヒートマップ（CW）")
    fig = build_industry_heatmap(month_data, 'RS_CW', f"業種 RS CW — {selected_month}")
    if fig: st.plotly_chart(fig, use_container_width=True)
    else:   st.warning("データがありません。")

with tab_ind_ew:
    st.header("業種 RS ヒートマップ（EW）")
    fig = build_industry_heatmap(month_data, 'RS_EW', f"業種 RS EW — {selected_month}")
    if fig: st.plotly_chart(fig, use_container_width=True)
    else:   st.warning("データがありません。")

with tab_sec_cmp:
    st.header("セクター RS 比較（最新 vs 前日）")
    tbl = build_latest_sector_table(all_data)
    if tbl is not None:
        rs_cols   = [c for c in tbl.columns if 'RS' in c and 'diff' not in c]
        diff_cols = [c for c in tbl.columns if 'diff' in c]
        styled = tbl.sort_values(rs_cols[0], ascending=False).style
        for c in rs_cols:   styled = styled.apply(color_rs_col,   subset=[c])
        for c in diff_cols: styled = styled.apply(color_diff_col, subset=[c])
        st.dataframe(styled, use_container_width=True)
    else:
        st.warning("データがありません。")

with tab_ind_cmp:
    st.header("業種 RS 比較（最新 vs 前日）")
    tbl = build_latest_industry_table(all_data)
    if tbl is not None:
        rs_cols   = [c for c in tbl.columns if 'RS' in c and 'diff' not in c]
        diff_cols = [c for c in tbl.columns if 'diff' in c]
        styled = tbl.sort_values(rs_cols[0], ascending=False).style
        for c in rs_cols:   styled = styled.apply(color_rs_col,   subset=[c])
        for c in diff_cols: styled = styled.apply(color_diff_col, subset=[c])
        st.dataframe(styled, use_container_width=True)
    else:
        st.warning("データがありません。")

with tab_mom_cw:
    st.header("🚀 モメンタム銘柄スクリーニング（CWベース）")
    render_momentum_tab(latest_stock_df, latest_disp_date, rs_mode='CW', tab_key='mom_cw')

with tab_mom_ew:
    st.header("🌊 モメンタム銘柄スクリーニング（EWベース）")
    render_momentum_tab(latest_stock_df, latest_disp_date, rs_mode='EW', tab_key='mom_ew')

with tab_mom_both:
    st.header("🎯 モメンタム銘柄スクリーニング（CW＋EW 両条件）")
    render_momentum_tab_both(latest_stock_df, latest_disp_date, tab_key='mom_both')

gc.collect()
