# =============================================================================
# 人事分析ダッシュボード
# HR Analytics Dashboard
#
# 実行方法:
#   pip install streamlit pandas plotly openpyxl
#   streamlit run hr_dashboard.py
#
# データ形式（Excel）:
#   行1: 帳票名 / 基準日（スキップ）
#   行2: 空行またはタイトル行（スキップ）
#   行3: ヘッダー（社員ID, 生年月日, 性別, 入社年月日, 退職年月日, 職種分類, 採用区分, 雇用区分）
#   行4以降: データ
# =============================================================================

import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from datetime import date, datetime
import warnings

warnings.filterwarnings("ignore")

# =============================================================================
# ページ設定
# =============================================================================
st.set_page_config(
    page_title="人事分析ダッシュボード",
    page_icon="👥",
    layout="wide",
    initial_sidebar_state="expanded",
)

# =============================================================================
# 定数定義
# =============================================================================
RETIREMENT_AGE = 65  # 定年年齢（歳）

# 年代区分の表示順序
AGE_GROUPS = ["20代未満", "20代", "30代", "40代", "50代", "60代以上"]


# =============================================================================
# ユーティリティ関数
# =============================================================================


def get_fiscal_year(d):
    """
    日付から年度を算出する
    ルール: 4月1日〜翌年3月31日 を同じ年度とする

    例)
      2024/4/1  → 2024年度 （4月以降なので当年）
      2024/10/1 → 2024年度 （4月以降なので当年）
      2025/3/1  → 2024年度 （3月は前の年度）
    """
    if d is None or pd.isnull(d):
        return None
    # Timestamp → date へ変換
    if isinstance(d, pd.Timestamp):
        d = d.date()
    elif isinstance(d, datetime):
        d = d.date()
    # 4月以降なら当年、3月以前なら前年
    return d.year if d.month >= 4 else d.year - 1


def calc_age(birth_date, ref_date):
    """
    誕生日と基準日から満年齢を計算する

    計算式:
      基準日の年 − 誕生年 で一旦算出し、
      基準日がその年の誕生日より前なら 1 を引く（誕生日未到達）
    """
    if birth_date is None or ref_date is None:
        return None
    if pd.isnull(birth_date) or pd.isnull(ref_date):
        return None
    if isinstance(birth_date, pd.Timestamp):
        birth_date = birth_date.date()
    if isinstance(ref_date, pd.Timestamp):
        ref_date = ref_date.date()
    try:
        age = ref_date.year - birth_date.year
        # 誕生日をまだ迎えていない場合は 1 を引く
        if (ref_date.month, ref_date.day) < (birth_date.month, birth_date.day):
            age -= 1
        return age
    except Exception:
        return None


def calc_tenure(hire_date, end_date):
    """
    入社日と終了日から勤続年数（小数）を計算する

    計算式:
      (終了日 − 入社日) の日数 ÷ 365.25
      ※ 365.25 は閏年を考慮した平均日数
    """
    if hire_date is None or end_date is None:
        return None
    if pd.isnull(hire_date) or pd.isnull(end_date):
        return None
    if isinstance(hire_date, pd.Timestamp):
        hire_date = hire_date.date()
    if isinstance(end_date, pd.Timestamp):
        end_date = end_date.date()
    try:
        return (end_date - hire_date).days / 365.25
    except Exception:
        return None


def to_age_group(age):
    """年齢（数値）から年代区分文字列を返す"""
    if age is None or pd.isnull(age):
        return "不明"
    a = int(age)
    if a < 20:
        return "20代未満"
    elif a < 30:
        return "20代"
    elif a < 40:
        return "30代"
    elif a < 50:
        return "40代"
    elif a < 60:
        return "50代"
    else:
        return "60代以上"


def retirement_fiscal_year(birth_date):
    """
    誕生日から定年退職年度を算出する

    計算式:
      誕生日 + RETIREMENT_AGE(65) 年 の日付 → その年度
    """
    if birth_date is None or pd.isnull(birth_date):
        return None
    if isinstance(birth_date, pd.Timestamp):
        birth_date = birth_date.date()
    try:
        retire_date = date(birth_date.year + RETIREMENT_AGE, birth_date.month, birth_date.day)
        return get_fiscal_year(retire_date)
    except ValueError:
        # 2/29 生まれの場合は 3/1 に繰り上げ
        retire_date = date(birth_date.year + RETIREMENT_AGE, 3, 1)
        return get_fiscal_year(retire_date)


# =============================================================================
# データ読み込み・前処理
# =============================================================================


def load_data(file):
    """
    アップロードされた Excel ファイルを読み込み、基本的な型変換・年度計算を行う

    対応ファイル形式:
      行1: 帳票名・基準日（スキップ）
      行2: 空白行 or タイトル（スキップ）
      行3: ヘッダー行（社員ID, 生年月日, …）
      行4以降: 実データ

    ※ 上記形式以外の場合は自動的に標準ヘッダーとして読み込みを試みる
    """
    try:
        # まず 3 行目をヘッダーとして試みる（header=2 は 0-indexed で 3 行目）
        df = pd.read_excel(file, header=2, dtype={"社員ID": str})

        # カラム名が期待通りでなければ header=0 で再試行
        expected = {"社員ID", "生年月日", "入社年月日"}
        if not expected.issubset(set(df.columns.str.strip())):
            file.seek(0) if hasattr(file, "seek") else None
            df = pd.read_excel(file, header=0, dtype={"社員ID": str})

        # カラム名の空白を除去
        df.columns = df.columns.str.strip()

        # ヘッダー行の残滓（"社員ID" という文字列がデータ行に混入している場合）を除去
        df = df[df["社員ID"] != "社員ID"].copy()
        df = df.dropna(subset=["社員ID", "入社年月日"]).copy()

    except Exception as e:
        st.error(f"ファイル読み込みエラー: {e}")
        return None

    # 日付型に変換
    for col in ["生年月日", "入社年月日", "退職年月日"]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")

    # 年度を追加
    df["入社年度"] = df["入社年月日"].apply(get_fiscal_year)
    df["退職年度"] = df["退職年月日"].apply(get_fiscal_year)

    return df


def enrich(df: pd.DataFrame, ref_date: date):
    """
    基準日をもとにデータを補完する

    追加フィールド:
      在籍フラグ  : 基準日時点で在籍しているか（bool）
        → 退職年月日が空 OR 退職年月日 > 基準日
      有効退職日  : 基準日以前の退職日（在籍中は NaT）
      年齢        : 在籍者 → 基準日時点 / 退職者 → 退職日時点
      年代        : 年齢から 20 代・30 代… に区分
      勤続年数    : 在籍者 → 基準日まで / 退職者 → 退職日まで（単位:年）
      退職区分    : 自然退職（退職時65歳以上）/ 通常退職（65歳未満）
      定年退職年度: 各社員の定年（65歳）を迎える年度
    """
    df = df.copy()
    if isinstance(ref_date, datetime):
        ref_date = ref_date.date()
    ref_ts = pd.Timestamp(ref_date)

    # ── 在籍フラグ ──────────────────────────────────────────────
    # 在籍中の条件（3つ全て満たすこと）:
    #   ① 入社年月日が存在する
    #   ② 入社年月日 ≤ 基準日（まだ入社していない人は含めない）
    #   ③ 退職年月日が空 または 退職年月日 > 基準日
    df["在籍フラグ"] = (
        df["入社年月日"].notna()
        & (df["入社年月日"] <= ref_ts)
        & (df["退職年月日"].isna() | (df["退職年月日"] > ref_ts))
    )

    # ── 有効退職日（基準日以前の退職日のみ） ─────────────────────
    df["有効退職日"] = df["退職年月日"].where(
        df["退職年月日"].notna() & (df["退職年月日"] <= ref_ts)
    )

    # ── 年齢 ──────────────────────────────────────────────────
    # 在籍者: 基準日時点の年齢
    # 退職者: 退職日時点の年齢
    def row_age(row):
        if row["在籍フラグ"]:
            return calc_age(row["生年月日"], ref_date)
        if pd.notna(row["有効退職日"]):
            return calc_age(row["生年月日"], row["有効退職日"])
        return None

    df["年齢"] = df.apply(row_age, axis=1)
    df["年代"] = df["年齢"].apply(to_age_group)

    # ── 勤続年数 ──────────────────────────────────────────────
    # 計算式: (基準日 or 退職日 − 入社日) ÷ 365.25
    def row_tenure(row):
        if row["在籍フラグ"]:
            return calc_tenure(row["入社年月日"], ref_date)
        if pd.notna(row["有効退職日"]):
            return calc_tenure(row["入社年月日"], row["有効退職日"])
        return None

    df["勤続年数"] = df.apply(row_tenure, axis=1)

    # ── 退職区分 ───────────────────────────────────────────────
    # 自然退職: 退職時の年齢が RETIREMENT_AGE（65歳）以上
    # 通常退職: 退職時の年齢が 65 歳未満
    def row_exit_type(row):
        if row["在籍フラグ"] or pd.isna(row["有効退職日"]):
            return None  # 在籍中は退職区分なし
        a = row["年齢"]
        if a is None:
            return "通常退職"
        return "自然退職" if a >= RETIREMENT_AGE else "通常退職"

    df["退職区分"] = df.apply(row_exit_type, axis=1)

    # ── 定年退職年度 ─────────────────────────────────────────────
    # 各社員が 65 歳になる年度を事前計算（将来予測で使用）
    df["定年退職年度"] = df["生年月日"].apply(retirement_fiscal_year)

    return df


# =============================================================================
# フィルタリング
# =============================================================================


def apply_filters(df: pd.DataFrame, filters: dict):
    """
    サイドバーのフィルター条件を適用してデータを絞り込む

    退職区分フィルターは在籍者には適用しない
    （在籍者は退職区分が None のため除外されてしまうのを防ぐ）
    """
    d = df.copy()

    # ① 入社年度（複数選択）
    if filters.get("入社年度"):
        d = d[d["入社年度"].isin(filters["入社年度"])]

    # ② 性別
    if filters.get("性別") and filters["性別"] != "全て":
        d = d[d["性別"] == filters["性別"]]

    # ③ 採用区分
    if filters.get("採用区分") and filters["採用区分"] != "全て":
        d = d[d["採用区分"] == filters["採用区分"]]

    # ④ 職種区分
    if filters.get("職種分類") and filters["職種分類"] != "全て":
        d = d[d["職種分類"] == filters["職種分類"]]

    # ⑤ 雇用区分
    if filters.get("雇用区分") and filters["雇用区分"] != "全て":
        d = d[d["雇用区分"] == filters["雇用区分"]]

    # ⑥ 退職区分（在籍者は常に含む）
    if filters.get("退職区分") and filters["退職区分"] != "全て":
        d = d[d["在籍フラグ"] | (d["退職区分"] == filters["退職区分"])]

    return d


# =============================================================================
# KPI 計算
# =============================================================================


def compute_kpis(df: pd.DataFrame, ref_date: date):
    """
    ダッシュボード上部に表示する KPI を計算する

    ■ 在籍者数
      基準日時点で在籍フラグ = True の件数

    ■ 離職率（直近完了年度）
      基準日の前年度（例: 基準日 2026/3/31 → 2025年度）を対象に計算
      総離職率   = 年度退職者数 ÷ 期首在籍者数 × 100
      通常離職率 = 通常退職者数 ÷ 期首在籍者数 × 100  （退職時 65 歳未満）
      自然退職率 = 自然退職者数 ÷ 期首在籍者数 × 100  （退職時 65 歳以上）

      ※ 期首在籍者数 = その年度の 4/1 時点で在籍していた人数
        （入社日 <= 4/1 かつ 退職日なし or 退職日 > 4/1）

    ■ 平均勤続年数（在籍者のみ）
      基準日 − 入社日（年）の平均

    ■ 早期離職率（自然退職を除く）
      X 年以内通常退職者数 ÷ フィルター後の全採用人数 × 100
    """
    if isinstance(ref_date, datetime):
        ref_date = ref_date.date()

    active_n = int(df["在籍フラグ"].sum())

    # 直近完了年度 = 基準日の年度の前年度
    ref_fy = get_fiscal_year(ref_date)
    prev_fy = ref_fy - 1

    # 期首在籍者数（前年度 4/1 時点）
    fy_start_ts = pd.Timestamp(date(prev_fy, 4, 1))
    bos = df[
        df["入社年月日"].notna()
        & (df["入社年月日"] <= fy_start_ts)
        & (df["退職年月日"].isna() | (df["退職年月日"] > fy_start_ts))
    ].shape[0]

    # 前年度退職者
    fy_exits = df[df["退職年度"] == prev_fy]
    n_total = len(fy_exits)
    n_normal = len(fy_exits[fy_exits["退職区分"] == "通常退職"])
    n_natural = len(fy_exits[fy_exits["退職区分"] == "自然退職"])

    tr = round(n_total / bos * 100, 1) if bos > 0 else 0.0
    nr = round(n_normal / bos * 100, 1) if bos > 0 else 0.0
    nat = round(n_natural / bos * 100, 1) if bos > 0 else 0.0

    # 平均勤続年数（在籍者のみ）
    avg_ten = df[df["在籍フラグ"]]["勤続年数"].mean()
    avg_ten = round(float(avg_ten), 1) if pd.notna(avg_ten) else 0.0

    # 早期離職率: X 年以内通常退職者 ÷ 全レコード数
    total_n = len(df)

    def early_rate(yrs: float) -> float:
        if total_n == 0:
            return 0.0
        n = df[
            (df["退職区分"] == "通常退職")
            & df["勤続年数"].notna()
            & (df["勤続年数"] <= yrs)
        ].shape[0]
        return round(n / total_n * 100, 1)

    return {
        "在籍者数": active_n,
        "総離職率": tr,
        "通常離職率": nr,
        "自然退職率": nat,
        "平均勤続年数": avg_ten,
        "早期離職_1年": early_rate(1),
        "早期離職_3年": early_rate(3),
        "早期離職_5年": early_rate(5),
        "参照年度": prev_fy,
        "期首在籍": bos,
    }


# =============================================================================
# 年度別サマリーテーブル（在籍数・退職数・入社数・離職率を年度ごとに集計）
# =============================================================================


def build_yearly_summary(df: pd.DataFrame):
    """
    年度ごとの在籍数・入社数・退職数・離職率をまとめたテーブルを作成する

    各列の定義:
      在籍者数: 年度末（翌年 3/31）時点で在籍している人数
      入社人数: その年度に入社した人数
      退職人数: その年度に退職した人数
      うち通常退職/自然退職
      期首在籍: その年度 4/1 時点の在籍者数（離職率の分母）
      総離職率: 退職人数 ÷ 期首在籍 × 100
    """
    hire_fy = sorted([int(y) for y in df["入社年度"].dropna().unique()])
    exit_fy = sorted([int(y) for y in df["退職年度"].dropna().unique()])
    all_fy = sorted(set(hire_fy + exit_fy))

    rows = []
    for fy in all_fy:
        fy_end = pd.Timestamp(date(fy + 1, 3, 31))
        fy_start = pd.Timestamp(date(fy, 4, 1))

        # 年度末在籍者数
        active_eoy = df[
            df["入社年月日"].notna()
            & (df["入社年月日"] <= fy_end)
            & (df["退職年月日"].isna() | (df["退職年月日"] > fy_end))
        ].shape[0]

        # 期首在籍者数
        bos = df[
            df["入社年月日"].notna()
            & (df["入社年月日"] <= fy_start)
            & (df["退職年月日"].isna() | (df["退職年月日"] > fy_start))
        ].shape[0]

        # 入社人数
        hires = df[df["入社年度"] == fy].shape[0]

        # 退職人数
        exits = df[df["退職年度"] == fy]
        n_exit = len(exits)
        n_normal = len(exits[exits["退職区分"] == "通常退職"])
        n_natural = len(exits[exits["退職区分"] == "自然退職"])

        # 離職率
        total_rate = round(n_exit / bos * 100, 1) if bos > 0 else 0.0
        normal_rate = round(n_normal / bos * 100, 1) if bos > 0 else 0.0
        natural_rate = round(n_natural / bos * 100, 1) if bos > 0 else 0.0

        rows.append(
            {
                "年度": f"{fy}年度",
                "期首在籍": bos,
                "入社人数": hires,
                "退職人数": n_exit,
                "　通常退職": n_normal,
                "　自然退職": n_natural,
                "年度末在籍": active_eoy,
                "総離職率(%)": total_rate,
                "通常離職率(%)": normal_rate,
                "自然退職率(%)": natural_rate,
            }
        )
    return pd.DataFrame(rows)


# =============================================================================
# グラフ関数
# =============================================================================


def fig_active_trend(df: pd.DataFrame):
    """
    ① 在籍者数推移（折れ線グラフ）

    各年度末（翌年 3/31）時点の在籍者数を折れ線で表示
    計算式: 入社日 <= 年度末 かつ (退職日なし or 退職日 > 年度末) の件数
    """
    fy_set = set()
    fy_set.update([int(y) for y in df["入社年度"].dropna().unique()])
    fy_set.update([int(y) for y in df["退職年度"].dropna().unique()])
    if not fy_set:
        return go.Figure().update_layout(title="① 在籍者数推移")

    all_fy = sorted(fy_set)
    counts = []
    for fy in all_fy:
        fy_end = pd.Timestamp(date(fy + 1, 3, 31))
        n = df[
            df["入社年月日"].notna()
            & (df["入社年月日"] <= fy_end)
            & (df["退職年月日"].isna() | (df["退職年月日"] > fy_end))
        ].shape[0]
        counts.append(n)

    fig = go.Figure(
        go.Scatter(
            x=[f"{y}年度" for y in all_fy],
            y=counts,
            mode="lines+markers",
            line=dict(color="royalblue", width=2),
            marker=dict(size=7),
            hovertemplate="%{x}: %{y}人<extra></extra>",
        )
    )
    fig.update_layout(
        title="① 在籍者数推移（年度末時点）",
        xaxis_title="年度",
        yaxis_title="在籍者数（人）",
        height=380,
    )
    return fig


def fig_population_pyramid(df: pd.DataFrame):
    """
    ② 在籍者年代構成（人口ピラミッド）

    在籍者を年代・性別で分けて表示
    男性: 左側（負の値として描画）、女性: 右側（正の値）
    """
    active = df[df["在籍フラグ"]].copy()
    if active.empty:
        return go.Figure().update_layout(title="② 在籍者年代構成")

    male = (
        active[active["性別"] == "男性"]
        .groupby("年代")
        .size()
        .reindex(AGE_GROUPS, fill_value=0)
    )
    female = (
        active[active["性別"] == "女性"]
        .groupby("年代")
        .size()
        .reindex(AGE_GROUPS, fill_value=0)
    )

    max_val = max(int(male.max()), int(female.max()), 1)
    # 目盛り: 5 刻み or 適切な刻み幅
    step = max(1, (max_val + 5) // 5)
    ticks = list(range(-(max_val + step), max_val + step + 1, step))

    fig = go.Figure()
    fig.add_trace(
        go.Bar(
            y=AGE_GROUPS,
            x=[-int(v) for v in male],
            orientation="h",
            name="男性",
            marker_color="steelblue",
            text=[f"{int(v)}人" for v in male],
            textposition="inside",
        )
    )
    fig.add_trace(
        go.Bar(
            y=AGE_GROUPS,
            x=[int(v) for v in female],
            orientation="h",
            name="女性",
            marker_color="salmon",
            text=[f"{int(v)}人" for v in female],
            textposition="inside",
        )
    )
    fig.update_layout(
        title="② 在籍者年代構成（人口ピラミッド）",
        barmode="overlay",
        xaxis=dict(
            title="人数（人）",
            tickvals=ticks,
            ticktext=[str(abs(t)) for t in ticks],
        ),
        yaxis_title="年代",
        legend=dict(orientation="h", y=-0.18, x=0.5, xanchor="center"),
        margin=dict(b=80),
        height=400,
    )
    return fig


def fig_turnover_trend(df: pd.DataFrame):
    """
    ③ 離職率推移（折れ線グラフ）

    各年度の離職率を 3 系列で表示:
      総離職率   = 年度退職者数 ÷ 期首在籍者数 × 100
      通常離職率 = 通常退職者数 ÷ 期首在籍者数 × 100
      自然退職率 = 自然退職者数 ÷ 期首在籍者数 × 100
    """
    exit_yrs = sorted([int(y) for y in df["退職年度"].dropna().unique()])
    if not exit_yrs:
        return go.Figure().update_layout(title="③ 離職率推移")

    t_rates, n_rates, nat_rates = [], [], []
    for fy in exit_yrs:
        fy_start = pd.Timestamp(date(fy, 4, 1))
        bos = df[
            df["入社年月日"].notna()
            & (df["入社年月日"] <= fy_start)
            & (df["退職年月日"].isna() | (df["退職年月日"] > fy_start))
        ].shape[0]
        if bos == 0:
            t_rates.append(0.0)
            n_rates.append(0.0)
            nat_rates.append(0.0)
            continue
        fx = df[df["退職年度"] == fy]
        t_rates.append(round(len(fx) / bos * 100, 1))
        n_rates.append(round(len(fx[fx["退職区分"] == "通常退職"]) / bos * 100, 1))
        nat_rates.append(round(len(fx[fx["退職区分"] == "自然退職"]) / bos * 100, 1))

    xl = [f"{y}年度" for y in exit_yrs]
    fig = go.Figure()
    fig.add_trace(
        go.Scatter(x=xl, y=t_rates, name="総離職率", mode="lines+markers", line=dict(color="crimson", width=2))
    )
    fig.add_trace(
        go.Scatter(x=xl, y=n_rates, name="通常離職率", mode="lines+markers", line=dict(color="orange", width=2))
    )
    fig.add_trace(
        go.Scatter(x=xl, y=nat_rates, name="自然退職率", mode="lines+markers", line=dict(color="green", width=2))
    )
    fig.update_layout(
        title="③ 離職率推移",
        xaxis_title="年度",
        yaxis_title="離職率（%）",
        hovermode="x unified",
        height=380,
    )
    return fig


def fig_hire_trend(df: pd.DataFrame):
    """
    ④ 採用人数推移（棒グラフ）

    年度ごとの採用人数（新卒・中途の積み上げ）を表示
    """
    if df.empty:
        return go.Figure().update_layout(title="④ 採用人数推移")

    hby = (
        df.groupby(["入社年度", "採用区分"])
        .size()
        .reset_index(name="人数")
        .sort_values("入社年度")
    )
    all_fy = sorted(df["入社年度"].dropna().unique())
    xl = [f"{int(y)}年度" for y in all_fy]

    fig = go.Figure()
    colors = {"新卒": "steelblue", "中途": "lightcoral"}
    for htype in ["新卒", "中途"]:
        sub = hby[hby["採用区分"] == htype].set_index("入社年度")
        y_vals = [int(sub.loc[fy, "人数"]) if fy in sub.index else 0 for fy in all_fy]
        fig.add_trace(
            go.Bar(
                x=xl,
                y=y_vals,
                name=htype,
                marker_color=colors.get(htype, "gray"),
                text=y_vals,
                textposition="inside",
            )
        )
    fig.update_layout(
        title="④ 採用人数推移",
        xaxis_title="入社年度",
        yaxis_title="採用人数（人）",
        barmode="stack",
        height=380,
    )
    return fig


def fig_cohort_retention(df: pd.DataFrame, hire_type: str, ref_date: date):
    """
    ⑤⑥ 定着率コホートグラフ（新卒 / 中途）

    各採用年度のコホートについて、入社後 0〜5 年の定着率を折れ線で表示

    計算式:
      定着率 = 入社 X 年後時点で在籍している人数 ÷ コホートの採用人数 × 100

    在籍判定:
      判定日 = 採用年度の 4/1 + X 年
      入社日 <= 判定日 かつ (退職日なし or 退職日 > 判定日) → 在籍
    """
    if isinstance(ref_date, datetime):
        ref_date = ref_date.date()

    cdf = df[df["採用区分"] == hire_type].copy() if hire_type != "全て" else df.copy()
    if cdf.empty:
        return go.Figure().update_layout(title=f"{'⑤' if hire_type=='新卒' else '⑥'} {hire_type}採用 定着率")

    cohort_years = sorted([int(y) for y in cdf["入社年度"].dropna().unique()])
    fig = go.Figure()

    for cohort_fy in cohort_years:
        cohort = cdf[cdf["入社年度"] == cohort_fy]
        total = len(cohort)
        if total == 0:
            continue

        x_pts, y_pts = [], []
        for yr in range(0, 6):
            check_ts = pd.Timestamp(date(cohort_fy + yr, 4, 1))
            if check_ts.date() > ref_date:
                break
            n_active = cohort[
                cohort["入社年月日"].notna()
                & (cohort["入社年月日"] <= check_ts)
                & (cohort["退職年月日"].isna() | (cohort["退職年月日"] > check_ts))
            ].shape[0]
            x_pts.append(yr)
            y_pts.append(round(n_active / total * 100, 1))

        if len(x_pts) > 1:
            fig.add_trace(
                go.Scatter(
                    x=x_pts,
                    y=y_pts,
                    mode="lines+markers",
                    name=f"{cohort_fy}年度",
                    marker=dict(size=5),
                    hovertemplate=f"{cohort_fy}年度: %{{y:.1f}}%<extra></extra>",
                )
            )

    num = "⑤" if hire_type == "新卒" else "⑥"
    fig.update_layout(
        title=f"{num} {hire_type}採用 定着率（コホート分析）",
        xaxis_title="経過年数（年）",
        yaxis=dict(title="定着率（%）", range=[0, 105]),
        hovermode="x unified",
        height=400,
    )
    return fig


def fig_tenure_hist(df: pd.DataFrame):
    """
    ⑦ 勤続年数分布（ヒストグラム）

    在籍者・退職者の勤続年数（年）の分布をヒストグラムで表示
    """
    data = df["勤続年数"].dropna()
    if data.empty:
        return go.Figure().update_layout(title="⑦ 勤続年数分布")

    fig = go.Figure(
        go.Histogram(
            x=data,
            nbinsx=25,
            marker_color="steelblue",
            opacity=0.8,
            name="勤続年数",
            hovertemplate="勤続 %{x:.0f}年: %{y}人<extra></extra>",
        )
    )
    avg_tenure = data.mean()
    # 平均値の垂直線
    fig.add_vline(
        x=avg_tenure,
        line_color="red",
        line_dash="dash",
        annotation_text=f"平均: {avg_tenure:.1f}年",
        annotation_position="top right",
    )
    fig.update_layout(
        title="⑦ 勤続年数分布",
        xaxis_title="勤続年数（年）",
        yaxis_title="人数（人）",
        bargap=0.05,
        height=380,
    )
    return fig


def fig_age_group_turnover(df: pd.DataFrame):
    """
    ⑧ 年代別離職率（棒グラフ）

    計算式:
      年代別離職率 = 各年代の退職者数 ÷ 各年代の全人数 × 100
      （在籍者 + 退職者の合計を分母とする）
    """
    total_by = df.groupby("年代").size()
    exits_by = df[df["退職区分"].notna()].groupby("年代").size()

    rates, normal_rates, natural_rates = [], [], []
    for ag in AGE_GROUPS:
        tot = int(total_by.get(ag, 0))
        ex_n = int(
            df[(df["年代"] == ag) & (df["退職区分"] == "通常退職")].shape[0]
        )
        ex_nat = int(
            df[(df["年代"] == ag) & (df["退職区分"] == "自然退職")].shape[0]
        )
        rates.append(round((ex_n + ex_nat) / tot * 100, 1) if tot > 0 else 0.0)
        normal_rates.append(round(ex_n / tot * 100, 1) if tot > 0 else 0.0)
        natural_rates.append(round(ex_nat / tot * 100, 1) if tot > 0 else 0.0)

    fig = go.Figure()
    fig.add_trace(
        go.Bar(x=AGE_GROUPS, y=normal_rates, name="通常退職", marker_color="orange", text=[f"{r}%" for r in normal_rates], textposition="inside")
    )
    fig.add_trace(
        go.Bar(x=AGE_GROUPS, y=natural_rates, name="自然退職", marker_color="green", text=[f"{r}%" for r in natural_rates], textposition="inside")
    )
    fig.update_layout(
        title="⑧ 年代別離職率",
        xaxis_title="年代",
        yaxis_title="離職率（%）",
        barmode="stack",
        height=380,
    )
    return fig


def fig_early_attrition(df: pd.DataFrame):
    """
    ⑨ 早期離職率（棒グラフ）

    採用年度別に 1 年・3 年・5 年以内の通常退職率を表示
    計算式: X 年以内通常退職者 ÷ コホート採用人数 × 100
    ※ 自然退職（65 歳以上）は除外
    """
    cohort_years = sorted([int(y) for y in df["入社年度"].dropna().unique()])
    if not cohort_years:
        return go.Figure().update_layout(title="⑨ 早期離職率")

    y1, y3, y5 = [], [], []
    for fy in cohort_years:
        cohort = df[df["入社年度"] == fy]
        total = len(cohort)
        if total == 0:
            y1.append(0.0)
            y3.append(0.0)
            y5.append(0.0)
            continue

        def rate(yrs, c=cohort, t=total):
            n = c[
                (c["退職区分"] == "通常退職")
                & c["勤続年数"].notna()
                & (c["勤続年数"] <= yrs)
            ].shape[0]
            return round(n / t * 100, 1)

        y1.append(rate(1))
        y3.append(rate(3))
        y5.append(rate(5))

    xl = [f"{y}年度" for y in cohort_years]
    fig = go.Figure()
    fig.add_trace(go.Bar(x=xl, y=y1, name="1年以内", marker_color="#aec7e8"))
    fig.add_trace(go.Bar(x=xl, y=y3, name="3年以内", marker_color="#1f77b4"))
    fig.add_trace(go.Bar(x=xl, y=y5, name="5年以内", marker_color="#0d4b8a"))
    fig.update_layout(
        title="⑨ 早期離職率（採用年度別・自然退職除く）",
        xaxis_title="入社年度",
        yaxis_title="離職率（%）",
        barmode="group",
        hovermode="x unified",
        height=380,
    )
    return fig


def fig_job_turnover(df: pd.DataFrame):
    """
    ⑩ 職種別離職率（積み上げ棒グラフ）

    横軸: 職種
    縦軸: 通常退職率 + 自然退職率の積み上げ
    計算式: 各退職区分の退職者数 ÷ 職種別総人数 × 100
    """
    if "職種分類" not in df.columns:
        return go.Figure().update_layout(title="⑩ 職種別離職率")

    # 職種の表示順を固定（未定は除外）
    JOB_ORDER = ["営業", "SE", "社会基盤", "ITAD", "EIT", "コーポレート"]
    _all_jobs = df["職種分類"].dropna().unique().tolist()
    jobs = [j for j in JOB_ORDER if j in _all_jobs]
    n_rates, nat_rates = [], []

    for jt in jobs:
        sub = df[df["職種分類"] == jt]
        total = len(sub)
        if total == 0:
            n_rates.append(0.0)
            nat_rates.append(0.0)
            continue
        n_rates.append(round(len(sub[sub["退職区分"] == "通常退職"]) / total * 100, 1))
        nat_rates.append(round(len(sub[sub["退職区分"] == "自然退職"]) / total * 100, 1))

    fig = go.Figure()
    fig.add_trace(
        go.Bar(x=jobs, y=n_rates, name="通常退職", marker_color="orange")
    )
    fig.add_trace(
        go.Bar(x=jobs, y=nat_rates, name="自然退職", marker_color="green")
    )
    fig.update_layout(
        title="⑩ 職種別離職率（通常・自然退職の積み上げ）",
        xaxis_title="職種",
        yaxis_title="離職率（%）",
        barmode="stack",
        height=380,
    )
    return fig


def fig_cohort_heatmap(df: pd.DataFrame):
    """
    ⑪ 採用年度 × 離職年度 ヒートマップ

    Y 軸: 採用年度（コホート）
    X 軸: 離職年度
    値  : 各コホートの退職者数
    """
    exits = df[df["退職年度"].notna() & df["入社年度"].notna()].copy()
    if exits.empty:
        return go.Figure().update_layout(title="⑪ 採用年度×離職年度ヒートマップ")

    pivot = (
        exits.groupby(["入社年度", "退職年度"])
        .size()
        .reset_index(name="人数")
    )
    pt = pivot.pivot(index="入社年度", columns="退職年度", values="人数").fillna(0)

    fig = go.Figure(
        go.Heatmap(
            z=pt.values,
            colorscale="Blues",
            x=[f"{int(c)}年度" for c in pt.columns],
            y=[f"{int(r)}年度" for r in pt.index],
            text=pt.values.astype(int),
            texttemplate="%{text}",
            hovertemplate="採用: %{y}<br>離職: %{x}<br>人数: %{z:.0f}人<extra></extra>",
        )
    )
    fig.update_layout(
        title="⑪ 採用年度×離職年度ヒートマップ",
        xaxis_title="離職年度",
        yaxis_title="採用年度",
        height=500,
    )
    return fig


def fig_gender_ratios(df: pd.DataFrame):
    """
    男女別在籍割合 / 採用割合（パイチャート）
    """
    active = df[df["在籍フラグ"]]
    active_gender = active.groupby("性別").size()

    hire_gender = df.groupby("性別").size()

    fig = go.Figure()
    # 左: 在籍割合
    fig.add_trace(
        go.Pie(
            labels=active_gender.index.tolist(),
            values=active_gender.values.tolist(),
            name="在籍",
            title="在籍者",
            domain=dict(x=[0, 0.45]),
            hole=0.4,
            marker_colors=["steelblue", "salmon"],
        )
    )
    # 右: 採用割合
    fig.add_trace(
        go.Pie(
            labels=hire_gender.index.tolist(),
            values=hire_gender.values.tolist(),
            name="採用",
            title="採用（全期間）",
            domain=dict(x=[0.55, 1.0]),
            hole=0.4,
            marker_colors=["steelblue", "salmon"],
        )
    )
    fig.update_layout(
        title="男女別在籍割合 / 採用割合",
        height=360,
        legend=dict(orientation="h", y=-0.1),
    )
    return fig


# =============================================================================
# 将来予測の共通計算
# =============================================================================


def _build_forecast_base(df: pd.DataFrame, ref_date: date,
                         recent_years: int = 5,
                         override_hire: float = None,
                         override_nr: float = None):
    """
    退職予測・人員数予測の共通データを計算する

    パラメータ:
      recent_years  : 平均採用数・離職率の計算に使う直近年数（デフォルト5年）
      override_hire : ユーザー指定の年間採用数（Noneなら自動計算）
      override_nr   : ユーザー指定の通常離職率（Noneなら自動計算）

    戻り値:
      active       : 在籍者 DataFrame
      ref_fy       : 基準日の年度
      avg_nr       : 通常離職率（直近N年平均 or ユーザー指定）
      avg_hire     : 年間採用数（直近N年平均 or ユーザー指定）
    """
    if isinstance(ref_date, datetime):
        ref_date = ref_date.date()
    ref_fy = get_fiscal_year(ref_date)

    active = df[df["在籍フラグ"]].copy()

    # ── 通常離職率: 直近 recent_years 年の平均 ──────────────────
    # 計算式: 通常退職者数 ÷ 期首在籍者数 の直近N年平均
    past_fys = sorted([int(y) for y in df["退職年度"].dropna().unique()])
    # 直近N年に絞る
    recent_fys = past_fys[-recent_years:] if len(past_fys) >= recent_years else past_fys
    nr_list = []
    for fy in recent_fys:
        fy_start = pd.Timestamp(date(fy, 4, 1))
        bos = df[
            df["入社年月日"].notna()
            & (df["入社年月日"] <= fy_start)
            & (df["退職年月日"].isna() | (df["退職年月日"] > fy_start))
        ].shape[0]
        if bos > 0:
            n = df[(df["退職年度"] == fy) & (df["退職区分"] == "通常退職")].shape[0]
            nr_list.append(n / bos)

    avg_nr = override_nr if override_nr is not None else (
        float(np.mean(nr_list)) if nr_list else 0.05
    )

    # ── 年間採用数: 直近 recent_years 年の平均 ──────────────────
    # 全期間平均ではなく直近N年を使うことで、現在の採用規模を反映する
    hire_by_fy = df.groupby("入社年度").size().sort_index()
    # 基準日の年度は年度途中のため除外
    hire_by_fy = hire_by_fy[hire_by_fy.index < ref_fy]
    recent_hires = hire_by_fy.tail(recent_years)
    avg_hire = override_hire if override_hire is not None else (
        float(recent_hires.mean()) if len(recent_hires) > 0 else 0.0
    )

    return active, ref_fy, avg_nr, avg_hire


def _count_mandatory_retirements(active_df: pd.DataFrame, fy: int) -> int:
    """
    指定年度に定年（65歳）を迎える在籍者を数える

    計算式:
      retirement_fiscal_year(生年月日) == fy の人数
    """
    return int((active_df["定年退職年度"] == fy).sum())


def fig_retirement_forecast(df: pd.DataFrame, ref_date: date,
                             recent_years: int = 5,
                             override_hire: float = None,
                             override_nr: float = None):
    """
    ⑫ 退職予測（積み上げ棒グラフ）

    基準日から 5 年間の退職予測を表示

    定年退職予測:
      現在の在籍者のうち、将来65歳になる年度ごとに集計

    通常退職予測:
      (期首在籍者数 − 定年退職者数) × 過去平均通常離職率
    """
    active, ref_fy, avg_nr, _ = _build_forecast_base(
        df, ref_date, recent_years, override_hire, override_nr)
    if active.empty:
        return go.Figure().update_layout(title="⑫ 退職予測")

    forecast_fys = list(range(ref_fy, ref_fy + 5))
    mandatory_list, normal_list = [], []
    forecast_active = active.copy()
    current_count = len(forecast_active)

    for fy in forecast_fys:
        mandatory = _count_mandatory_retirements(forecast_active, fy)
        normal = int(round(max(0, current_count - mandatory) * avg_nr))

        mandatory_list.append(mandatory)
        normal_list.append(normal)

        # 定年退職者を除去して次年度へ更新（近似計算）
        forecast_active = forecast_active[forecast_active["定年退職年度"] != fy].copy()
        current_count = max(0, len(forecast_active) - normal)

    xl = [f"{y}年度" for y in forecast_fys]
    fig = go.Figure()
    fig.add_trace(
        go.Bar(x=xl, y=mandatory_list, name="定年退職（予測）", marker_color="#2ca02c",
               text=mandatory_list, textposition="inside")
    )
    fig.add_trace(
        go.Bar(x=xl, y=normal_list, name="通常退職（予測）", marker_color="#ff7f0e",
               text=normal_list, textposition="inside")
    )
    fig.update_layout(
        title=f"⑫ 退職予測（{ref_fy}〜{ref_fy+4}年度）",
        xaxis_title="年度",
        yaxis_title="予測退職者数（人）",
        barmode="stack",
        height=420,
        margin=dict(t=80),
        annotations=[
            dict(
                x=0.5, y=1.12, xref="paper", yref="paper",
                text=f"※ 過去平均通常離職率: {round(avg_nr*100, 1)}%",
                showarrow=False,
                font=dict(size=11, color="gray"),
                xanchor="center",
            )
        ],
    )
    return fig


def fig_headcount_forecast(df: pd.DataFrame, ref_date: date,
                            recent_years: int = 5,
                            override_hire: float = None,
                            override_nr: float = None):
    """
    ⑬ 人員数予測（折れ線グラフ）

    基準日から 5 年間の人員数推移を予測

    シンプルモデル（年度ごとに繰り返し計算）:
      翌年度人員 = 今年度人員 + 平均採用数 − 定年退職者数 − 通常退職予測数

    パラメータ:
      平均採用数  : 過去の年度別採用人数の平均
      通常離職率  : 過去の通常離職率の平均
    """
    active, ref_fy, avg_nr, avg_hire = _build_forecast_base(
        df, ref_date, recent_years, override_hire, override_nr)
    if active.empty:
        return go.Figure().update_layout(title="⑬ 人員数予測")

    forecast_fys = list(range(ref_fy, ref_fy + 6))
    headcounts = [len(active)]
    forecast_active = active.copy()

    for fy in forecast_fys[:-1]:
        current = headcounts[-1]
        mandatory = _count_mandatory_retirements(forecast_active, fy)
        normal = int(round(max(0, current - mandatory) * avg_nr))
        next_count = current - mandatory - normal + int(round(avg_hire))
        headcounts.append(max(0, next_count))

        # 定年退職者を除去
        forecast_active = forecast_active[forecast_active["定年退職年度"] != fy].copy()

    xl = [f"{y}年度" for y in forecast_fys]
    fig = go.Figure()
    fig.add_trace(
        go.Scatter(
            x=xl, y=headcounts,
            mode="lines+markers",
            name="予測人員数",
            line=dict(color="royalblue", width=2, dash="dot"),
            marker=dict(size=8),
            hovertemplate="%{x}: %{y}人<extra></extra>",
        )
    )
    # 現在値を強調（赤い星）
    fig.add_trace(
        go.Scatter(
            x=[xl[0]], y=[headcounts[0]],
            mode="markers", name="現在人員数",
            marker=dict(color="red", size=14, symbol="star"),
        )
    )
    fig.update_layout(
        title=f"⑬ 人員数予測（{ref_fy}〜{ref_fy+5}年度）",
        xaxis_title="年度",
        yaxis_title="人員数（人）",
        height=420,
        margin=dict(t=80),
        annotations=[
            dict(
                x=0.5, y=1.12, xref="paper", yref="paper",
                text=(
                    f"※ 平均採用数: {round(avg_hire)}人/年　"
                    f"通常離職率: {round(avg_nr*100, 1)}%"
                ),
                showarrow=False,
                font=dict(size=11, color="gray"),
                xanchor="center",
            )
        ],
    )
    return fig


# =============================================================================
# メイン関数
# =============================================================================


def main():
    st.title("👥 人事分析ダッシュボード")
    st.caption(
        "社員データ（Excel）をアップロードし、基準日を設定すると各種指標が自動計算されます。"
    )

    # セッション状態の初期化
    if "use_sample" not in st.session_state:
        st.session_state["use_sample"] = False

    # =========================================================================
    # サイドバー（上部: データ取得 / 下部: フィルター）
    # =========================================================================
    with st.sidebar:
        st.header("⚙️ 設定・フィルター")

        # ── データアップロード ───────────────────────────────────────
        st.subheader("📂 データ取得")
        uploaded_file = st.file_uploader(
            "Excelファイルをアップロード",
            type=["xlsx", "xls"],
            help=(
                "必須列: 社員ID / 生年月日 / 性別 / 入社年月日 / 退職年月日 "
                "/ 職種分類 / 採用区分 / 雇用区分\n\n"
                "退職年月日: 在籍中は空欄のままにしてください"
            ),
        )
        if uploaded_file is not None:
            st.session_state["use_sample"] = False

        # ── 基準日 ──────────────────────────────────────────────────
        st.subheader("📅 基準日設定")
        reference_date = st.date_input(
            "基準日（この時点の情報を表示）",
            value=date(2026, 3, 31),
            help=(
                "在籍者数・年齢・勤続年数はこの日付を基準に計算されます。\n"
                "基準日より後の退職日を持つ社員は「在籍中」として扱います。"
            ),
        )

    # =========================================================================
    # データ読み込み
    # =========================================================================
    raw_df = None
    if uploaded_file is not None:
        raw_df = load_data(uploaded_file)
    elif st.session_state.get("use_sample"):
        raw_df = _generate_sample_data()

    if raw_df is None:
        # データ未読み込み時のガイダンス
        st.info("👆 左のパネルから Excel ファイルをアップロードしてください。")
        st.markdown(
            """
            ### 必要なデータ形式（Excel）

            | 社員ID | 生年月日 | 性別 | 入社年月日 | 退職年月日 | 職種分類 | 採用区分 | 雇用区分 |
            |--------|---------|------|-----------|-----------|---------|---------|---------|
            | E001 | 1985/4/1 | 男性 | 2010/4/1 | （在籍は空欄）| SE | 新卒 | 正社員 |
            | E002 | 1990/7/15 | 女性 | 2015/4/1 | 2022/3/31 | 営業 | 新卒 | 正社員 |

            **ポイント**:
            - 退職年月日: **在籍中は空欄**（値なし）
            - 採用区分: `新卒` / `中途`
            - 雇用区分: `正社員` / `契約社員`
            - 行 3 がヘッダー（社員ID, 生年月日, …）の形式にも対応

            ---
            """
        )
        with st.sidebar:
            st.markdown("---")
            if st.button("📋 サンプルデータで試す", use_container_width=True):
                st.session_state["use_sample"] = True
                st.rerun()
        return

    # ── 基準日でデータを補完 ────────────────────────────────────────
    df_all = enrich(raw_df, reference_date)

    # =========================================================================
    # サイドバー: フィルター（データ読み込み後に表示）
    # =========================================================================
    with st.sidebar:
        st.markdown("---")
        st.subheader("🔍 フィルター")
        st.caption("未選択・「全て」の場合は全データが対象になります")

        # フィルター1: 入社年度（複数選択）
        all_fy = sorted([int(y) for y in df_all["入社年度"].dropna().unique()])
        selected_fy = st.multiselect(
            "入社年度（複数選択可）",
            options=all_fy,
            default=[],
            format_func=lambda y: f"{y}年度",
        )

        # フィルター2: 性別
        gender_opts = ["全て"] + sorted(df_all["性別"].dropna().unique().tolist())
        sel_gender = st.selectbox("性別", gender_opts)

        # フィルター3: 採用区分（新卒を先頭に）
        _hire_all = df_all["採用区分"].dropna().unique().tolist()
        _hire_order = ["新卒", "中途"]
        hire_list = [h for h in _hire_order if h in _hire_all]
        hire_list += [h for h in sorted(_hire_all) if h not in _hire_order]
        hire_opts = ["全て"] + hire_list
        sel_hire = st.selectbox("採用区分", hire_opts)

        # フィルター4: 職種区分（表示順固定・未定除外）
        JOB_ORDER = ["営業", "SE", "社会基盤", "ITAD", "EIT", "コーポレート"]
        _job_all = df_all["職種分類"].dropna().unique().tolist()
        job_list = [j for j in JOB_ORDER if j in _job_all]
        job_opts = ["全て"] + job_list
        sel_job = st.selectbox("職種区分", job_opts)

        # フィルター5: 雇用区分（正社員を先頭に）
        _emp_all = df_all["雇用区分"].dropna().unique().tolist()
        _emp_order = ["正社員", "契約社員"]
        emp_list = [e for e in _emp_order if e in _emp_all]
        emp_list += [e for e in sorted(_emp_all) if e not in _emp_order]
        emp_opts = ["全て"] + emp_list
        sel_emp = st.selectbox("雇用区分", emp_opts)

        # フィルター6: 退職区分
        sel_exit = st.selectbox(
            "退職区分",
            ["全て", "通常退職", "自然退職"],
        )

        st.markdown("---")
        # ── 将来予測パラメーター ─────────────────────────────────
        st.subheader("将来予測の設定")
        st.caption("採用・離職の実績ベース期間と想定採用数を調整できます")

        # 直近何年の実績を使うか
        recent_years = st.slider(
            "実績参照期間（直近N年）",
            min_value=1, max_value=10, value=5,
            help="採用数・離職率の平均を計算する直近年数。大きいほど長期トレンドを反映"
        )

        # 採用数の実績計算値を表示
        hire_by_fy = df_all.groupby("入社年度").size().sort_index()
        ref_fy_now = get_fiscal_year(reference_date)
        hire_excl = hire_by_fy[hire_by_fy.index < ref_fy_now]
        calc_avg_hire = float(hire_excl.tail(recent_years).mean()) if len(hire_excl) > 0 else 0
        st.caption(f"直近{recent_years}年の平均採用数: **{calc_avg_hire:.0f}人/年**")

        # 採用数の手動上書き（任意）
        use_custom_hire = st.checkbox("採用数を手動で指定する")
        override_hire = None
        if use_custom_hire:
            override_hire = float(st.number_input(
                "年間想定採用数（人）",
                min_value=0, max_value=200,
                value=int(round(calc_avg_hire)),
                step=1
            ))

        st.markdown("---")
        st.caption(f"総レコード数: {len(df_all):,} 件")

    # フィルター適用
    filters = {
        "入社年度": selected_fy,
        "性別": sel_gender,
        "採用区分": sel_hire,
        "職種分類": sel_job,
        "雇用区分": sel_emp,
        "退職区分": sel_exit,
    }
    df = apply_filters(df_all, filters)

    # =========================================================================
    # KPI セクション
    # =========================================================================
    kpis = compute_kpis(df, reference_date)

    st.subheader("📊 KPI サマリー")
    st.info(
        f"基準日: {reference_date}　｜　"
        f"参照年度（直近完了）: {kpis['参照年度']}年度　｜　"
        f"フィルター適用後レコード: {len(df):,}件　｜　"
        f"期首在籍者数: {kpis['期首在籍']}人"
    )

    # 1行目 KPI
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("在籍者数", f"{kpis['在籍者数']:,}人")
    c2.metric("総離職率", f"{kpis['総離職率']}%")
    c3.metric("通常離職率", f"{kpis['通常離職率']}%")
    c4.metric("自然退職率", f"{kpis['自然退職率']}%")
    c5.metric("平均勤続年数", f"{kpis['平均勤続年数']}年")

    # 2行目 KPI（早期離職率）
    c6, c7, c8, _ = st.columns(4)
    c6.metric("1年以内早期離職率", f"{kpis['早期離職_1年']}%")
    c7.metric("3年以内早期離職率", f"{kpis['早期離職_3年']}%")
    c8.metric("5年以内早期離職率", f"{kpis['早期離職_5年']}%")

    st.markdown("---")

    # =========================================================================
    # 年度別サマリーテーブル（展開可）
    # =========================================================================
    with st.expander("📋 年度別サマリーテーブル（展開）"):
        summary = build_yearly_summary(df)
        st.dataframe(summary, use_container_width=True, hide_index=True)
        st.caption(
            "在籍者数=年度末時点, 期首在籍=その年度4/1時点, "
            "離職率=退職者数÷期首在籍者数×100"
        )

    # =========================================================================
    # グラフセクション
    # =========================================================================

    # ── 在籍者数推移 / 人口ピラミッド ───────────────────────────
    st.subheader("在籍者数の推移と年代構成")
    col1, col2 = st.columns(2)
    with col1:
        st.plotly_chart(fig_active_trend(df), use_container_width=True)
    with col2:
        st.plotly_chart(fig_population_pyramid(df), use_container_width=True)

    # ── 離職率推移 / 採用人数推移 ────────────────────────────────
    st.subheader("離職率と採用人数の推移")
    col3, col4 = st.columns(2)
    with col3:
        st.plotly_chart(fig_turnover_trend(df), use_container_width=True)
    with col4:
        st.plotly_chart(fig_hire_trend(df), use_container_width=True)

    # ── 定着率コホート（新卒 / 中途） ────────────────────────────
    st.subheader("定着率（コホート分析）")
    col5, col6 = st.columns(2)
    with col5:
        st.plotly_chart(
            fig_cohort_retention(df, "新卒", reference_date), use_container_width=True
        )
    with col6:
        st.plotly_chart(
            fig_cohort_retention(df, "中途", reference_date), use_container_width=True
        )

    # ── 勤続年数分布 / 年代別離職率 ─────────────────────────────
    st.subheader("勤続年数分布と年代別離職率")
    col7, col8 = st.columns(2)
    with col7:
        st.plotly_chart(fig_tenure_hist(df), use_container_width=True)
    with col8:
        st.plotly_chart(fig_age_group_turnover(df), use_container_width=True)

    # ── 早期離職率 / 職種別離職率 ────────────────────────────────
    st.subheader("早期離職率と職種別分析")
    col9, col10 = st.columns(2)
    with col9:
        st.plotly_chart(fig_early_attrition(df), use_container_width=True)
    with col10:
        st.plotly_chart(fig_job_turnover(df), use_container_width=True)

    # ── 男女別割合 ──────────────────────────────────────────────
    st.subheader("男女別在籍割合と採用割合")
    st.plotly_chart(fig_gender_ratios(df), use_container_width=True)

    # ── コホートヒートマップ ──────────────────────────────────────
    st.subheader("採用年度 × 離職年度ヒートマップ")
    st.plotly_chart(fig_cohort_heatmap(df), use_container_width=True)

    # ── 将来予測 ─────────────────────────────────────────────────
    st.subheader("将来退職予測と人員数予測")
    st.caption(f"定年年齢: {RETIREMENT_AGE}歳 ／ 予測期間: 基準日から5年間")
    col12, col13 = st.columns(2)
    with col12:
        st.plotly_chart(
            fig_retirement_forecast(df, reference_date,
                                    recent_years=recent_years,
                                    override_hire=override_hire),
            use_container_width=True)
    with col13:
        st.plotly_chart(
            fig_headcount_forecast(df, reference_date,
                                   recent_years=recent_years,
                                   override_hire=override_hire),
            use_container_width=True)

    # ── データプレビュー ──────────────────────────────────────
    with st.expander("🔍 データプレビュー（フィルター適用後・先頭100件）"):
        preview_cols = [
            c for c in [
                "社員ID", "性別", "採用区分", "雇用区分", "職種分類",
                "入社年度", "退職年度", "年齢", "年代",
                "勤続年数", "退職区分", "在籍フラグ",
            ]
            if c in df.columns
        ]
        st.dataframe(
            df[preview_cols].head(100).reset_index(drop=True),
            use_container_width=True,
        )
        st.caption(f"フィルター適用後総件数: {len(df):,} 件")


# =============================================================================
# サンプルデータ生成（テスト用）
# =============================================================================


def _generate_sample_data():
    """
    テスト・デモ用のサンプル社員データを生成する
    実際の運用では不要（Excelファイルをアップロードしてください）
    """
    rng = np.random.default_rng(42)
    n = 500

    genders = rng.choice(["男性", "女性"], n, p=[0.6, 0.4])
    job_types = rng.choice(["SE", "営業", "管理", "企画", "コーポレート"], n)
    hire_types = rng.choice(["新卒", "中途"], n, p=[0.55, 0.45])
    emp_types = rng.choice(["正社員", "契約社員"], n, p=[0.85, 0.15])

    hire_years = rng.integers(2000, 2025, n)
    hire_months = rng.choice([4, 7, 10, 1], n)
    hire_dates = [date(int(y), int(m), 1) for y, m in zip(hire_years, hire_months)]

    birth_dates = []
    for hd in hire_dates:
        age_at_hire = int(rng.integers(22, 50))
        by = hd.year - age_at_hire
        bm = int(rng.integers(1, 13))
        bd = int(rng.integers(1, 29))
        try:
            birth_dates.append(date(by, bm, bd))
        except ValueError:
            birth_dates.append(date(by, bm, 1))

    base_ref = date(2026, 3, 31)
    exit_dates = []
    for hd in hire_dates:
        if rng.random() < 0.35:
            yrs = int(rng.integers(1, 20))
            ed = date(hd.year + yrs, hd.month, 1)
            exit_dates.append(ed if ed <= base_ref else None)
        else:
            exit_dates.append(None)

    return pd.DataFrame(
        {
            "社員ID": [f"E{str(i+1).zfill(5)}" for i in range(n)],
            "生年月日": birth_dates,
            "性別": genders,
            "入社年月日": hire_dates,
            "退職年月日": exit_dates,
            "職種分類": job_types,
            "採用区分": hire_types,
            "雇用区分": emp_types,
        }
    )


# =============================================================================
# エントリーポイント
# =============================================================================

if __name__ == "__main__":
    main()
