# -*- coding: utf-8 -*-
"""
个人记账 APP：Streamlit 本地记账工具。

功能：
- 收支录入：区分收入/支出，支持分类选择、金额、交易日期、备注
- 数据持久化：使用 pandas + CSV 存储账单，页面重启不丢失
- 账单管理：展示表格、单条删除、筛选、分页
- 预算设置：支持支出分类月度预算，显示剩余、超支红色提醒
- 可视化：月度收支饼图、近 30 天消费趋势折线图、支出分类柱状图
- 导出 Excel、清空账单

说明：首次运行时会自动生成 transactions.csv 和 budgets.csv，无需手动创建。"""

import os
import uuid
from io import BytesIO
from datetime import datetime, date, timedelta

import pandas as pd
import plotly.express as px
import streamlit as st

# ------------ 常量配置 ------------
CSV_FILE = "transactions.csv"  # 账单存储文件
BUDGET_FILE = "budgets.csv"  # 预算存储文件
DATE_FORMAT = "%Y-%m-%d"  # 交易日期格式
EXPENSE_CATEGORIES = ["餐饮", "交通", "购物", "学习", "娱乐", "医疗"]
INCOME_CATEGORIES = ["工资", "兼职"]
ALL_CATEGORIES = EXPENSE_CATEGORIES + INCOME_CATEGORIES
PAGE_SIZE = 10  # 分页每页条数

# Streamlit 页面配置
st.set_page_config(page_title="个人记账APP", layout="wide")


@st.cache_data
def ensure_files_exist():
    """确保 transactions.csv 和 budgets.csv 存在。"""
    if not os.path.exists(CSV_FILE):
        pd.DataFrame(
            columns=["id", "date", "type", "category", "amount", "note", "created_at"]
        ).to_csv(CSV_FILE, index=False, encoding="utf-8-sig")
    if not os.path.exists(BUDGET_FILE):
        pd.DataFrame(
            {"category": EXPENSE_CATEGORIES, "budget": [0.0] * len(EXPENSE_CATEGORIES)}
        ).to_csv(BUDGET_FILE, index=False, encoding="utf-8-sig")
    return True


def clear_cache():
    """清理 Streamlit 缓存，保证数据修改后页面能及时刷新。"""
    try:
        st.cache_data.clear()
    except Exception:
        pass


@st.cache_data
def load_transactions():
    """读取本地 CSV 账单数据，并转换日期字段。"""
    ensure_files_exist()
    df = pd.read_csv(CSV_FILE, encoding="utf-8-sig")
    if df.empty:
        return pd.DataFrame(
            columns=["id", "date", "type", "category", "amount", "note", "created_at"]
        )
    df["date"] = pd.to_datetime(df["date"], errors="coerce").dt.date
    if "created_at" in df.columns:
        df["created_at"] = pd.to_datetime(df["created_at"], errors="coerce")
    return df


def save_transactions(df: pd.DataFrame):
    """保存账单数据到 CSV。"""
    df.to_csv(CSV_FILE, index=False, encoding="utf-8-sig")
    clear_cache() 


@st.cache_data
def load_budgets():
    """读取预算数据，并补齐所有支出分类。"""
    ensure_files_exist()
    bd = pd.read_csv(BUDGET_FILE, encoding="utf-8-sig")
    for cat in EXPENSE_CATEGORIES:
        if cat not in bd["category"].values:
            bd = pd.concat(
                [bd, pd.DataFrame({"category": [cat], "budget": [0.0]})],
                ignore_index=True,
            )
    bd = bd.drop_duplicates(subset=["category"], keep="last")
    bd = bd.set_index("category").reindex(EXPENSE_CATEGORIES).reset_index()
    bd["budget"] = bd["budget"].fillna(0.0)
    return bd


def save_budgets(bd: pd.DataFrame):
    """保存预算数据到 CSV。"""
    bd.to_csv(BUDGET_FILE, index=False, encoding="utf-8-sig")
    clear_cache()


def add_transaction(record: dict):
    """追加一条交易记录并保存。"""
    df = load_transactions()
    df = pd.concat([df, pd.DataFrame([record])], ignore_index=True)
    save_transactions(df)


def delete_transaction(tx_id: str):
    """根据 ID 删除账单记录。"""
    df = load_transactions()
    df = df[df["id"] != tx_id]
    save_transactions(df)


def clear_all_transactions():
    """清空所有账单并重置 CSV 文件。"""
    pd.DataFrame(
        columns=["id", "date", "type", "category", "amount", "note", "created_at"]
    ).to_csv(CSV_FILE, index=False, encoding="utf-8-sig")
    clear_cache()


def export_excel(df: pd.DataFrame) -> bytes:
    """导出账单数据为 Excel 文件字节流。"""
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="账单")
    return output.getvalue()


def compute_summary(df: pd.DataFrame) -> dict:
    """计算今日/本周/本月收入、支出和结余。"""
    today = date.today()
    summary = {
        "today_income": 0.0,
        "today_expense": 0.0,
        "week_income": 0.0,
        "week_expense": 0.0,
        "month_income": 0.0,
        "month_expense": 0.0,
    }
    if df.empty:
        return summary
    df = df.copy()
    df["date"] = pd.to_datetime(df["date"], errors="coerce").dt.date
    today_df = df[df["date"] == today]
    summary["today_income"] = today_df[today_df["type"] == "收入"]["amount"].sum()
    summary["today_expense"] = today_df[today_df["type"] == "支出"]["amount"].sum()
    week_start = today - timedelta(days=today.weekday())
    week_df = df[(df["date"] >= week_start) & (df["date"] <= today)]
    summary["week_income"] = week_df[week_df["type"] == "收入"]["amount"].sum()
    summary["week_expense"] = week_df[week_df["type"] == "支出"]["amount"].sum()
    month_start = today.replace(day=1)
    month_df = df[(df["date"] >= month_start) & (df["date"] <= today)]
    summary["month_income"] = month_df[month_df["type"] == "收入"]["amount"].sum()
    summary["month_expense"] = month_df[month_df["type"] == "支出"]["amount"].sum()
    return summary


def filter_transactions(
    df: pd.DataFrame,
    start_date: date,
    end_date: date,
    type_filter: str,
    category_filter: list,
) -> pd.DataFrame:
    """根据日期、收支类型和分类筛选账单。"""
    if df.empty:
        return df
    filtered = df.copy()
    filtered["date"] = pd.to_datetime(filtered["date"], errors="coerce").dt.date
    filtered = filtered[(filtered["date"] >= start_date) & (filtered["date"] <= end_date)]
    if type_filter in ["收入", "支出"]:
        filtered = filtered[filtered["type"] == type_filter]
    if category_filter:
        filtered = filtered[filtered["category"].isin(category_filter)]
    return filtered.sort_values(by=["date", "created_at"], ascending=[False, False])


def render_sidebar():
    """渲染侧边栏：新增账单、预算设置、筛选和导出/清空按钮。"""
    st.sidebar.header("新增账单 / 预算 / 筛选")

    with st.sidebar.form("entry_form"):
        ttype = st.selectbox("收支类型", options=["支出", "收入"], key="entry_type")
        categories = ALL_CATEGORIES
        category = st.selectbox("分类", options=categories, index=0, key="entry_category")
        amount = st.number_input("金额", min_value=0.0, format="%.2f", key="entry_amount")
        tx_date = st.date_input("交易日期", value=date.today(), key="entry_date")
        note = st.text_input("备注（可选）", key="entry_note")
        if st.form_submit_button("添加账单"):
            if amount <= 0:
                st.sidebar.error("请填写大于 0 的金额。")
            else:
                add_transaction(
                    {
                        "id": str(uuid.uuid4()),
                        "date": tx_date.strftime(DATE_FORMAT),
                        "type": ttype,
                        "category": category,
                        "amount": float(amount),
                        "note": note,
                        "created_at": datetime.now().isoformat(),
                    }
                )
                st.sidebar.success("已保存账单，数据写入本地 CSV。")

    st.sidebar.markdown("---")
    st.sidebar.subheader("支出预算设置")
    budgets = load_budgets()
    with st.sidebar.form("budget_form"):
        budget_inputs = {}
        for cat in EXPENSE_CATEGORIES:
            default_value = float(
                budgets.loc[budgets["category"] == cat, "budget"].squeeze()
            ) if cat in budgets["category"].values else 0.0
            budget_inputs[cat] = st.number_input(
                f"{cat} 预算", min_value=0.0, value=default_value, key=f"budget_{cat}"
            )
        if st.form_submit_button("保存预算"):
            save_budgets(
                pd.DataFrame(
                    {"category": list(budget_inputs.keys()), "budget": list(budget_inputs.values())}
                )
            )
            st.sidebar.success("预算已保存。")

    st.sidebar.markdown("---")
    st.sidebar.subheader("筛选条件")
    start_date = st.sidebar.date_input(
        "开始日期", value=date.today() - timedelta(days=30), key="filter_start"
    )
    end_date = st.sidebar.date_input("结束日期", value=date.today(), key="filter_end")
    type_filter = st.sidebar.selectbox(
        "收支类型", options=["全部", "收入", "支出"], index=0, key="filter_type"
    )
    if type_filter == "收入":
        category_options = INCOME_CATEGORIES
    elif type_filter == "支出":
        category_options = EXPENSE_CATEGORIES
    else:
        category_options = ALL_CATEGORIES
    category_filter = st.sidebar.multiselect(
        "分类", options=category_options, default=category_options, key="filter_category"
    )

    st.sidebar.markdown("---")
    df_all = load_transactions()
    if df_all.empty:
        st.sidebar.info("暂无账单数据，添加后即可导出。")
    else:
        st.sidebar.download_button(
            "导出全部为 Excel",
            data=export_excel(df_all),
            file_name=f"账单_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    if st.sidebar.button("一键清空所有账单"):
        if st.sidebar.checkbox("确认清空（不可恢复）", key="confirm_clear"):
            clear_all_transactions()
            st.sidebar.success("已清空所有账单，CSV 已重置。")

    st.sidebar.markdown("---")
    st.sidebar.caption("账单保存于当前目录 transactions.csv，预算保存于 budgets.csv。")
    return start_date, end_date, type_filter, category_filter


def render_summary_cards(summary: dict):
    """渲染顶部统计卡片，展示收入、支出和结余。"""
    col1, col2, col3 = st.columns(3)
    col1.metric("今日收入", f"¥{summary['today_income']:,.2f}")
    col1.metric("今日支出", f"¥{summary['today_expense']:,.2f}")
    col2.metric("本周收入", f"¥{summary['week_income']:,.2f}")
    col2.metric("本周支出", f"¥{summary['week_expense']:,.2f}")
    month_balance = summary['month_income'] - summary['month_expense']
    col3.metric(
        "本月结余",
        f"¥{month_balance:,.2f}",
        delta=(f"¥{month_balance:,.2f}" if month_balance >= 0 else f"-¥{abs(month_balance):,.2f}"),
    )


def render_budget_panel(df: pd.DataFrame, budgets: pd.DataFrame):
    """显示每个支出分类预算、已用和剩余状态。"""
    st.subheader("支出预算概览")
    if df.empty:
        st.info("暂无账单数据，预算剩余无法统计。")
        return
    expense_df = df[df['type'] == '支出']
    spent_series = (
        expense_df.groupby('category')['amount'].sum().reindex(EXPENSE_CATEGORIES).fillna(0.0)
    )
    for _, row in budgets.iterrows():
        category = row['category']
        budget = float(row['budget'])
        used = float(spent_series.get(category, 0.0))
        remain = budget - used
        if remain < 0:
            st.markdown(
                f"**{category}：预算 ¥{budget:,.2f}，已用 ¥{used:,.2f}，剩余 ¥{remain:,.2f}（超支）**"
            )
        else:
            st.write(
                f"{category}：预算 ¥{budget:,.2f}，已用 ¥{used:,.2f}，剩余 ¥{remain:,.2f}"
            )


def render_charts(df: pd.DataFrame):
    """根据筛选结果绘制收支饼图、消费趋势图和支出分类柱状图。"""
    st.subheader("数据可视化")
    if df.empty:
        st.info("当前筛选结果为空，添加或调整筛选条件后可查看图表。")
        return
    df_plot = df.copy()
    df_plot['date'] = pd.to_datetime(df_plot['date'], errors='coerce')
    pie_data = df_plot.groupby('type')['amount'].sum().reset_index()
    fig_pie = px.pie(
        pie_data,
        names='type',
        values='amount',
        title='收支构成饼图',
        hole=0.4,
        color='type',
        color_discrete_map={'收入': '#2ca02c', '支出': '#d62728'},
    )
    end_date = date.today()
    start_date = end_date - timedelta(days=29)
    trend_df = df_plot[
        (df_plot['type'] == '支出')
        & (df_plot['date'].dt.date >= start_date)
        & (df_plot['date'].dt.date <= end_date)
    ]
    trend_series = (
        trend_df.groupby(trend_df['date'].dt.date)['amount']
        .sum()
        .reindex(pd.date_range(start_date, end_date), fill_value=0)
    )
    fig_line = px.line(
        x=trend_series.index,
        y=trend_series.values,
        labels={'x': '日期', 'y': '支出金额'},
        title='近 30 天支出趋势',
    )
    expense_by_cat = (
        df_plot[df_plot['type'] == '支出']
        .groupby('category')['amount']
        .sum()
        .reindex(EXPENSE_CATEGORIES)
        .fillna(0.0)
        .reset_index()
    )
    fig_bar = px.bar(
        expense_by_cat,
        x='category',
        y='amount',
        title='各类别支出占比柱状图',
        labels={'amount': '支出金额', 'category': '分类'},
        color='amount',
        color_continuous_scale='OrRd',
    )
    col1, col2 = st.columns(2)
    col1.plotly_chart(fig_pie, use_container_width=True)
    col2.plotly_chart(fig_line, use_container_width=True)
    st.plotly_chart(fig_bar, use_container_width=True)


def render_transaction_table(df: pd.DataFrame):
    """渲染账单列表，支持分页和单条删除。"""
    st.subheader("账单列表")
    if df.empty:
        st.info("当前暂无账单，欢迎添加第一条记录。")
        return
    if 'page' not in st.session_state:
        st.session_state.page = 1
    total_pages = max(1, (len(df) + PAGE_SIZE - 1) // PAGE_SIZE)
    st.write(f"共 {len(df)} 条记录，页码 {st.session_state.page}/{total_pages}")
    paging_cols = st.columns([1, 1, 1])
    if paging_cols[0].button('上一页', key='prev_page') and st.session_state.page > 1:
        st.session_state.page -= 1
    if paging_cols[2].button('下一页', key='next_page') and st.session_state.page < total_pages:
        st.session_state.page += 1
    start_idx = (st.session_state.page - 1) * PAGE_SIZE
    page_df = df.iloc[start_idx:start_idx + PAGE_SIZE].copy()
    page_df['created_at'] = pd.to_datetime(page_df['created_at'], errors='coerce')
    page_df['created_at'] = page_df['created_at'].dt.strftime('%Y-%m-%d %H:%M:%S')
    displayed = page_df[
        ['date', 'type', 'category', 'amount', 'note', 'created_at']
    ].rename(
        columns={
            'date': '交易日期',
            'type': '收支类型',
            'category': '分类',
            'amount': '金额',
            'note': '备注',
            'created_at': '创建时间',
        }
    )
    st.dataframe(displayed, use_container_width=True)
    for row in page_df.itertuples():
        cols = st.columns([5, 1])
        cols[0].markdown(
            f"**{row.date} | {row.type} | {row.category}**  \n金额：¥{row.amount:,.2f}  \n备注：{row.note}"
        )
        if cols[1].button('删除', key=f"delete_{row.id}"):
            delete_transaction(row.id)
            st.success('已删除该条账单，页面将自动刷新。')
            return


def render_main_page():
    """渲染整个应用主页面。"""
    transactions = load_transactions()
    start_date, end_date, type_filter, category_filter = render_sidebar()
    filtered = filter_transactions(transactions, start_date, end_date, type_filter, category_filter)
    summary = compute_summary(transactions)

    st.title('个人记账 APP')
    st.markdown(
        '使用本地 CSV 持久化账单，记录收入/支出，支持筛选、预算、可视化、导出 Excel 和清空功能。'
    )
    render_summary_cards(summary)
    render_budget_panel(transactions, load_budgets())
    render_charts(filtered)
    render_transaction_table(filtered)


if __name__ == '__main__':
    render_main_page()
