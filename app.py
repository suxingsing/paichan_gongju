import streamlit as st
from datetime import datetime, timedelta
import pandas as pd
import io

# ------------------------------
# 核心排产计算函数（完全匹配你的需求）
# ------------------------------
def calculate_schedule(
    uph, total_demand, initial_stock, work_hours_per_shift,
    start_date, target_days, end_date, exclude_sunday, custom_holidays, schedule_mode
):
    # 1. 基础参数计算
    net_demand = max(total_demand - initial_stock, 0)
    single_shift_daily_cap = uph * work_hours_per_shift  # 单班组单日产能

    # 边界情况处理
    if net_demand == 0:
        return pd.DataFrame(), 0, 0, "期初库存已完全覆盖总需求，无需排产", 0, 0
    if single_shift_daily_cap <= 0:
        return pd.DataFrame(), 0, 0, "错误：当前配置的单班组单日产能为0，请检查UPH和工时设置", 0, 0

    schedule_list = []
    cumulative_capacity = 0
    holiday_set = set(pd.to_datetime(custom_holidays).date) if custom_holidays else set()

    # ------------------------------
    # 模式1：正排模式（从开工日期往后推）
    # ------------------------------
    if schedule_mode == "正排模式（从开工日期往后推）":
        current_date = start_date
        while cumulative_capacity < net_demand:
            is_workday = True
            if exclude_sunday and current_date.weekday() == 6:
                is_workday = False
            if current_date in holiday_set:
                is_workday = False
            
            if is_workday:
                cumulative_capacity += single_shift_daily_cap
                schedule_list.append({
                    "排产日期": current_date.strftime("%Y-%m-%d"),
                    "星期": ["周一", "周二", "周三", "周四", "周五", "周六", "周日"][current_date.weekday()],
                    "班组1是否生产": "是",
                    "班组2是否生产": "否",
                    "当日总产能": single_shift_daily_cap,
                    "累计总产能": cumulative_capacity,
                    "需求完成状态": "已完成" if cumulative_capacity >= net_demand else "排产中"
                })
            
            current_date += timedelta(days=1)
        extra_shift_days = 0
        total_workdays = len(schedule_list)

    # ------------------------------
    # 模式2：倒排模式（从交付截止日期往前推）
    # ------------------------------
    elif schedule_mode == "倒排模式（从交付截止日期往前推）":
        current_date = end_date
        while cumulative_capacity < net_demand:
            is_workday = True
            if exclude_sunday and current_date.weekday() == 6:
                is_workday = False
            if current_date in holiday_set:
                is_workday = False
            
            if is_workday:
                cumulative_capacity += single_shift_daily_cap
                schedule_list.append({
                    "排产日期": current_date.strftime("%Y-%m-%d"),
                    "星期": ["周一", "周二", "周三", "周四", "周五", "周六", "周日"][current_date.weekday()],
                    "班组1是否生产": "是",
                    "班组2是否生产": "否",
                    "当日总产能": single_shift_daily_cap,
                    "累计总产能": cumulative_capacity,
                    "需求完成状态": "已完成" if cumulative_capacity >= net_demand else "排产中"
                })
            
            current_date -= timedelta(days=1)
            if current_date < datetime(1970, 1, 1).date():
                return pd.DataFrame(), 0, 0, "错误：交付截止日期前的可用工作日产能不足，无法满足需求", 0, 0

        schedule_list.reverse()
        extra_shift_days = 0
        total_workdays = len(schedule_list)

    # ------------------------------
    # 模式3：指定天数排产（班组1干满，班组2补前x天）
    # ------------------------------
    elif schedule_mode == "指定天数排产（班组1干满，班组2补前x天）":
        # 第一步：计算目标天数内的可用工作日
        current_date = start_date
        workday_list = []
        day_count = 0
        
        while day_count < target_days:
            is_workday = True
            if exclude_sunday and current_date.weekday() == 6:
                is_workday = False
            if current_date in holiday_set:
                is_workday = False
            
            if is_workday:
                workday_list.append(current_date)
            day_count += 1
            current_date += timedelta(days=1)
        
        total_workdays = len(workday_list)
        if total_workdays == 0:
            return pd.DataFrame(), 0, 0, "错误：目标天数内没有可用工作日，请调整开始日期或排除规则", 0, 0

        # 第二步：核心计算（完全匹配你的需求）
        # 1. 班组1干满所有工作日的总产能
        shift1_full_total = total_workdays * single_shift_daily_cap
        
        # 2. 判断是否需要班组2
        if shift1_full_total >= net_demand:
            # 班组1自己就能完成，不需要班组2
            extra_shift_days = 0
            for workday in workday_list:
                if cumulative_capacity >= net_demand:
                    # 即使班组1干满了，但如果提前完成了，后面的天数还是要排（因为你要求班组1干满）
                    # 这里按你的要求：班组1必须干满所有天数
                    pass
                
                daily_cap = single_shift_daily_cap
                cumulative_capacity += daily_cap
                
                schedule_list.append({
                    "排产日期": workday.strftime("%Y-%m-%d"),
                    "星期": ["周一", "周二", "周三", "周四", "周五", "周六", "周日"][workday.weekday()],
                    "班组1是否生产": "是",
                    "班组2是否生产": "否",
                    "当日总产能": daily_cap,
                    "累计总产能": cumulative_capacity,
                    "需求完成状态": "已完成" if cumulative_capacity >= net_demand else "排产中"
                })
        else:
            # 班组1自己不够，需要班组2在前x天同时生产
            gap = net_demand - shift1_full_total
            # 计算班组2需要干的天数（向上取整）
            extra_shift_days = (gap + single_shift_daily_cap - 1) // single_shift_daily_cap
            
            # 生成排产表
            for day_index, workday in enumerate(workday_list):
                # 班组1每天都生产
                shift1_production = single_shift_daily_cap
                # 班组2只在前extra_shift_days天生产
                shift2_production = single_shift_daily_cap if (day_index < extra_shift_days) else 0
                
                daily_cap = shift1_production + shift2_production
                cumulative_capacity += daily_cap
                
                schedule_list.append({
                    "排产日期": workday.strftime("%Y-%m-%d"),
                    "星期": ["周一", "周二", "周三", "周四", "周五", "周六", "周日"][workday.weekday()],
                    "班组1是否生产": "是",
                    "班组2是否生产": "是" if day_index < extra_shift_days else "否",
                    "当日总产能": daily_cap,
                    "累计总产能": cumulative_capacity,
                    "需求完成状态": "已完成" if cumulative_capacity >= net_demand else "排产中"
                })

    # 结果整理
    schedule_df = pd.DataFrame(schedule_list)
    final_total_capacity = cumulative_capacity

    return schedule_df, total_workdays, final_total_capacity, "排产计算完成", extra_shift_days, net_demand

# ------------------------------
# 网页界面配置
# ------------------------------
st.set_page_config(page_title="量产排产自动化计算工具", page_icon="📅", layout="wide")
st.title("📅 量产排产自动化计算工具")
st.divider()

# ------------------------------
# 1. 排产模式与基础参数配置
# ------------------------------
st.subheader("一、排产模式与核心参数配置")
# 模式选择
schedule_mode = st.radio(
    "排产模式选择",
    options=[
        "正排模式（从开工日期往后推）", 
        "倒排模式（从交付截止日期往前推）",
        "指定天数排产（班组1干满，班组2补前x天）"
    ],
    index=2,
    horizontal=True
)

# 日期配置
date_col1, date_col2, date_col3 = st.columns(3)
with date_col1:
    start_date = st.date_input("排产开工日期", value=datetime.today().date())
with date_col2:
    if schedule_mode == "倒排模式（从交付截止日期往前推）":
        end_date = st.date_input("交付截止日期", value=datetime.today().date() + timedelta(days=30))
    else:
        end_date = None
    if "指定天数" in schedule_mode:
        target_days = st.number_input("目标完成天数（自然日，班组1必须干满）", min_value=1, value=15, step=1)
    else:
        target_days = None

# 核心需求参数
base_col1, base_col2, base_col3, base_col4 = st.columns(4)
with base_col1:
    total_demand = st.number_input("总需求量（件）", min_value=0, value=160000, step=1000)
with base_col2:
    initial_stock = st.number_input("期初库存（件）", min_value=0, value=6000, step=1000)
with base_col3:
    exclude_sunday = st.checkbox("自动排除周日", value=True)
with base_col4:
    custom_holidays = st.date_input(
        "自定义节假日（多选）",
        value=[],
        help="按住Ctrl键可多选多个日期"
    )

st.divider()

# ------------------------------
# 2. 班组与产能参数配置
# ------------------------------
st.subheader("二、班组与产能参数配置")
shift_col1, shift_col2 = st.columns(2)

with shift_col1:
    uph = st.number_input("单班组UPH（单位小时产量）", min_value=0, value=400, step=10)
with shift_col2:
    work_hours_per_shift = st.number_input("单班组单日工作小时数", min_value=1, value=10, step=1)

st.divider()

# ------------------------------
# 3. 排产计算与结果展示
# ------------------------------
st.subheader("三、排产计算结果")
calc_col1, calc_col2 = st.columns([1, 5])
with calc_col1:
    calc_button = st.button("开始排产计算", type="primary", use_container_width=True)

if calc_button:
    schedule_df, total_workdays, final_capacity, message, extra_shift_days, net_demand = calculate_schedule(
        uph=uph,
        total_demand=total_demand,
        initial_stock=initial_stock,
        work_hours_per_shift=work_hours_per_shift,
        start_date=start_date,
        target_days=target_days,
        end_date=end_date,
        exclude_sunday=exclude_sunday,
        custom_holidays=custom_holidays,
        schedule_mode=schedule_mode
    )

    if "错误" in message:
        st.error(message)
    else:
        st.success(message)
        if not schedule_df.empty:
            # 关键指标
            metric_col1, metric_col2, metric_col3, metric_col4 = st.columns(4)
            with metric_col1:
                st.metric("本次排产净需求", f"{net_demand:,} 件")
            with metric_col2:
                st.metric("可用工作日总数", f"{total_workdays} 天")
            with metric_col3:
                st.metric("最终排产总产能", f"{final_capacity:,} 件")
            with metric_col4:
                if "指定天数" in schedule_mode:
                    st.metric("班组2需要生产天数", f"{extra_shift_days} 天")

            # 排产表
            st.markdown("#### 详细排产计划表")
            st.dataframe(schedule_df, use_container_width=True, height=400)

            # Excel导出
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                schedule_df.to_excel(writer, sheet_name='排产计划表', index=False)
            st.download_button(
                label="📥 下载排产计划Excel文件",
                data=buffer,
                file_name=f"量产排产计划表_{datetime.now().strftime('%Y%m%d%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )