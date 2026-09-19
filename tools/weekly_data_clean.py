import pandas as pd
from pathlib import Path

# ==============================================================================
# 20260503修改：将 agg 字段列表、agg 字典、二级分组规则抽到模块级常量
# 旧代码：agg 字典在 process_xlsx 函数内写了两遍（一级分组和二级分组各一份），
#         subgroup_rules 也硬编码在函数内部
# 新代码：统一定义为模块级常量，函数内直接引用，修改产品规则只需改这里
# ==============================================================================

# 需要 sum 聚合的原始数值列
_AGG_COLS = [
    '订单量', '销量', '销售额', '促销销量', '促销订单量', '促销销售额', '促销折扣',
    '退款量', '退款金额', '展示', '点击', '广告订单量', '广告花费', '广告销售额'
]
# 由 _AGG_COLS 自动生成 agg 字典，避免重复书写
# 旧代码：手动写了两份 {'订单量':'sum', '销量':'sum', ...} 字典
_AGG_DICT = {col: 'sum' for col in _AGG_COLS}

# 二级分组规则：各 listing标签 下按品名关键词模糊匹配
# 旧代码：此字典硬编码在 process_xlsx 函数内部，不易维护
SUBGROUP_RULES = {
    'SL': ['3.6FT', '5.0FT', '5.5FT'],
    'Toy': ['Short w/oBall 1Pack', 'Long w/oBall 1Pack', 'w/Ball 1Pack', 'Long w/oBall 2Pack'],
    'ToyDH': ['1Pack', '2Pack'],
    'DL': ['w/oHandle AL', 'w/Handle AL', 'w/oHandle ZN', 'w/Handle ZN'],
    'DSL': ['3FT', '6FT'],
    'MFL': ['AL', 'ZN'],
    'SFM': ['SFM 1', 'SFM 2', 'SFM 3'],
    'SC': ['SC S', 'SC M', 'SC L'],
    'VHAN1': ['VHAN1 S', 'VHAN1 M', 'VHAN1 L'],
    'VHAN2': ['VHAN2 S', 'VHAN2 M', 'VHAN2 L', 'VHAN2 XL']
}


# ==============================================================================
# 20260503修改：派生指标计算抽成独立函数，消除一级/二级分组中的重复逻辑
# 旧代码：退款率/CPC/促销折扣 的计算在 grouped 和 sub_group 两处各写了一遍
# 新代码：统一调用 _calc_derived()，DataFrame 和 dict 两种形态分别处理
# ==============================================================================

def _calc_derived_df(df):
    """对 DataFrame 做向量化派生指标计算（替代逐行 apply+lambda，性能更好）"""
    # 旧代码（逐行 apply，数据量大时慢）：
    # grouped['退款率'] = grouped.apply(lambda r: r['退款量']/r['销量'] if r['销量']>0 else 0, axis=1)
    # grouped['CPC']   = grouped.apply(lambda r: r['广告花费']/r['点击'] if r['点击']>0 else 0, axis=1)
    # grouped['促销折扣'] = grouped.apply(lambda r: r['促销销售额']/r['促销销量'] if r['促销销量']>0 else 0, axis=1)
    df['退款率'] = df['退款量'].div(df['销量'].replace(0, float('nan'))).fillna(0)
    df['CPC'] = df['广告花费'].div(df['点击'].replace(0, float('nan'))).fillna(0)
    return df


def _calc_derived_dict(d):
    """对二级分组汇总后的 dict 计算派生指标"""
    # 旧代码（直接写在循环体内，与一级分组逻辑重复）：
    # sub_group['退款率'] = sub_group['退款量']/sub_group['销量'] if sub_group['销量'] > 0 else 0
    # sub_group['CPC']   = sub_group['广告花费']/sub_group['点击'] if sub_group['点击'] > 0 else 0
    # sub_group['促销折扣'] = sub_group['促销销售额']/sub_group['促销销量'] if sub_group['促销销量'] > 0 else 0
    d['退款率'] = d['退款量'] / d['销量'] if d['销量'] > 0 else 0
    d['CPC'] = d['广告花费'] / d['点击'] if d['点击'] > 0 else 0
    return d


# 父ASIN 文件中需要聚合的六列
_PARENT_ASIN_COLS = [
    'Sessions-Browser', 'Sessions-Mobile', 'Sessions-Total',
    'PV-Browser', 'PV-Mobile', 'PV-Total'
]


def _load_parent_asin(parent_asin_file) -> pd.DataFrame:
    """
    读取父ASIN文件，按 listing标签 聚合六列流量数据。
    - 跳过 listing标签 为空的行
    - 返回以 listing标签 为索引的 DataFrame
    - 文件不存在时返回空 DataFrame（调用方按全部留空处理）
    """
    if parent_asin_file is None or not Path(parent_asin_file).exists():
        return pd.DataFrame(columns=_PARENT_ASIN_COLS)

    df = pd.read_excel(parent_asin_file)
    df.columns = df.columns.str.strip()

    # 跳过 listing标签 为空的行
    df = df[df['listing标签'].notna() & (df['listing标签'].astype(str).str.strip() != '')]
    if df.empty:
        return pd.DataFrame(columns=_PARENT_ASIN_COLS)

    df['listing标签'] = df['listing标签'].astype(str).str.strip()

    # 六列转数值
    for col in _PARENT_ASIN_COLS:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
        else:
            df[col] = 0

    aggregated = df.groupby('listing标签')[_PARENT_ASIN_COLS].sum()
    return aggregated  # index = listing标签


def process_xlsx(input_file, output_file, parent_asin_file=None):
    # 读取数据
    df = pd.read_excel(input_file, dtype=str).fillna('')

    # 列名清理，确保列名无空格
    df.columns = df.columns.str.strip()

    # 必须确保这些列存在，否则报错
    required_cols = ['listing标签', '品名', '订单量', '销量', '销售额', '促销销量', '促销订单量',
                     '促销销售额', '退款量', '退款金额', '展示', '点击', '广告订单量', '广告花费', '广告销售额']
    for col in required_cols:
        if col not in df.columns:
            raise ValueError(f"缺少必要列: {col}")

    # ==============================================================================
    # 20260503修改：to_num 中空字符串处理改用语义更清晰的 fillna
    # 旧代码：s = series.str.replace(',','').replace('', '0')
    #         第二个 .replace 是 Series.replace（值替换），与前面的 str.replace 混用，语义不一致
    # 新代码：统一用 .fillna('0') 补空，意图更明确
    # ==============================================================================
    def to_num(series, is_int=False):
        # 旧代码：s = series.str.replace(',','').replace('', '0')
        s = series.str.replace(',', '').fillna('0').replace('', '0')
        if is_int:
            return pd.to_numeric(s, errors='coerce').fillna(0).astype(int)
        else:
            return pd.to_numeric(s, errors='coerce').fillna(0).astype(float)

    int_cols = ['订单量', '销量', '促销销量', '促销订单量', '退款量', '展示', '点击', '广告订单量']
    float_cols = ['销售额', '促销销售额', '促销折扣', '退款金额', '广告花费', '广告销售额']

    for col in int_cols:
        df[col] = to_num(df[col].astype(str), True)
    for col in float_cols:
        df[col] = to_num(df[col].astype(str), False)

    # 品名去空格
    df['品名'] = df['品名'].str.strip()
    df['listing标签'] = df['listing标签'].str.strip()

    # 退款金额和广告花费转正数（绝对值）
    df['退款金额'] = df['退款金额'].abs()
    df['广告花费'] = df['广告花费'].abs()

    # --- 一级分组：按 listing标签 ---
    # 20260503修改：agg 字典改用模块级常量 _AGG_DICT，不再手动列举
    # 旧代码：df.groupby('listing标签').agg({'订单量':'sum', '销量':'sum', ...（13列手写）})
    grouped = df.groupby('listing标签').agg(_AGG_DICT).reset_index()

    # 20260503修改：派生指标改用向量化函数，不再逐行 apply+lambda
    grouped = _calc_derived_df(grouped)

    # --- 合并父ASIN流量数据（Sessions / PV）---
    # 20260506新增：读取父ASIN文件，按 listing标签 聚合后 left join 到一级汇总
    # 父ASIN文件不存在时 df_parent 为空，六列全部留空（NaN）
    df_parent = _load_parent_asin(parent_asin_file)
    if not df_parent.empty:
        grouped = grouped.join(df_parent, on='listing标签', how='left')
        print(f"✅ 已合并父ASIN流量数据，匹配标签: {df_parent.index.tolist()}")
    else:
        for col in _PARENT_ASIN_COLS:
            grouped[col] = None
        if parent_asin_file is not None:
            print(f"⚠️ 父ASIN文件无有效数据，六列留空")
        else:
            print(f"ℹ️ 未提供父ASIN文件，六列留空")

    # --- 二级分组：按品名规格 ---
    # 20260503修改：subgroup_rules 改从模块级常量 SUBGROUP_RULES 读取
    subgroup_results = []
    for listing_tag, names in SUBGROUP_RULES.items():
        df_filtered = df[df['listing标签'] == listing_tag]
        if df_filtered.empty:
            continue
        for name in names:
            # 模糊匹配品名包含规则中的name，忽略大小写
            mask = df_filtered['品名'].str.contains(name, case=False, na=False)
            df_sub = df_filtered[mask]
            if df_sub.empty:
                continue

            # 20260503修改：agg 字典改用模块级常量 _AGG_DICT
            sub_group = df_sub.agg(_AGG_DICT).to_dict()

            # 绝对值（二级分组原始数据可能含负数，与一级分组对齐）
            sub_group['退款金额'] = abs(sub_group['退款金额'])
            sub_group['广告花费'] = abs(sub_group['广告花费'])

            # 20260503修改：派生指标改用统一函数，不再重复写三行计算逻辑
            sub_group = _calc_derived_dict(sub_group)

            # 填充标签
            sub_group['listing标签'] = listing_tag
            sub_group['品名'] = name

            subgroup_results.append(sub_group)

    df_subgroups_all = pd.DataFrame(subgroup_results)

    # --- 数据格式化 ---
    def format_data(df_out):
        # 整数列
        for col in ['订单量', '销量', '促销销量', '促销订单量', '退款量', '展示', '点击', '广告订单量']:
            if col in df_out.columns:
                df_out[col] = df_out[col].fillna(0).astype(int)
        # 保留两位小数列
        for col in ['销售额', '促销销售额', '促销折扣', '退款金额', '广告花费', '广告销售额', 'CPC']:
            if col in df_out.columns:
                df_out[col] = df_out[col].fillna(0).round(2)
        # 百分比列，保留两位小数
        if '退款率' in df_out.columns:
            df_out['退款率'] = df_out['退款率'].fillna(0).round(4)  # 先四位小数
            # df_out['退款率'] = df_out['退款率'] * 100  # 转换成百分比
            # df_out['退款率'] = df_out['退款率'].round(2)
        return df_out

    grouped = format_data(grouped)
    df_subgroups_all = format_data(df_subgroups_all)

    # --- 导出 ---
    # 指定导出列顺序
    # 20260506新增：标签汇总sht 末尾追加六列父ASIN流量数据
    export_columns = [
        'listing标签', '品名', '销量', '订单量', '销售额', '促销销量', '促销订单量', '促销销售额', '促销折扣',
        '退款量', '退款率', '退款金额', '展示', '点击', '广告订单量', '广告花费', '广告销售额', 'CPC',
        'Sessions-Browser', 'Sessions-Mobile', 'Sessions-Total',
        'PV-Browser', 'PV-Mobile', 'PV-Total'
    ]

    # 只保留并按顺序导出指定列（如果有些列不存在则自动跳过）
    grouped_export = grouped[[col for col in export_columns if col in grouped.columns]]
    df_subgroups_all_export = df_subgroups_all[[col for col in export_columns if col in df_subgroups_all.columns]]

    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        grouped_export.to_excel(writer, sheet_name='标签汇总sht', index=False)
        if not df_subgroups_all_export.empty:
            df_subgroups_all_export.to_excel(writer, sheet_name='标签品名汇总sht', index=False)

        # 设置Excel格式（整数、浮点、百分比）
        workbook = writer.book
        int_fmt = workbook.add_format({'num_format': '0'})
        float_fmt = workbook.add_format({'num_format': '0.00'})
        pct_fmt = workbook.add_format({'num_format': '0.00%'})

        # 格式化第一个sheet
        ws1 = writer.sheets['标签汇总sht']
        fmt_map = {
            '订单量': int_fmt, '销量': int_fmt, '销售额': float_fmt,
            '促销销量': int_fmt, '促销订单量': int_fmt, '促销销售额': float_fmt, '促销折扣': float_fmt,
            '退款量': int_fmt, '退款率': pct_fmt, '退款金额': float_fmt,
            '展示': int_fmt, '点击': int_fmt, '广告订单量': int_fmt,
            '广告花费': float_fmt, '广告销售额': float_fmt, 'CPC': float_fmt,
            # 20260506新增：父ASIN流量列格式（整数）
            'Sessions-Browser': int_fmt, 'Sessions-Mobile': int_fmt, 'Sessions-Total': int_fmt,
            'PV-Browser': int_fmt, 'PV-Mobile': int_fmt, 'PV-Total': int_fmt,
        }
        for col_num, col_name in enumerate(grouped_export.columns):
            if col_name in fmt_map:
                ws1.set_column(col_num, col_num, 15, fmt_map[col_name])
            else:
                ws1.set_column(col_num, col_num, 20)

        # 格式化第二个sheet
        if not df_subgroups_all_export.empty:
            ws2 = writer.sheets['标签品名汇总sht']
            for col_num, col_name in enumerate(df_subgroups_all_export.columns):
                if col_name in fmt_map:
                    ws2.set_column(col_num, col_num, 15, fmt_map[col_name])
                else:
                    ws2.set_column(col_num, col_num, 20)

    print(f"处理完成，结果已保存至 {output_file}")


def run():
    base_path = Path(__file__).resolve().parents[1]
    data_dir = base_path / "data" / "weekly_data_clean"
    input_dir = data_dir / "input"
    output_dir = data_dir / "output"

    input_dir.mkdir(parents=True, exist_ok=True)
    output_dir.mkdir(parents=True, exist_ok=True)

    print(f"\n📂 输入文件目录: {input_dir}")
    print(f"📂 输出文件目录: {output_dir}")

    # input_file = input("请输入输入文件名（如 产品表现MSKU-W39.xlsx）: ").strip()
    # output_file = input("请输入输出文件名（如 output_aggregated-W39.xlsx）: ").strip()
    week_number = input("请输入国家-周数（如 US-26W11）: ").strip()
    input_file = '产品表现MSKU-' + week_number + '.xlsx'
    output_file = 'output_aggregated-' + week_number + '.xlsx'

    input_path = input_dir / input_file
    output_path = output_dir / output_file

    if not input_path.exists():
        print(f"❌ 文件不存在: {input_path}")
        return

    # 20260506新增：自动查找同周次的父ASIN文件（可选），找不到时六列留空
    parent_asin_file_name = '产品表现父ASIN-' + week_number + '.xlsx'
    parent_asin_path = input_dir / parent_asin_file_name
    if parent_asin_path.exists():
        print(f"📄 找到父ASIN文件: {parent_asin_file_name}")
    else:
        print(f"ℹ️ 未找到父ASIN文件（{parent_asin_file_name}），Sessions/PV 列将留空")
        parent_asin_path = None

    process_xlsx(input_path, output_path, parent_asin_file=parent_asin_path)
    print(f"✅ 周报已生成：{output_path}")