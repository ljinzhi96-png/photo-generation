
"""
多人一诉数据可视化（支持同比、环比汇总表）
输入文件格式：包含 子场景、一级业务、二级业务、每月 四列
每月列格式：事件名称（件数），多个事件用顿号或逗号分隔
示例：虎溪富力城占道经营问题（3件）、曾家曾凤路占道经营问题（3件）
"""

import pandas as pd
import matplotlib.pyplot as plt
import re
import numpy as np
from pathlib import Path
import argparse
import warnings
import matplotlib.font_manager as fm

warnings.filterwarnings('ignore')

# 设置中文字体
# plt.rcParams['font.sans-serif'] = ['SimHei', 'Microsoft YaHei', 'WenQuanYi Micro Hei']
# plt.rcParams['axes.unicode_minus'] = False


# 获取当前文件所在目录，并构造字体文件的完整路径
current_dir = Path(__file__).parent
font_path = current_dir / 'fonts' / 'SimHei.ttf'   # 确保文件名一致

# 检查字体文件是否存在
if font_path.exists():
    # 注册字体到 matplotlib 全局字体管理器
    fm.fontManager.addfont(str(font_path))
    # 获取字体属性
    prop = fm.FontProperties(fname=str(font_path))
    # 设置 matplotlib 默认字体为该字体的名称
    plt.rcParams['font.family'] = prop.get_name()
    print(f"成功加载中文字体：{prop.get_name()}")
else:
    # 如果找不到字体文件，则回退到系统通用中文字体（用于本地测试）
    plt.rcParams['font.sans-serif'] = ['WenQuanYi Micro Hei', 'DejaVu Sans', 'SimHei', 'Microsoft YaHei']
    print(f"警告：字体文件 {font_path} 不存在，使用备用字体")

# 解决负号显示问题
plt.rcParams['axes.unicode_minus'] = False


# 子场景映射
SUB_SCENE_MAP = {
    '街面清': '城市综合治理',
    '安居住': '住房城建管理',
    '邻里和': '邻里矛盾调处',
    '市场谐': '商民权益保障',
    '政民安': '公共服务供给'
}
SUB_SCENE_NAMES = list(SUB_SCENE_MAP.keys())


def parse_event_count(cell_value):
    """解析每月列，返回 [(事件名称, 件数), ...]"""
    if pd.isna(cell_value) or cell_value == '':
        return []
    results = []
    parts = re.split(r'[、，]', str(cell_value))
    for part in parts:
        part = part.strip()
        if not part:
            continue
        match = re.search(r'(.+)（(\d+)件[^）]*）', part)
        if match:
            event_name = match.group(1).strip()
            count = int(match.group(2))
            results.append((event_name, count))
        else:
            results.append((part, 1))
    return results


def clean_biz_name(name):
    """清洗二级业务名称：去掉括号内的数字（包括中英文括号）"""
    if not isinstance(name, str):
        return str(name)
    # 删除中文括号内的数字，如（10）、（5）
    # name = re.sub(r'[（(]\d+[）)]', '', name)
    # # 删除英文括号内的数字，如(10)
    # name = re.sub(r'\(\d+\)', '', name)
    # 删除中文括号 （...） 及其内部任意字符
    name = re.sub(r'（[^）]*）', '', name)
    # 删除英文括号 (...) 及其内部任意字符
    name = re.sub(r'\([^)]*\)', '', name)
    name = name.strip()
    # 如果清洗后为空（极少情况），保留原名称
    if not name:
        return str(name).strip()
    return name


def load_summary_data(excel_path):
    """
    加载汇总表，返回四个字典：
    - sub_total: {子场景: 总件数}
    - biz_total: {二级业务(清洗后): 总件数}
    - event_total: {事件名称: 总件数}
    - sub_biz: {子场景: {二级业务(清洗后): 总件数}}
    """
    if excel_path is None or not Path(excel_path).exists():
        return None, None, None, None

    df = pd.read_excel(excel_path, dtype=str,engine='openpyxl')
    df['子场景'] = df['子场景'].ffill()
    df['一级业务'] = df['一级业务'].ffill()
    df['二级业务'] = df['二级业务'].ffill()
    df = df[df['二级业务'].notna() & (df['二级业务'] != '')]

    sub_total = {s: 0 for s in SUB_SCENE_NAMES}
    biz_total = {}
    event_total = {}
    sub_biz = {s: {} for s in SUB_SCENE_NAMES}

    for _, row in df.iterrows():
        sub_scene = row['子场景']
        # 清洗二级业务名称
        raw_second_biz = row['二级业务']
        second_biz = clean_biz_name(raw_second_biz)
        monthly_text = row['每月']
        events = parse_event_count(monthly_text)
        if not events:
            continue
        biz_sum = 0
        for event_name, cnt in events:
            event_total[event_name] = event_total.get(event_name, 0) + cnt
            biz_sum += cnt
            if sub_scene in sub_total:
                sub_total[sub_scene] += cnt
        biz_total[second_biz] = biz_total.get(second_biz, 0) + biz_sum
        if sub_scene in sub_biz:
            sub_biz[sub_scene][second_biz] = sub_biz[sub_scene].get(second_biz, 0) + biz_sum
    return sub_total, biz_total, event_total, sub_biz


def extract_month_from_path(file_path):
    """从文件路径中提取形如 202603 的6位数字，若失败返回'对比期'"""
    if file_path is None:
        return '对比期'
    stem = Path(file_path).stem
    # match = re.search(r'(\d{6})', str(file_path))
    # match = re.match(r'(\d{6})', stem)  # 匹配开头的连续6位数字
    match = re.search(r'(\d{6})', stem)   #
    if match:
        return match.group(1)
    return '对比期'



def plot_overall_fig1(biz_total, output_dir):
    """整体图1：排名前五的二级业务（柱状图）"""
    if not biz_total:
        print("警告：当前月无二级业务数据，跳过整体图1")
        return
    sorted_items = sorted(biz_total.items(), key=lambda x: x[1], reverse=True)[:5]
    if not sorted_items:
        return
    labels, values = zip(*sorted_items)

    plt.figure(figsize=(12, 6))

    bars = plt.bar(labels, values, color='steelblue',width=0.4)
    plt.title('高频事件排名前五的二级业务', fontsize=14)
    plt.xlabel('二级业务', fontsize=12)
    plt.ylabel('件数', fontsize=12)
    plt.xticks(rotation=0, ha='center')
    for bar in bars:
        h = bar.get_height()
        plt.text(bar.get_x() + bar.get_width() / 2., h + 0.5, f'{int(h)}', ha='center', va='bottom')

    max_val = max(values)
    plt.ylim(top=max_val + 5)
    plt.tight_layout()
    plt.savefig(output_dir / '整体分析_图1_高频二级业务.png', dpi=150)
    plt.close()
    print("已生成：整体分析_图1_高频二级业务.png")


def plot_overall_fig2(event_total, output_dir):
    """整体图2：排名前五的高频事件（柱状图）"""
    if not event_total:
        print("警告：当前月无事件数据，跳过整体图2")
        return
    sorted_items = sorted(event_total.items(), key=lambda x: x[1], reverse=True)[:5]
    if not sorted_items:
        return
    labels, values = zip(*sorted_items)

    plt.figure(figsize=(12, 6))
    bars = plt.bar(labels, values, color='coral',width=0.4)
    plt.title('排名前五的高频事件', fontsize=14)
    plt.xlabel('高频事件名称', fontsize=12)
    plt.ylabel('件数', fontsize=12)
    plt.xticks(rotation=10, ha='center')
    for bar in bars:
        h = bar.get_height()
        plt.text(bar.get_x() + bar.get_width() / 2., h + 0.5, f'{int(h)}', ha='center', va='bottom')

    max_val = max(values)
    plt.ylim(top=max_val + 5)
    plt.tight_layout()
    plt.savefig(output_dir / '整体分析_图2_高频事件.png', dpi=150)
    plt.close()
    print("已生成：整体分析_图2_高频事件.png")




def plot_overall_fig3(current_sub, yoy_sub, mom_sub, output_dir):
    """整体图3：五大子场景同环比增长率（双折线图），x轴标签合并当月件数"""
    if not current_sub:
        print("警告：当前月子场景数据为空，跳过整体图3")
        return
    scenes = SUB_SCENE_NAMES
    cur_counts = [current_sub.get(s, 0) for s in scenes]
    yoy_rates = []
    mom_rates = []
    for i, s in enumerate(scenes):
        cur = cur_counts[i]
        yoy_val = yoy_sub.get(s, 0) if yoy_sub else 0
        mom_val = mom_sub.get(s, 0) if mom_sub else 0
        yoy_rates.append(((cur - yoy_val) / yoy_val * 100) if yoy_val > 0 else (100 if cur > 0 else 0))
        mom_rates.append(((cur - mom_val) / mom_val * 100) if mom_val > 0 else (100 if cur > 0 else 0))

    base_labels = [SUB_SCENE_MAP.get(s, s) for s in scenes]
    # 合并当月件数到标签
    new_labels = [f"{base_labels[i]}（{cur_counts[i]}件）" for i in range(len(scenes))]
    x = np.arange(len(scenes))

    fig, ax = plt.subplots(figsize=(12, 6))
    # 绘制折线图
    line1, = ax.plot(x, yoy_rates, marker='o', linestyle='-', linewidth=2, markersize=8,
                     color='#3A8C7A', label='同比增长率 (%)')
    line2, = ax.plot(x, mom_rates, marker='s', linestyle='--', linewidth=2, markersize=8,
                     color='#C03C2B', label='环比增长率 (%)')

    # 标注折点数值
    for i, (yoy, mom) in enumerate(zip(yoy_rates, mom_rates)):
        ax.text(i, yoy + 5.0, f'{yoy:.1f}%', ha='center', va='bottom',
                fontsize=12, color='#3A8C7A')
        ax.text(i, mom + 5.0, f'{mom:.1f}%', ha='center', va='bottom',
                fontsize=12, color='#C03C2B')

    ax.set_xlabel('子场景', fontsize=12)
    ax.set_ylabel('增长率 (%)', fontsize=12)
    ax.set_title('高频事件所属领域同环比变化情况', fontsize=14)
    ax.set_xticks(x)
    # ax.set_xticklabels(new_labels, rotation=15, ha='right')
    ax.set_xticklabels(new_labels, rotation=0, ha='center')
    ax.legend(loc='lower center', bbox_to_anchor=(0.5, 1.05), ncol=2)

    # 根据数据范围自动调整y轴，无需为额外标注留空间
    # all_rates = yoy_rates + mom_rates
    # y_max = max(max(all_rates), 10) + 30
    # y_min = min(min(all_rates), -10) - 15  # 稍微留一点空间即可
    # ax.set_ylim(y_min, y_max)

    all_rates = yoy_rates + mom_rates
    max_rate = max(all_rates)
    min_rate = min(all_rates)
    if max_rate > 600:
        y_max = max(max_rate, 10) + 60
        y_min = min(min_rate, -10) - 30
    else:
        y_max = max(max_rate, 10) + 30
        y_min = min(min_rate, -10) - 15
    ax.set_ylim(y_min, y_max)


    ax.grid(True, linestyle='--', alpha=0.6)
    plt.tight_layout()
    plt.savefig(output_dir / '整体分析_图3_子场景同环比.png', dpi=150)
    plt.close()
    print("已生成：整体分析_图3_子场景同环比.png")



def plot_sub_fig1(sub_scene, biz_dict, output_dir):
    """子场景图1：二级业务占比饼图"""
    if not biz_dict:
        print(f"子场景 {sub_scene} 无二级业务数据，跳过图1")
        return
    sorted_items = sorted(biz_dict.items(), key=lambda x: x[1], reverse=True)
    if len(sorted_items) > 8:
        top = sorted_items[:7]
        other_sum = sum(c for _, c in sorted_items[7:])
        top.append(('其他', other_sum))
        biz_dict = dict(top)
    else:
        biz_dict = dict(sorted_items)

    labels = list(biz_dict.keys())
    sizes = list(biz_dict.values())
    plt.figure(figsize=(10, 8))
    colors = plt.cm.Set3(np.linspace(0, 1, len(labels)))
    plt.pie(sizes, labels=labels, autopct='%1.1f%%', colors=colors, startangle=90,textprops={'fontsize': 12})
    plt.title(f'{SUB_SCENE_MAP.get(sub_scene, sub_scene)}领域高频事件分布情况', fontsize=14)
    plt.tight_layout()
    plt.savefig(output_dir / f'{sub_scene}_图1_高频事件分布.png', dpi=150)
    plt.close()
    print(f"已生成：{sub_scene}_图1_高频事件分布.png")


def plot_sub_change(sub_scene, current_biz, compare_biz, output_dir, label, compare_label):
    """
    子场景同比/环比变化图（三组：当前月柱、对比期柱、增长率折线）
    current_biz: 当前月二级业务字典 {二级业务: 件数}
    compare_biz: 对比期（去年同月或上月）的二级业务字典
    label: '同比' 或 '环比'
    compare_label: 对比期标签，如 '202503'
    """
    all_biz = set(current_biz.keys()) | set(compare_biz.keys())
    sorted_biz = sorted(all_biz, key=lambda x: current_biz.get(x, 0), reverse=True)[:10]
    if not sorted_biz:
        print(f"子场景 {sub_scene} 无有效二级业务数据，跳过{label}图")
        return

    cur_vals = [current_biz.get(b, 0) for b in sorted_biz]
    comp_vals = [compare_biz.get(b, 0) for b in sorted_biz]
    growth = []
    for c, p in zip(cur_vals, comp_vals):
        if p == 0:
            g = 100 if c > 0 else 0
        else:
            g = (c - p) / p * 100
        growth.append(g)
    fig, ax1 = plt.subplots(figsize=(12, 6))
    x = np.arange(len(sorted_biz))
    width = 0.2

    # 当前月柱状图（蓝色）
    bars1 = ax1.bar(x - width/2, cur_vals, width, color='steelblue', label='本月高频诉求量', zorder=2)
    # 对比期柱状图（橙色）
    bars2 = ax1.bar(x + width/2, comp_vals, width, color='orange', label=f'{compare_label}高频诉求量', zorder=2)

    ax1.set_xlabel('二级业务', fontsize=12)
    ax1.set_ylabel('事件数（件）', fontsize=12, color='black')
    ax1.set_xticks(x)
    # ax1.set_xticklabels(sorted_biz, rotation=30, ha='right', fontsize=10)
    ax1.set_xticklabels(sorted_biz, rotation=0, ha='center', fontsize=10)

    # 添加数值标签
    for bar in bars1:
        h = bar.get_height()
        if h > 0:
            ax1.text(bar.get_x() + bar.get_width()/2., h + 0.3, f'{int(h)}', ha='center', va='bottom', fontsize=10, color='steelblue')
    for bar in bars2:
        h = bar.get_height()
        if h > 0:
            ax1.text(bar.get_x() + bar.get_width()/2., h + 0.3, f'{int(h)}', ha='center', va='bottom', fontsize=10, color='orange')

    # 折线图（增长率）
    ax2 = ax1.twinx()
    line = ax2.plot(x, growth, 'ro-', linewidth=2, markersize=6, label=f'{label}增长率', zorder=3)
    ax2.set_ylabel('增长率 (%)', fontsize=12, color='red')
    ax2.tick_params(axis='y', labelcolor='red')
    for i, g in enumerate(growth):
        ax2.text(i, g + 3, f'{g:.1f}%', ha='center', va='bottom', fontsize=10, color='red')

    # 合并图例
    lines1, labels1 = ax1.get_legend_handles_labels()
    lines2, labels2 = ax2.get_legend_handles_labels()
    # ax1.legend(lines1 + lines2, labels1 + labels2, loc='upper left')
    ax1.legend(lines1 + lines2, labels1 + labels2,
               loc='lower center', bbox_to_anchor=(0.5, 1.12), ncol=3)

    # 增加上部留白：根据当前数据最大值动态设置纵轴上限
    max_cur = max(cur_vals) if cur_vals else 0
    max_comp = max(comp_vals) if comp_vals else 0
    max_val = max(max_cur, max_comp)
    if max_val > 0:
        ax1.set_ylim(top=max_val + 10)

    # 增加折线图上留白（右轴）
    positive_growth = [g for g in growth if g > 0]
    max_growth = max(positive_growth) if positive_growth else 0
    if max_growth > 0:
        if max_growth >450:
            ax2.set_ylim(top=max_growth + 50)
        else:
        # 可以根据需要选择固定增量或比例，例如固定增加 5 个百分点
            ax2.set_ylim(top=max_growth + 20)
    else:
        # 如果没有正增长率，也设置一个合理的上界（如 10%）
        ax2.set_ylim(top=10)

    min_growth = min(growth) if growth else 0
    ax2.set_ylim(bottom=min_growth - 15)      #下方留白


    title_name = SUB_SCENE_MAP.get(sub_scene, sub_scene)
    change_type = '同比' if label == '同比' else '环比'
    plt.title(f'{title_name}领域高频事件{change_type}变化情况', fontsize=14)
    plt.tight_layout()
    suffix = '同比' if label == '同比' else '环比'
    plt.savefig(output_dir / f'{sub_scene}_图{3 if label == "环比" else 2}_{suffix}.png', dpi=150)
    plt.close()
    print(f"已生成：{sub_scene}_图{3 if label == '环比' else 2}_{suffix}.png")


def plot_sub_missing(sub_scene, output_dir, label):
    """生成缺失数据的提示图"""
    title_name = SUB_SCENE_MAP.get(sub_scene, sub_scene)
    fig, ax = plt.subplots(figsize=(10, 6))
    ax.text(0.5, 0.5, f'暂无{label}历史数据\n无法生成{label}变化图',
            transform=ax.transAxes, ha='center', va='center', fontsize=14, color='gray')
    ax.set_title(f'{title_name}领域高频事件{label}变化情况（数据缺失）', fontsize=14)
    plt.tight_layout()
    suffix = "同比" if label == "同比" else "环比"
    plt.savefig(output_dir / f'{sub_scene}_图{3 if label == "环比" else 2}_{suffix}.png', dpi=150)
    plt.close()
    print(f"已生成：{sub_scene}_图{3 if label == '环比' else 2}_{suffix}.png（数据缺失）")


def run_analysis(current_path, yoy_path=None, mom_path=None, output_dir="output_charts"):
    """
    参数：
        current_path: 当前月 Excel 文件路径（字符串或 Path 对象）
        yoy_path:     同比（去年同月）Excel 文件路径，可选
        mom_path:     环比（上月）Excel 文件路径，可选
        output_dir:   输出图片的目录（字符串或 Path 对象）
    """
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    # 加载当前月数据
    cur_sub, cur_biz, cur_event, cur_sub_biz = load_summary_data(current_path)
    if cur_sub is None:
        raise ValueError("当前月数据加载失败，请检查文件格式")

    # 提取月份标签
    current_month = extract_month_from_path(current_path)

    # 同比数据
    yoy_sub, yoy_biz, yoy_event, yoy_sub_biz = None, None, None, None
    yoy_label = None
    if yoy_path is not None and Path(yoy_path).exists():
        yoy_sub, yoy_biz, yoy_event, yoy_sub_biz = load_summary_data(yoy_path)
        if yoy_sub is not None:
            yoy_label = extract_month_from_path(yoy_path)
            print(f"同比数据加载成功，对比月份：{yoy_label}")
        else:
            print("警告：同比文件无效，将忽略同比图")

    # 环比数据
    mom_sub, mom_biz, mom_event, mom_sub_biz = None, None, None, None
    mom_label = None
    if mom_path is not None and Path(mom_path).exists():
        mom_sub, mom_biz, mom_event, mom_sub_biz = load_summary_data(mom_path)
        if mom_sub is not None:
            mom_label = extract_month_from_path(mom_path)
            print(f"环比数据加载成功，对比月份：{mom_label}")
        else:
            print("警告：环比文件无效，将忽略环比图")

    # 整体分析图
    plot_overall_fig1(cur_biz, output_dir)
    plot_overall_fig2(cur_event, output_dir)
    plot_overall_fig3(cur_sub, yoy_sub if yoy_sub else {}, mom_sub if mom_sub else {}, output_dir)

    # 子场景图表
    for scene in SUB_SCENE_NAMES:
        cur_dict = cur_sub_biz.get(scene, {})
        plot_sub_fig1(scene, cur_dict, output_dir)

        # 同比
        if yoy_sub_biz is not None and scene in yoy_sub_biz:
            plot_sub_change(scene, cur_dict, yoy_sub_biz.get(scene, {}), output_dir, '同比', yoy_label or '去年同月')
        else:
            plot_sub_missing(scene, output_dir, '同比')

        # 环比
        if mom_sub_biz is not None and scene in mom_sub_biz:
            plot_sub_change(scene, cur_dict, mom_sub_biz.get(scene, {}), output_dir, '环比', mom_label or '上月')
        else:
            plot_sub_missing(scene, output_dir, '环比')

    print(f"所有图表已保存至：{output_dir.absolute()}")