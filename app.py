import streamlit as st
import tempfile
import zipfile
import os
from pathlib import Path
from chart_generator import run_analysis  # 导入你刚才写的函数

st.set_page_config(page_title="多人一诉可视化工具", layout="wide")
st.title("📊 多人一诉数据分析工具")
st.markdown("上传三个 Excel 文件（当前月、同比、环比），自动生成所有图表并打包下载。")

# 文件上传区域
col1, col2, col3 = st.columns(3)
with col1:
    current_file = st.file_uploader("📁 当前月数据 (必填)", type=["xlsx"], key="current")
with col2:
    yoy_file = st.file_uploader("📁 同比数据 (去年同月, 可选)", type=["xlsx"], key="yoy")
with col3:
    mom_file = st.file_uploader("📁 环比数据 (上月, 可选)", type=["xlsx"], key="mom")

if st.button("🚀 生成图表", type="primary"):
    if current_file is None:
        st.error("请至少上传当前月数据文件")
        st.stop()

    # 用临时目录存放上传的文件和生成的图片
    with tempfile.TemporaryDirectory() as tmpdir:
        tmp_path = Path(tmpdir)

        # 保存当前文件
        current_path = tmp_path / "current.xlsx"
        with open(current_path, "wb") as f:
            f.write(current_file.getbuffer())

        # 保存同比文件（如果有）
        yoy_path = None
        if yoy_file is not None:
            yoy_path = tmp_path / "yoy.xlsx"
            with open(yoy_path, "wb") as f:
                f.write(yoy_file.getbuffer())

        # 保存环比文件（如果有）
        mom_path = None
        if mom_file is not None:
            mom_path = tmp_path / "mom.xlsx"
            with open(mom_path, "wb") as f:
                f.write(mom_file.getbuffer())

        # 输出目录
        img_dir = tmp_path / "charts"
        img_dir.mkdir()

        # 调用绘图函数
        with st.spinner("正在生成图表，请稍候..."):
            run_analysis(
                current_path=current_path,
                yoy_path=yoy_path,
                mom_path=mom_path,
                output_dir=img_dir
            )

        # 收集所有图片
        image_files = list(img_dir.glob("*.png"))
        if not image_files:
            st.warning("没有生成任何图片，请检查数据格式")
            st.stop()

        st.success(f"✅ 成功生成 {len(image_files)} 张图表")

        # 预览前几张
        with st.expander("📸 点击预览前 6 张图"):
            for img_path in image_files[:6]:
                st.image(str(img_path), caption=img_path.name, use_container_width=True)

        # 打包为 zip 并提供下载按钮
        zip_path = tmp_path / "charts.zip"
        with zipfile.ZipFile(zip_path, 'w') as zipf:
            for img in image_files:
                zipf.write(img, arcname=img.name)

        with open(zip_path, "rb") as f:
            st.download_button(
                label="📦 下载全部图表 (ZIP)",
                data=f,
                file_name="multi_petition_charts.zip",
                mime="application/zip"
            )