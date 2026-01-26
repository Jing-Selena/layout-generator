#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
布局图生成器 Web应用
使用Streamlit创建的用户界面
"""

import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.patches as patches
import os
import tempfile
import zipfile
from io import BytesIO
from layout_generator import LayoutGenerator

# 设置页面配置
st.set_page_config(
    page_title="布局图生成器",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 设置中文字体
plt.rcParams['font.sans-serif'] = ['Arial Unicode MS', 'SimHei', 'STHeiti']
plt.rcParams['axes.unicode_minus'] = False

# 标题
st.title("📊 布局图生成器")
st.markdown("---")

# 侧边栏说明
with st.sidebar:
    st.header("📖 使用说明")
    st.markdown("""
    ### 功能说明
    本工具可以根据商品资料表和落位明细清单自动生成货架布局图。
    
    ### 文件要求
    1. **商品资料表**（Excel文件）
       - 文件名需包含"商品资料表"
       - 必须包含列：`商品编码`
       - 应包含列：`项目商品类别`、`项目中类`、`项目小类`、`项目细类`、`品牌名称`
    
    2. **落位明细清单**（Excel文件）
       - 文件名需包含"落位明细清单"
       - 必须包含列：`*商品编码`（或`商品编码`）
       - 必须包含列：`*货架序号`（或`货架序号`）
       - 必须包含列：`*层数`（或`层数`）
       - 必须包含列：`*位置`（或`位置`）
    
    ### 生成内容
    工具会生成以下布局图：
    - 货架框架图
    - 项目中类布局图
    - 项目小类布局图
    - 项目细类布局图
    - 销售类别布局图
    - 品牌布局图
    """)

# 文件上传区域
st.header("📁 文件上传")

col1, col2 = st.columns(2)

with col1:
    st.subheader("商品资料表")
    product_file = st.file_uploader(
        "请上传商品资料表（Excel文件）",
        type=['xlsx', 'xls'],
        key="product_file"
    )

with col2:
    st.subheader("落位明细清单")
    layout_file = st.file_uploader(
        "请上传落位明细清单（Excel文件）",
        type=['xlsx', 'xls'],
        key="layout_file"
    )

# 生成按钮
st.markdown("---")
generate_button = st.button("🚀 生成布局图", type="primary", use_container_width=True)

# 处理生成逻辑
if generate_button:
    if product_file is None or layout_file is None:
        st.error("❌ 请上传两个Excel文件后再生成布局图！")
    else:
        with st.spinner("正在处理数据，请稍候..."):
            try:
                # 创建临时目录
                with tempfile.TemporaryDirectory() as temp_dir:
                    # 保存上传的文件
                    product_path = os.path.join(temp_dir, product_file.name)
                    layout_path = os.path.join(temp_dir, layout_file.name)
                    
                    with open(product_path, "wb") as f:
                        f.write(product_file.getbuffer())
                    
                    with open(layout_path, "wb") as f:
                        f.write(layout_file.getbuffer())
                    
                    # 创建生成器实例
                    generator = LayoutGenerator(temp_dir)
                    
                    # 加载数据
                    generator.load_data(product_path, layout_path)
                    
                    # 匹配数据
                    generator.match_data()
                    
                    # 获取货架模板名称
                    template_name = generator.get_shelf_template_name()
                    
                    # 获取货架信息
                    shelf_info, shelf_col, layer_col, position_col = generator.get_shelf_info()
                    
                    # 生成布局图
                    generated_files = []
                    
                    # 生成货架框架图
                    if template_name:
                        framework_filename = f"{template_name}-货架框架图.png"
                    else:
                        framework_filename = "货架框架图.png"
                    framework_path = os.path.join(temp_dir, framework_filename)
                    generator.generate_shelf_framework(shelf_info, framework_path)
                    generated_files.append((framework_path, framework_filename))
                    
                    # 生成五个维度的商品布局图
                    dimensions = [
                        ("项目中类", "项目中类"),
                        ("项目小类", "项目小类"),
                        ("项目细类", "项目细类"),
                        ("项目商品类别", "销售类别"),
                        ("品牌名称", "品牌")
                    ]
                    
                    for field_name, display_name in dimensions:
                        if template_name:
                            output_filename = f"{template_name}-{display_name}布局图.png"
                        else:
                            output_filename = f"{display_name}布局图.png"
                        output_path = os.path.join(temp_dir, output_filename)
                        generator.generate_product_layout(
                            shelf_info, shelf_col, layer_col, position_col,
                            field_name, display_name, template_name, output_path
                        )
                        generated_files.append((output_path, output_filename))
                    
                    # 创建ZIP文件
                    zip_buffer = BytesIO()
                    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                        for file_path, filename in generated_files:
                            if os.path.exists(file_path):
                                zip_file.write(file_path, filename)
                    
                    zip_buffer.seek(0)
                    
                    # 显示成功信息
                    st.success("✅ 布局图生成成功！")
                    
                    # 显示货架信息
                    st.info(f"📊 找到 {len(shelf_info)} 个货架")
                    for shelf, layers in shelf_info.items():
                        st.text(f"  货架 {shelf}: {len(layers)} 层")
                    
                    if template_name:
                        st.info(f"🏷️ 货架模板名称: {template_name}")
                    
                    # 在页面上展示所有布局图
                    st.markdown("---")
                    st.header("🖼️ 布局图预览")
                    
                    # 使用标签页展示不同类型的图
                    tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
                        "📐 货架框架图",
                        "📊 项目中类",
                        "📊 项目小类",
                        "📊 项目细类",
                        "📊 销售类别",
                        "📊 品牌"
                    ])
                    
                    # 定义每个标签页对应的文件索引（货架框架图是第0个，然后是5个维度图）
                    tab_configs = [
                        (0, tab1, "货架框架图"),  # 货架框架图
                        (1, tab2, "项目中类"),    # 项目中类
                        (2, tab3, "项目小类"),    # 项目小类
                        (3, tab4, "项目细类"),    # 项目细类
                        (4, tab5, "销售类别"),    # 销售类别
                        (5, tab6, "品牌"),        # 品牌
                    ]
                    
                    # 显示每个标签页的图片
                    for idx, (file_idx, tab, tab_name) in enumerate(tab_configs):
                        if file_idx < len(generated_files):
                            file_path, filename = generated_files[file_idx]
                            if os.path.exists(file_path):
                                with tab:
                                    st.subheader(f"📊 {filename.replace('.png', '')}")
                                    with open(file_path, "rb") as f:
                                        image_data = f.read()
                                    st.image(image_data, use_container_width=True)
                                    
                                    # 在每个图片下方添加下载按钮
                                    col1, col2, col3 = st.columns([1, 1, 1])
                                    with col2:
                                        st.download_button(
                                            label=f"📥 下载 {filename}",
                                            data=image_data,
                                            file_name=filename,
                                            mime="image/png",
                                            key=f"download_{idx}_{filename}",
                                            use_container_width=True
                                        )
                    
                    # 下载按钮区域
                    st.markdown("---")
                    st.header("📥 批量下载")
                    
                    # 下载ZIP文件
                    st.download_button(
                        label="📦 下载所有布局图（ZIP压缩包）",
                        data=zip_buffer,
                        file_name=f"{template_name if template_name else '布局图'}-所有图纸.zip",
                        mime="application/zip",
                        use_container_width=True
                    )
                    
            except Exception as e:
                st.error(f"❌ 生成布局图时出错: {str(e)}")
                import traceback
                st.code(traceback.format_exc())

# 页脚
st.markdown("---")
st.markdown(
    "<div style='text-align: center; color: gray;'>布局图生成器 v1.0</div>",
    unsafe_allow_html=True
)
