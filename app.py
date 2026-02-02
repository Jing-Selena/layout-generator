#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
布局图生成器 Web应用
使用Streamlit创建的用户界面
"""

import streamlit as st
import os
import tempfile
from layout_generator import LayoutGenerator

# 设置页面配置（中文字体由 layout_generator 内的 _setup_chinese_font 统一设置）
st.set_page_config(
    page_title="布局图生成器",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

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
    工具会生成以下布局图及对应的 Excel 表（每图一个 Excel，可单独下载图片或 Excel）：
    - 项目中类 / 项目小类 / 项目细类 / 销售类别 / 品牌 布局图 + 合并明细表
    Excel 列：template_name、shelf_id、layer_id、pos_id、value、dimension_name
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

# 生成逻辑：成功后将结果写入 session_state，之后无论点多少次下载都不会回弹
if generate_button and product_file is not None and layout_file is not None:
    with st.spinner("正在处理数据，请稍候..."):
        try:
            with tempfile.TemporaryDirectory() as temp_dir:
                product_path = os.path.join(temp_dir, product_file.name)
                layout_path = os.path.join(temp_dir, layout_file.name)
                with open(product_path, "wb") as f:
                    f.write(product_file.getbuffer())
                with open(layout_path, "wb") as f:
                    f.write(layout_file.getbuffer())

                generator = LayoutGenerator(temp_dir)
                generator.load_data(product_path, layout_path)
                generator.match_data()
                template_name = generator.get_shelf_template_name()
                shelf_info, shelf_col, layer_col, position_col = generator.get_shelf_info()

                dimensions = [
                    ("项目中类", "项目中类"),
                    ("项目小类", "项目小类"),
                    ("项目细类", "项目细类"),
                    ("项目商品类别", "销售类别"),
                    ("品牌名称", "品牌")
                ]
                results = []
                download_png_name = f"{template_name}_布局图.png" if template_name else "布局图.png"
                download_xlsx_name = f"{template_name}_布局图.xlsx" if template_name else "布局图.xlsx"
                for field_name, display_name in dimensions:
                    type_code = LayoutGenerator.DIMENSION_TYPE_MAP.get(display_name, display_name)
                    if template_name:
                        output_filename = f"{template_name}_布局图_{type_code}.png"
                    else:
                        output_filename = f"布局图_{type_code}.png"
                    output_path = os.path.join(temp_dir, output_filename)
                    generator.generate_product_layout(
                        shelf_info, shelf_col, layer_col, position_col,
                        field_name, display_name, template_name, output_path
                    )
                    with open(output_path, "rb") as f:
                        png_bytes = f.read()
                    xlsx_path = output_path.rsplit(".", 1)[0] + ".xlsx"
                    xlsx_bytes, xlsx_name = None, None
                    if os.path.exists(xlsx_path):
                        with open(xlsx_path, "rb") as f:
                            xlsx_bytes = f.read()
                        xlsx_name = download_xlsx_name
                    results.append((png_bytes, download_png_name, xlsx_bytes, xlsx_name))

                st.session_state["layout_results"] = results
                st.session_state["layout_meta"] = {
                    "template_name": template_name,
                    "shelf_info": shelf_info,
                    "shelf_count": len(shelf_info),
                }
        except Exception as e:
            st.error(f"❌ 生成布局图时出错: {str(e)}")
            import traceback
            st.code(traceback.format_exc())
elif generate_button and (product_file is None or layout_file is None):
    st.error("❌ 请上传两个Excel文件后再生成布局图！")

# 只要有结果就始终展示结果区（下载后不回到顶部、可继续看下一个图并下载）
if st.session_state.get("layout_results"):
    meta = st.session_state.get("layout_meta") or {}
    st.success("✅ 布局图生成成功！")
    st.info(f"📊 找到 {meta.get('shelf_count', 0)} 个货架")
    if meta.get("template_name"):
        st.info(f"🏷️ 货架模板名称: {meta['template_name']}")
    for shelf, layers in (meta.get("shelf_info") or {}).items():
        st.text(f"  货架 {shelf}: {len(layers)} 层")

    st.markdown("---")
    st.header("🖼️ 布局图预览")
    st.caption("可切换标签查看不同维度图，点击下载后不会跳回上方，可继续查看或下载其他图。")

    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "📊 项目中类",
        "📊 项目小类",
        "📊 项目细类",
        "📊 销售类别",
        "📊 品牌"
    ])
    tab_list = [tab1, tab2, tab3, tab4, tab5]
    results = st.session_state["layout_results"]

    for idx, tab in enumerate(tab_list):
        if idx >= len(results):
            break
        png_bytes, png_name, xlsx_bytes, xlsx_name = results[idx]
        with tab:
            st.subheader(f"📊 {png_name.replace('.png', '')}")
            st.image(png_bytes, use_container_width=True)
            c1, c2 = st.columns(2)
            with c1:
                st.download_button(
                    label=f"📥 下载图片 {png_name}",
                    data=png_bytes,
                    file_name=png_name,
                    mime="image/png",
                    key=f"dl_img_{idx}_{png_name}",
                    use_container_width=True
                )
            with c2:
                if xlsx_bytes is not None and xlsx_name:
                    st.download_button(
                        label=f"📥 下载Excel {xlsx_name}",
                        data=xlsx_bytes,
                        file_name=xlsx_name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"dl_xlsx_{idx}_{xlsx_name}",
                        use_container_width=True
                    )
                else:
                    st.caption("本图无对应 Excel")

    st.markdown("---")
    if st.button("🔄 清空结果，重新生成", key="clear_results"):
        for k in ("layout_results", "layout_meta"):
            if k in st.session_state:
                del st.session_state[k]
        st.rerun()

# 页脚
st.markdown("---")
st.markdown(
    "<div style='text-align: center; color: gray;'>布局图生成器 v1.0</div>",
    unsafe_allow_html=True
)
