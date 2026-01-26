#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
布局图生成器
根据商品资料表和落位明细清单生成货架布局图
"""

import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.patches as patches
from matplotlib import font_manager
import os
import glob
from typing import Dict, List, Tuple
import numpy as np

# 设置中文字体
plt.rcParams['font.sans-serif'] = ['Arial Unicode MS', 'SimHei', 'STHeiti']
plt.rcParams['axes.unicode_minus'] = False


class LayoutGenerator:
    """布局图生成器类"""
    
    def __init__(self, folder_path: str = None):
        """
        初始化布局图生成器
        
        Args:
            folder_path: 文件夹路径，如果为None则使用当前目录
        """
        self.folder_path = folder_path or os.getcwd()
        self.product_df = None
        self.layout_df = None
        self.merged_df = None
        
    def find_excel_files(self) -> Tuple[str, str]:
        """
        查找商品资料表和落位明细清单文件
        
        Returns:
            (product_file, layout_file): 商品资料表文件路径和落位明细清单文件路径
        """
        files = glob.glob(os.path.join(self.folder_path, "*.xlsx"))
        
        product_file = None
        layout_file = None
        
        for file in files:
            filename = os.path.basename(file)
            if "商品资料表" in filename:
                product_file = file
            elif "落位明细清单" in filename:
                layout_file = file
        
        if product_file is None:
            raise FileNotFoundError("未找到包含'商品资料表'的Excel文件")
        if layout_file is None:
            raise FileNotFoundError("未找到包含'落位明细清单'的Excel文件")
        
        return product_file, layout_file
    
    def load_data(self, product_file: str = None, layout_file: str = None):
        """
        加载Excel数据
        
        Args:
            product_file: 商品资料表文件路径
            layout_file: 落位明细清单文件路径
        """
        if product_file is None or layout_file is None:
            product_file, layout_file = self.find_excel_files()
        
        print(f"正在读取商品资料表: {product_file}")
        self.product_df = pd.read_excel(product_file)
        print(f"商品资料表列名: {self.product_df.columns.tolist()}")
        print(f"商品资料表行数: {len(self.product_df)}")
        
        print(f"正在读取落位明细清单: {layout_file}")
        self.layout_df = pd.read_excel(layout_file)
        print(f"落位明细清单列名: {self.layout_df.columns.tolist()}")
        print(f"落位明细清单行数: {len(self.layout_df)}")
        
    def match_data(self):
        """
        根据商品编码匹配数据
        """
        if self.product_df is None or self.layout_df is None:
            raise ValueError("请先加载数据")
        
        # 查找商品编码列（可能的列名）
        product_code_col = None
        for col in self.product_df.columns:
            if "商品编码" in str(col):
                product_code_col = col
                break
        
        layout_code_col = None
        # 优先匹配带*的商品编码列
        for col in self.layout_df.columns:
            col_str = str(col)
            if "*商品编码" in col_str or (col_str.startswith("*") and "商品编码" in col_str):
                layout_code_col = col
                break
        # 如果没有找到，再找普通的商品编码列
        if layout_code_col is None:
            for col in self.layout_df.columns:
                if "商品编码" in str(col):
                    layout_code_col = col
                    break
        
        if product_code_col is None:
            raise ValueError("商品资料表中未找到商品编码列")
        if layout_code_col is None:
            raise ValueError("落位明细清单中未找到商品编码列")
        
        print(f"商品资料表编码列: {product_code_col}")
        print(f"落位明细清单编码列: {layout_code_col}")
        
        # 查找需要的字段
        required_fields = ["项目商品类别", "项目小类", "项目细类", "品牌名称"]
        available_fields = []
        for field in required_fields:
            for col in self.product_df.columns:
                if field in str(col):
                    available_fields.append((field, col))
                    break
        
        print(f"找到的字段映射: {available_fields}")
        
        # 合并数据
        merge_on_left = layout_code_col
        merge_on_right = product_code_col
        
        self.merged_df = self.layout_df.merge(
            self.product_df,
            left_on=merge_on_left,
            right_on=merge_on_right,
            how='left',
            suffixes=('', '_product')
        )
        
        print(f"合并后数据行数: {len(self.merged_df)}")
        print(f"合并后列名: {self.merged_df.columns.tolist()}")
    
    def get_shelf_template_name(self) -> str:
        """
        获取货架模板名称
        
        Returns:
            货架模板名称字符串
        """
        if self.merged_df is None:
            raise ValueError("请先匹配数据")
        
        # 查找货架模板名称列
        template_name_col = None
        for col in self.merged_df.columns:
            if "货架模板名称" in str(col):
                template_name_col = col
                break
        
        if template_name_col is None:
            return ""
        
        # 获取第一个非空值
        template_name = ""
        for _, row in self.merged_df.iterrows():
            if pd.notna(row[template_name_col]):
                template_name = str(row[template_name_col]).strip()
                break
        
        return template_name
        
    def get_shelf_info(self) -> Dict:
        """
        获取货架信息
        
        Returns:
            货架信息字典，格式: {货架序号: {层数: [位置列表]}}
        """
        if self.merged_df is None:
            raise ValueError("请先匹配数据")
        
        # 查找货架序号、层数、位置列（优先匹配带*的列）
        shelf_col = None
        layer_col = None
        position_col = None
        
        # 先找带*的列
        for col in self.merged_df.columns:
            col_str = str(col)
            if "*货架序号" in col_str or (col_str.startswith("*") and "货架序号" in col_str):
                shelf_col = col
            elif "*层数" in col_str or (col_str.startswith("*") and "层数" in col_str and "组件" not in col_str):
                layer_col = col
            elif "*位置" in col_str or (col_str.startswith("*") and "位置" in col_str and "垫高" not in col_str):
                position_col = col
        
        # 如果没找到带*的，再找普通的
        if shelf_col is None:
            for col in self.merged_df.columns:
                col_str = str(col)
                if "货架序号" in col_str:
                    shelf_col = col
                    break
        if layer_col is None:
            for col in self.merged_df.columns:
                col_str = str(col)
                if "层数" in col_str and "组件" not in col_str:
                    layer_col = col
                    break
        if position_col is None:
            for col in self.merged_df.columns:
                col_str = str(col)
                if "位置" in col_str and "垫高" not in col_str and "模板" not in col_str:
                    position_col = col
                    break
        
        if shelf_col is None or layer_col is None or position_col is None:
            raise ValueError("未找到货架序号、层数或位置列")
        
        print(f"货架序号列: {shelf_col}")
        print(f"层数列: {layer_col}")
        print(f"位置列: {position_col}")
        
        shelf_info = {}
        for _, row in self.merged_df.iterrows():
            shelf = row[shelf_col]
            layer = row[layer_col]
            position = row[position_col]
            
            if pd.isna(shelf) or pd.isna(layer) or pd.isna(position):
                continue
            
            shelf = str(shelf).strip()
            layer = int(layer) if pd.notna(layer) else 0
            position = str(position).strip()
            
            if shelf not in shelf_info:
                shelf_info[shelf] = {}
            if layer not in shelf_info[shelf]:
                shelf_info[shelf][layer] = []
            if position not in shelf_info[shelf][layer]:
                shelf_info[shelf][layer].append(position)
        
        # 对位置进行排序
        for shelf in shelf_info:
            for layer in shelf_info[shelf]:
                try:
                    shelf_info[shelf][layer] = sorted(
                        shelf_info[shelf][layer],
                        key=lambda x: int(x) if str(x).isdigit() else float('inf')
                    )
                except:
                    shelf_info[shelf][layer] = sorted(shelf_info[shelf][layer])
        
        return shelf_info, shelf_col, layer_col, position_col
    
    def generate_shelf_framework(self, shelf_info: Dict, output_path: str = "货架框架图.png"):
        """
        生成货架框架图
        
        Args:
            shelf_info: 货架信息字典
            output_path: 输出文件路径
        """
        num_shelves = len(shelf_info)
        if num_shelves == 0:
            raise ValueError("未找到任何货架信息，请检查数据")
        
        if num_shelves == 1:
            fig, ax = plt.subplots(1, 1, figsize=(5, 8))
            axes = [ax]
        else:
            fig, axes = plt.subplots(1, num_shelves, figsize=(5 * num_shelves, 8))
        
        for idx, (shelf_num, layers) in enumerate(sorted(shelf_info.items())):
            ax = axes[idx]
            max_layers = max(layers.keys()) if layers else 1
            max_positions = max(len(positions) for positions in layers.values()) if layers else 1
            
            # 绘制货架框架
            for layer in range(1, max_layers + 1):
                if layer in layers:
                    positions = layers[layer]
                    num_positions = len(positions)
                    
                    # 绘制层
                    y_bottom = max_layers - layer
                    y_top = y_bottom + 0.8
                    
                    # 绘制分隔线
                    for i in range(num_positions + 1):
                        x = i * (1.0 / (num_positions + 1))
                        ax.plot([x, x], [y_bottom, y_top], 'k-', linewidth=1)
                    
                    # 绘制上下边框
                    ax.plot([0, 1], [y_bottom, y_bottom], 'k-', linewidth=2)
                    ax.plot([0, 1], [y_top, y_top], 'k-', linewidth=2)
                    
                    # 标注位置
                    for i, pos in enumerate(positions):
                        x_center = (i + 0.5) * (1.0 / (num_positions + 1))
                        ax.text(x_center, y_bottom + 0.4, str(pos), 
                               ha='center', va='center', fontsize=8)
            
            ax.set_xlim(-0.1, 1.1)
            ax.set_ylim(-0.1, max_layers + 0.1)
            ax.set_title(f'货架 {shelf_num}', fontsize=12, fontweight='bold')
            ax.set_xlabel('位置', fontsize=10)
            ax.set_ylabel('层数', fontsize=10)
            ax.set_aspect('equal')
            ax.grid(True, alpha=0.3)
            ax.invert_yaxis()
        
        plt.tight_layout()
        plt.savefig(output_path, dpi=300, bbox_inches='tight')
        print(f"货架框架图已保存至: {output_path}")
        plt.close()
    
    def generate_product_layout(self, shelf_info: Dict, shelf_col: str, 
                                layer_col: str, position_col: str,
                                dimension_field: str, dimension_name: str,
                                template_name: str = "",
                                output_path: str = "商品布局图.png"):
        """
        生成商品布局图（按维度合并商品，参考格式）
        
        Args:
            shelf_info: 货架信息字典
            shelf_col: 货架序号列名
            layer_col: 层数列名
            position_col: 位置列名
            dimension_field: 维度字段名称（如"项目商品类别"、"项目小类"、"项目细类"）
            dimension_name: 维度显示名称（用于标题和文件名）
            output_path: 输出文件路径
        """
        if self.merged_df is None:
            raise ValueError("请先匹配数据")
        
        # 查找维度字段
        category_col = None
        for col in self.merged_df.columns:
            if dimension_field in str(col):
                category_col = col
                break
        
        if category_col is None:
            print(f"警告: 未找到维度字段 '{dimension_field}'，跳过生成 {dimension_name} 布局图")
            return
        
        print(f"生成 {dimension_name} 布局图，使用的分类字段: {category_col}")
        
        # 为每个类别分配颜色
        category_color_map = {}
        # 使用柔和的颜色调色板
        color_palette = [
            '#E3F2FD',  # 浅蓝色
            '#C8E6C9',  # 浅绿色
            '#FFF9C4',  # 浅黄色
            '#FFE0B2',  # 浅橙色
            '#F8BBD0',  # 浅粉色
            '#B2EBF2',  # 浅青色
            '#D1C4E9',  # 浅紫色
            '#FFCCBC',  # 浅红橙色
            '#C5E1A5',  # 浅黄绿色
            '#BBDEFB',  # 浅天蓝色
            '#F0F4C3',  # 浅黄绿色
            '#FFCDD2',  # 浅红色
            '#E1BEE7',  # 浅紫色
            '#B2DFDB',  # 浅青绿色
            '#FFECB3',  # 浅琥珀色
        ]
        
        # 为每个货架分配背景颜色（更柔和的色调）
        shelf_bg_colors = [
            '#E8F4F8',  # 浅蓝色（更柔和）
            '#FFFACD',  # 浅黄色（更柔和）
            '#FFE4B5',  # 浅橙色（更柔和）
            '#FFE4E1',  # 浅粉色（更柔和）
            '#E0F7FA',  # 浅青色（更柔和）
            '#F3E5F5',  # 浅紫色（更柔和）
        ]
        
        num_shelves = len(shelf_info)
        if num_shelves == 0:
            raise ValueError("未找到任何货架信息，请检查数据")
        
        # 计算每个货架的最大层数
        max_layers_all = max(max(layers.keys()) if layers else 1 
                            for layers in shelf_info.values())
        
        # 创建单个图形，所有货架并排显示
        fig_width = num_shelves * 2.5
        fig_height = max_layers_all * 1.2 + 1
        fig, ax = plt.subplots(1, 1, figsize=(fig_width, fig_height))
        ax.set_xlim(0, num_shelves)
        ax.set_ylim(0, max_layers_all)
        ax.axis('off')
        
        # 存储所有类别用于图例
        all_categories = set()
        
        # 为每个货架绘制
        for idx, (shelf_num, layers) in enumerate(sorted(shelf_info.items())):
            shelf_x_center = idx + 0.5
            shelf_width = 0.9
            shelf_x_left = idx + 0.05
            shelf_x_right = idx + 0.95
            
            # 绘制货架背景
            bg_color = shelf_bg_colors[idx % len(shelf_bg_colors)]
            bg_rect = patches.Rectangle(
                (shelf_x_left, 0), shelf_width, max_layers_all,
                linewidth=2, edgecolor='#666666', facecolor=bg_color,
                alpha=0.3, zorder=0
            )
            ax.add_patch(bg_rect)
            
            # 收集该货架的所有商品数据
            shelf_products = []
            for _, row in self.merged_df.iterrows():
                if str(row[shelf_col]).strip() != str(shelf_num).strip():
                    continue
                
                layer = int(row[layer_col]) if pd.notna(row[layer_col]) else 0
                position = str(row[position_col]).strip()
                
                if layer not in layers or position not in layers[layer]:
                    continue
                
                category = ""
                if category_col and pd.notna(row[category_col]):
                    category = str(row[category_col]).strip()
                    all_categories.add(category)
                
                shelf_products.append({
                    'layer': layer,
                    'position': position,
                    'category': category
                })
            
            # 按层数和位置排序
            shelf_products.sort(key=lambda x: (x['layer'], x['position']))
            
            # 按层分组绘制
            current_y = max_layers_all - 0.1
            
            for layer in sorted(layers.keys()):
                layer_products = [p for p in shelf_products if p['layer'] == layer]
                if not layer_products:
                    current_y -= 1.0
                    continue
                
                positions = sorted(layers[layer], key=lambda x: int(x) if str(x).isdigit() else float('inf'))
                num_positions = len(positions)
                block_height = 0.9
                block_width = shelf_width / num_positions
                
                # 创建位置到商品的映射
                pos_to_product = {}
                for product in layer_products:
                    pos_to_product[product['position']] = product
                
                # 合并相邻相同类别的块
                merged_blocks = []
                i = 0
                while i < len(positions):
                    pos = positions[i]
                    product = pos_to_product.get(pos)
                    if not product:
                        i += 1
                        continue
                    
                    category = product['category']
                    start_idx = i
                    
                    # 查找连续相同类别的块
                    while i < len(positions) and pos_to_product.get(positions[i], {}).get('category') == category:
                        i += 1
                    
                    end_idx = i
                    merged_blocks.append({
                        'start_idx': start_idx,
                        'end_idx': end_idx,
                        'category': category,
                        'positions': positions[start_idx:end_idx]
                    })
                
                # 绘制合并后的块
                for block_info in merged_blocks:
                    start_idx = block_info['start_idx']
                    end_idx = block_info['end_idx']
                    category = block_info['category']
                    
                    # 分配颜色
                    if category and category not in category_color_map:
                        color_idx = len(category_color_map) % len(color_palette)
                        category_color_map[category] = color_palette[color_idx]
                    
                    block_color = category_color_map.get(category, '#FFFFFF')
                    
                    # 计算块的边界
                    x_left = shelf_x_left + start_idx * block_width
                    x_right = shelf_x_left + end_idx * block_width
                    y_bottom = current_y - block_height
                    y_top = current_y
                    
                    # 使用FancyBboxPatch创建圆角矩形块
                    block = patches.FancyBboxPatch(
                        (x_left + 0.01, y_bottom + 0.01),
                        x_right - x_left - 0.02, block_height - 0.02,
                        boxstyle="round,pad=0.02",
                        linewidth=1.2,
                        edgecolor='#000000',
                        facecolor=block_color,
                        alpha=1.0,
                        zorder=1
                    )
                    ax.add_patch(block)
                    
                    # 添加文本
                    if category:
                        # 根据块的大小调整字体
                        block_width_actual = x_right - x_left
                        if block_width_actual > 0.3:
                            fontsize = 10
                        elif block_width_actual > 0.2:
                            fontsize = 9
                        else:
                            fontsize = 8
                        
                        text = category
                        # 如果文本太长，尝试换行
                        if len(text) > 6 and block_width_actual < 0.25:
                            # 在中间位置找合适的分割点
                            mid = len(text) // 2
                            for j in range(mid, len(text)):
                                if text[j] in ['类', '型', '种', '品', '饮', '料']:
                                    text = text[:j+1] + '\n' + text[j+1:]
                                    break
                        
                        ax.text(
                            (x_left + x_right) / 2,
                            (y_bottom + y_top) / 2,
                            text,
                            ha='center',
                            va='center',
                            fontsize=fontsize,
                            fontweight='bold',
                            color='#000000',
                            zorder=2
                        )
                
                current_y -= 1.0
        
        # 添加标题
        if template_name:
            title = f'{template_name}-{dimension_name}布局图'
        else:
            title = f'{dimension_name}布局图'
        plt.title(title, fontsize=16, fontweight='bold', pad=20)
        
        # 添加图例
        if category_color_map:
            # 创建图例元素
            legend_elements = []
            # 按类别名称排序，确保图例显示有序
            sorted_categories = sorted(category_color_map.items(), key=lambda x: x[0] if x[0] else '')
            
            for category, color in sorted_categories:
                if category:  # 只显示非空类别
                    # 限制类别名称长度，避免图例过长
                    display_name = category[:20] + '...' if len(category) > 20 else category
                    legend_elements.append(
                        patches.Patch(facecolor=color, edgecolor='#000000', linewidth=0.5, label=display_name)
                    )
            
            # 如果类别太多，分多列显示
            ncol = min(4, max(1, len(legend_elements) // 10 + 1))
            
            # 在图的下方添加图例
            if legend_elements:
                fig.legend(
                    handles=legend_elements,
                    loc='lower center',
                    ncol=ncol,
                    fontsize=8,
                    frameon=True,
                    fancybox=True,
                    shadow=True,
                    bbox_to_anchor=(0.5, -0.05),
                    title=f'{dimension_name}图例',
                    title_fontsize=10
                )
        
        plt.tight_layout()
        # 调整布局以容纳图例
        plt.subplots_adjust(bottom=0.15 + 0.05 * (len(category_color_map) // 10))
        plt.savefig(output_path, dpi=300, bbox_inches='tight', facecolor='white')
        print(f"{dimension_name}布局图已保存至: {output_path}")
        plt.close()
    
    def run(self, product_file: str = None, layout_file: str = None):
        """
        运行完整的生成流程
        
        Args:
            product_file: 商品资料表文件路径（可选）
            layout_file: 落位明细清单文件路径（可选）
        """
        try:
            # 1. 加载数据
            self.load_data(product_file, layout_file)
            
            # 2. 匹配数据
            self.match_data()
            
            # 3. 获取货架模板名称
            template_name = self.get_shelf_template_name()
            if template_name:
                print(f"货架模板名称: {template_name}")
            
            # 4. 获取货架信息
            shelf_info, shelf_col, layer_col, position_col = self.get_shelf_info()
            print(f"\n找到 {len(shelf_info)} 个货架")
            for shelf, layers in shelf_info.items():
                print(f"  货架 {shelf}: {len(layers)} 层")
            
            # 5. 生成货架框架图
            if template_name:
                framework_filename = f"{template_name}-货架框架图.png"
            else:
                framework_filename = "货架框架图.png"
            self.generate_shelf_framework(shelf_info, framework_filename)
            
            # 6. 生成五个维度的商品布局图
            dimensions = [
                ("项目中类", "项目中类"),
                ("项目小类", "项目小类"),
                ("项目细类", "项目细类"),
                ("项目商品类别", "销售类别"),  # 字段名是"项目商品类别"，但显示为"销售类别"
                ("品牌名称", "品牌")  # 字段名是"品牌名称"，但显示为"品牌"
            ]
            
            for field_name, display_name in dimensions:
                if template_name:
                    output_filename = f"{template_name}-{display_name}布局图.png"
                else:
                    output_filename = f"{display_name}布局图.png"
                self.generate_product_layout(
                    shelf_info, shelf_col, layer_col, position_col,
                    field_name, display_name, template_name, output_filename
                )
            
            print("\n✅ 所有图纸生成完成！")
            
        except Exception as e:
            print(f"❌ 错误: {str(e)}")
            import traceback
            traceback.print_exc()


def main():
    """主函数"""
    import sys
    
    # 获取文件夹路径
    if len(sys.argv) > 1:
        folder_path = sys.argv[1]
    else:
        folder_path = os.path.dirname(os.path.abspath(__file__))
    
    generator = LayoutGenerator(folder_path)
    generator.run()


if __name__ == "__main__":
    main()
