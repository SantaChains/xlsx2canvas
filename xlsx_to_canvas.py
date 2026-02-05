"""
XLSX表格转单词Canvas工具
功能：将Excel表格中的单词数据转换为Obsidian Canvas格式
作者：SantaChains with Trae
日期：2026-02-04
"""

import sys
import json
import os
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QFileDialog, QGroupBox, QGridLayout, 
    QMessageBox, QTextEdit, QComboBox, QScrollArea
)
from PyQt5.QtCore import Qt
import pandas as pd


class XLSXToCanvasConverter(QMainWindow):
    """
    XLSX转Canvas转换器主窗口
    """
    
    def __init__(self):
        super().__init__()
        self.setWindowTitle('XLSX表格转单词Canvas工具')
        self.setGeometry(100, 100, 900, 700)
        self.setMinimumSize(800, 600)
        
        # 初始化变量
        self.file_path = ''
        self.df = None
        self.columns = []
        
        # 创建中心部件
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # 主布局 - 紧凑布局，移除多余空白
        main_layout = QVBoxLayout(central_widget)
        main_layout.setSpacing(8)
        main_layout.setContentsMargins(12, 12, 12, 12)
        
        # 文件选择区域 - 紧凑布局
        file_section = QGroupBox('文件选择')
        file_layout = QHBoxLayout()
        file_layout.setContentsMargins(8, 8, 8, 8)
        file_layout.setSpacing(8)
        
        self.file_label = QLabel('未选择文件')
        self.file_label.setMinimumWidth(200)
        self.file_label.setStyleSheet('''
            padding: 6px 10px;
            background-color: #f9fafb;
            border: 1px solid #e5e7eb;
            border-radius: 10px;
            font-size: 13px;
            color: #6b7280;
        ''')
        
        select_button = QPushButton('选择XLSX文件')
        select_button.setFixedSize(110, 30)
        select_button.clicked.connect(self.select_file)
        
        file_layout.addWidget(self.file_label)
        file_layout.addWidget(select_button)
        file_layout.addStretch()
        file_section.setLayout(file_layout)
        main_layout.addWidget(file_section, 0)
        
        # 列映射区域
        mapping_section = QGroupBox('列映射')
        mapping_layout = QVBoxLayout()
        mapping_layout.setContentsMargins(8, 8, 8, 8)
        mapping_layout.setSpacing(4)
        
        # 映射控件网格 - 紧凑布局
        mapping_grid = QGridLayout()
        mapping_grid.setSpacing(4)
        mapping_grid.setVerticalSpacing(4)
        mapping_grid.setColumnStretch(1, 1)
        
        # 单词列（必填）
        word_label = QLabel('单词列:')
        word_label.setStyleSheet('font-weight: 600; font-size: 13px; min-width: 80px;')
        self.word_combo = QComboBox()
        self.word_combo.setMinimumWidth(150)
        self.word_combo.setFixedHeight(28)
        self.word_combo.currentTextChanged.connect(self.on_mapping_changed)
        mapping_grid.addWidget(word_label, 0, 0)
        mapping_grid.addWidget(self.word_combo, 0, 1)
        
        # 含义列（必填）
        meaning_label = QLabel('含义列:')
        meaning_label.setStyleSheet('font-weight: 600; font-size: 13px; min-width: 80px;')
        self.meaning_combo = QComboBox()
        self.meaning_combo.setMinimumWidth(150)
        self.meaning_combo.setFixedHeight(28)
        self.meaning_combo.currentTextChanged.connect(self.on_mapping_changed)
        mapping_grid.addWidget(meaning_label, 1, 0)
        mapping_grid.addWidget(self.meaning_combo, 1, 1)
        
        # 音标列（可选）
        phonetic_label = QLabel('音标列:')
        phonetic_label.setStyleSheet('font-size: 13px; min-width: 80px; color: #666;')
        self.phonetic_combo = QComboBox()
        self.phonetic_combo.setMinimumWidth(150)
        self.phonetic_combo.setFixedHeight(28)
        self.phonetic_combo.currentTextChanged.connect(self.on_mapping_changed)
        mapping_grid.addWidget(phonetic_label, 2, 0)
        mapping_grid.addWidget(self.phonetic_combo, 2, 1)
        
        # 派生词列（可选）
        derivatives_label = QLabel('派生词列:')
        derivatives_label.setStyleSheet('font-size: 13px; min-width: 80px; color: #666;')
        self.derivatives_combo = QComboBox()
        self.derivatives_combo.setMinimumWidth(150)
        self.derivatives_combo.setFixedHeight(28)
        self.derivatives_combo.currentTextChanged.connect(self.on_mapping_changed)
        mapping_grid.addWidget(derivatives_label, 3, 0)
        mapping_grid.addWidget(self.derivatives_combo, 3, 1)
        
        # 例句列（可选）
        example_label = QLabel('例句列:')
        example_label.setStyleSheet('font-size: 13px; min-width: 80px; color: #666;')
        self.example_combo = QComboBox()
        self.example_combo.setMinimumWidth(150)
        self.example_combo.setFixedHeight(28)
        self.example_combo.currentTextChanged.connect(self.on_mapping_changed)
        mapping_grid.addWidget(example_label, 4, 0)
        mapping_grid.addWidget(self.example_combo, 4, 1)
        
        mapping_layout.addLayout(mapping_grid)
        
        # 状态标签 - 紧凑
        self.status_label = QLabel('请选择对应的列')
        self.status_label.setStyleSheet('''
            color: #666; 
            font-size: 12px;
            padding: 4px 8px;
            background-color: #f5f5f5;
            border-radius: 10px;
            margin-top: 4px;
        ''')
        mapping_layout.addWidget(self.status_label)
        
        # 保存控件引用
        self.mapping_widgets = {
            'word': self.word_combo,
            'meaning': self.meaning_combo,
            'phonetic': self.phonetic_combo,
            'derivatives': self.derivatives_combo,
            'example': self.example_combo
        }
        
        mapping_section.setLayout(mapping_layout)
        main_layout.addWidget(mapping_section, 0)
        
        # 预览区域 - 自适应高度
        preview_section = QGroupBox('数据预览')
        preview_layout = QVBoxLayout()
        preview_layout.setContentsMargins(8, 8, 8, 8)
        
        self.preview_text = QTextEdit()
        self.preview_text.setReadOnly(True)
        self.preview_text.setMinimumHeight(80)
        self.preview_text.setPlaceholderText("请先选择Excel文件，加载的数据将在这里显示。")
        self.preview_text.setStyleSheet('''
            QTextEdit {
                background-color: white;
                border: 1px solid #e5e7eb;
                border-radius: 12px;
                padding: 8px;
                font-size: 13px;
                font-family: "Microsoft YaHei", "Segoe UI", sans-serif;
            }
        ''')
        preview_layout.addWidget(self.preview_text)
        preview_section.setLayout(preview_layout)
        main_layout.addWidget(preview_section, 1)
        
        # 操作按钮区域 - 紧凑布局
        button_section = QGroupBox('操作')
        button_layout = QHBoxLayout()
        button_layout.setContentsMargins(8, 8, 8, 8)
        button_layout.setSpacing(12)
        button_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        help_button = QPushButton('使用帮助')
        help_button.setFixedSize(90, 28)
        help_button.clicked.connect(self.show_help)
        
        reverse_button = QPushButton('Canvas转XLSX')
        reverse_button.setFixedSize(110, 28)
        reverse_button.clicked.connect(self.reverse_canvas_to_xlsx)
        
        clear_button = QPushButton('清空')
        clear_button.setFixedSize(70, 28)
        clear_button.clicked.connect(self.clear_all)
        
        generate_button = QPushButton('生成Canvas文件')
        generate_button.setFixedSize(120, 28)
        generate_button.clicked.connect(self.generate_canvas)
        
        button_layout.addWidget(help_button)
        button_layout.addWidget(reverse_button)
        button_layout.addWidget(clear_button)
        button_layout.addWidget(generate_button)
        button_section.setLayout(button_layout)
        main_layout.addWidget(button_section, 0)
        
        # 状态标签 - 紧凑
        status_layout = QHBoxLayout()
        status_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.status_label = QLabel('就绪')
        self.status_label.setStyleSheet('''
            color: #6b7280;
            font-size: 12px;
            padding: 4px 12px;
            background-color: white;
            border: 1px solid #e5e7eb;
            border-radius: 12px;
        ''')
        self.status_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        status_layout.addWidget(self.status_label)
        main_layout.addLayout(status_layout)
        
        # 设置样式
        self.setStyleSheet('''
            QMainWindow {
                background-color: #f3f4f6;
            }
            QGroupBox {
                font-weight: bold;
                font-size: 16px;
                color: #374151;
                border: 1px solid #e5e7eb;
                border-radius: 16px;
                margin-top: 10px;
                padding-top: 16px;
                padding-bottom: 12px;
                padding-left: 12px;
                padding-right: 12px;
                background-color: white;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 12px;
                padding: 0 6px;
                color: #374151;
            }
            QPushButton {
                background-color: #3b82f6;
                color: white;
                border: none;
                border-radius: 10px;
                font-size: 14px;
                font-weight: 500;
            }
            QPushButton:hover {
                background-color: #2563eb;
            }
            QPushButton:pressed {
                background-color: #1d4ed8;
            }
            QComboBox {
                background-color: white;
                border: 1px solid #d1d5db;
                border-radius: 10px;
                padding: 6px 10px;
                font-size: 14px;
                min-height: 32px;
            }
            QComboBox:hover {
                border-color: #3b82f6;
            }
            QComboBox::drop-down {
                border-radius: 10px;
            }
            QLabel {
                color: #374151;
                font-size: 14px;
            }
        ''')
    
    def select_file(self):
        """选择XLSX文件"""
        file_dialog = QFileDialog()
        file_path, _ = file_dialog.getOpenFileName(
            self, "选择XLSX文件", "", "Excel文件 (*.xlsx)"
        )
        
        if file_path:
            self.file_path = file_path
            self.file_label.setText(os.path.basename(file_path))
            self.load_xlsx()
    
    def load_xlsx(self):
        """加载XLSX文件"""
        try:
            # 获取所有sheet名称
            xl_file = pd.ExcelFile(self.file_path, engine='openpyxl')
            sheet_names = xl_file.sheet_names
            
            # 尝试读取第一个有数据的sheet
            self.df = None
            for sheet_name in sheet_names:
                df_temp = pd.read_excel(self.file_path, sheet_name=sheet_name, engine='openpyxl')
                if not df_temp.empty:
                    self.df = df_temp
                    break
            
            # 检查是否为空
            if self.df is None or self.df.empty:
                QMessageBox.warning(self, "警告", "Excel文件为空或没有数据\n请检查文件是否包含数据")
                self.status_label.setText("文件为空")
                return
            
            # 检测是否没有标题行（列名是数据值）
            first_row = self.df.iloc[0].tolist()
            columns = list(self.df.columns)
            
            # 如果列名包含数据值（看起来像数据而不是标题），说明没有标题行
            is_header_missing = False
            for i, col in enumerate(columns):
                if pd.isna(col) or str(col).strip() == '':
                    # 空列名，假设没有标题行
                    is_header_missing = True
                    break
            
            # 特殊检测：如果第一行的值看起来像数据（如中国人、美国人等）
            if not is_header_missing:
                # 检查列名是否和第一行数据相同
                first_data_row = [str(self.df.iloc[0][col]).strip() for col in columns]
                if first_data_row == [str(col).strip() for col in columns]:
                    is_header_missing = True
            
            if is_header_missing:
                # 没有标题行，使用列索引作为列名
                self.df.columns = [f'列{i+1}' for i in range(len(columns))]
                columns = list(self.df.columns)
                QMessageBox.information(self, "提示", 
                    "检测到Excel文件没有标题行，已自动添加默认列名\n"
                    "请在映射区域选择对应的列")
            
            self.columns = columns
            
            # 更新列映射控件
            self.set_columns(self.columns)
            
            # 预览数据
            preview_text = f"共 {len(self.df)} 行数据\n"
            preview_text += f"列名：{', '.join(self.columns)}\n\n"
            preview_text += "前5行数据预览：\n"
            preview_text += str(self.df.head())
            self.preview_text.setText(preview_text)
            
            self.status_label.setText(f"成功加载文件，共 {len(self.df)} 条记录")
            
        except Exception as e:
            error_msg = f"加载文件失败：{str(e)}"
            QMessageBox.critical(self, "错误", error_msg)
            self.status_label.setText("加载失败")
            # 清空数据
            self.df = None
            self.columns = []
            self.set_columns([])
    
    def clear_all(self):
        """清空所有内容"""
        self.file_path = ''
        self.df = None
        self.columns = []
        
        self.file_label.setText('未选择文件')
        self.set_columns([])
        self.preview_text.clear()
        self.status_label.setText('就绪')
    
    def generate_canvas(self):
        """生成Canvas文件"""
        if not self.file_path or self.df is None:
            QMessageBox.warning(self, "警告", "请先选择并加载XLSX文件")
            return
        
        # 获取映射配置
        mapping = self.get_mapping()
        
        # 验证必填字段
        errors = self.validate_mapping()
        if errors:
            QMessageBox.warning(self, "警告", "\n".join(errors))
            return
        
        # 准备数据
        nodes = []
        x_pos = -440
        y_pos = -100
        
        for idx, row in self.df.iterrows():
            word = str(row[mapping['word']]).strip()
            if not word:
                continue
            
            # 构建文本内容
            text = word
            
            # 音标
            if mapping['phonetic'] and mapping['phonetic'] in self.df.columns:
                phonetic = str(row[mapping['phonetic']]).strip()
                if phonetic and phonetic != 'nan':
                    text += f"\n[{phonetic}]"
            
            # 派生词
            if mapping['derivatives'] and mapping['derivatives'] in self.df.columns:
                derivatives = str(row[mapping['derivatives']]).strip()
                if derivatives and derivatives != 'nan':
                    text += f"\n*{derivatives}*"
            
            # 含义
            meaning = str(row[mapping['meaning']]).strip()
            if meaning and meaning != 'nan':
                text += f"\n**含义：** {meaning}"
            
            # 例句
            if mapping['example'] and mapping['example'] in self.df.columns:
                example = str(row[mapping['example']]).strip()
                if example and example != 'nan':
                    text += f"\n**例句：** {example}"
            
            # 创建节点
            node = {
                "id": str(idx + 1),
                "type": "text",
                "text": text,
                "x": x_pos,
                "y": y_pos,
                "width": 565,
                "height": 300
            }
            nodes.append(node)
            
            # 更新位置
            x_pos += 600
            if x_pos > 800:
                x_pos = -440
                y_pos += 340
        
        # 生成canvas数据
        canvas_data = {"nodes": nodes, "edges": []}
        
        # 保存文件
        save_path, _ = QFileDialog.getSaveFileName(
            self, "保存Canvas文件", "单词.canvas", "Canvas文件 (*.canvas)"
        )
        
        if save_path:
            try:
                with open(save_path, 'w', encoding='utf-8') as f:
                    json.dump(canvas_data, f, indent=1, ensure_ascii=False)
                
                QMessageBox.information(self, "成功", f"Canvas文件已生成：{save_path}")
                self.status_label.setText(f"成功生成Canvas文件，包含 {len(nodes)} 个单词")
                
            except Exception as e:
                QMessageBox.critical(self, "错误", f"保存文件失败：{str(e)}")
    
    def set_columns(self, columns):
        """设置可用的列"""
        for combo in self.mapping_widgets.values():
            combo.clear()
            combo.addItem('未选择', '')
        
        if columns:
            for col in columns:
                for combo in self.mapping_widgets.values():
                    combo.addItem(col, col)
            for combo in self.mapping_widgets.values():
                combo.setEnabled(True)
        else:
            for combo in self.mapping_widgets.values():
                combo.addItem('请先选择Excel文件', '')
                combo.setEnabled(False)
    
    def get_mapping(self):
        """获取当前映射"""
        mapping = {}
        for attr, combo in self.mapping_widgets.items():
            mapping[attr] = combo.currentData() or ''
        return mapping
    
    def validate_mapping(self):
        """验证映射是否正确"""
        errors = []
        mapping = self.get_mapping()
        
        if not mapping['word']:
            errors.append('请选择单词列')
        if not mapping['meaning']:
            errors.append('请选择含义列')
        
        return errors
    
    def on_mapping_changed(self):
        """映射改变时的回调"""
        pass
    
    def reverse_canvas_to_xlsx(self):
        """Canvas逆向转XLSX功能"""
        # 选择Canvas文件
        canvas_path, _ = QFileDialog.getOpenFileName(
            self, "选择Canvas文件", "", "Canvas文件 (*.canvas)"
        )
        
        if not canvas_path:
            return
        
        try:
            # 读取Canvas文件
            with open(canvas_path, 'r', encoding='utf-8') as f:
                canvas_data = json.load(f)
            
            nodes = canvas_data.get('nodes', [])
            if not nodes:
                QMessageBox.warning(self, "警告", "Canvas文件中没有节点数据")
                return
            
            # 解析节点数据
            data = []
            for node in nodes:
                if node.get('type') != 'text':
                    continue
                
                text = node.get('text', '')
                if not text:
                    continue
                
                # 解析文本内容
                lines = text.split('\n')
                row_data = {
                    '单词': '',
                    '音标': '',
                    '派生词': '',
                    '含义': '',
                    '例句': ''
                }
                
                # 第1行：单词
                if lines:
                    row_data['单词'] = lines[0].strip()
                
                # 解析其他行
                for line in lines[1:]:
                    line_stripped = line.strip()
                    if not line_stripped:
                        continue
                    
                    # 音标：[xxx]
                    if line_stripped.startswith('[') and line_stripped.endswith(']'):
                        row_data['音标'] = line_stripped[1:-1]
                    # 派生词：*xxx*
                    elif line_stripped.startswith('*') and line_stripped.endswith('*'):
                        row_data['派生词'] = line_stripped[1:-1]
                    # 例句：**例句：** xxx - 必须先检查例句
                    elif line_stripped.startswith('**例句：**'):
                        row_data['例句'] = line_stripped.replace('**例句：**', '').strip()
                    # 含义：**含义：** xxx 或 **n.** xxx 等词性标注
                    elif line_stripped.startswith('**含义：**'):
                        row_data['含义'] = line_stripped.replace('**含义：**', '').strip()
                    elif line_stripped.startswith('**') and (
                        line_stripped[2:4] in ['n.', 'v.', 'a.', 'ad.', 'pr.', 'co']
                    ):
                        # 匹配 **n.** 含义 或 **v.** 含义 等格式
                        row_data['含义'] = line_stripped.replace('**', '').strip()
                
                data.append(row_data)
            
            if not data:
                QMessageBox.warning(self, "警告", "未能从Canvas文件中提取有效数据")
                return
            
            # 创建DataFrame
            df = pd.DataFrame(data)
            
            # 选择保存位置
            save_path, _ = QFileDialog.getSaveFileName(
                self, "保存XLSX文件", "words_from_canvas.xlsx", "Excel文件 (*.xlsx)"
            )
            
            if save_path:
                df.to_excel(save_path, index=False, engine='openpyxl')
                QMessageBox.information(self, "成功", f"XLSX文件已生成：{save_path}\n共转换 {len(data)} 个单词")
                self.status_label.setText(f"成功转换 {len(data)} 个单词到XLSX")
                
        except Exception as e:
            QMessageBox.critical(self, "错误", f"转换失败：{str(e)}")
    
    def show_help(self):
        """显示使用帮助"""
        help_text = """XLSX表格转单词Canvas工具 - 使用帮助

[说明] XLSX转Canvas步骤：
1. 点击"选择XLSX文件"按钮，选择包含单词数据的Excel文件
2. 在"列映射配置"区域，将Excel表格的列与目标属性进行匹配：
   - 单词列 (必填)：包含单词本身的列
   - 含义列 (必填)：包含单词含义的列
   - 音标列 (可选)：包含音标的列
   - 派生词列 (可选)：包含派生词的列
   - 例句列 (可选)：包含例句的列
3. 确认映射配置正确（状态显示"OK 映射配置正确"）
4. 点击"生成Canvas文件"按钮，选择保存位置

[多语言支持]
本工具支持各种语言的单词转换，包括但不限于：
- 中文：中国人 [zhong guo ren] 含义：名词,中国人
- 日语：日本人 [nihonjin] 含义：名词,日本人
- 韩语：Hangukin [hangukin] 含义：名词,韩国人
- 英语：American [american] 含义：n.美国人
- 法语：Francais [franse] 含义：n.法国人

只需将对应语言的单词、音标、含义分别放入Excel的不同列即可。

[Canvas转XLSX步骤]
1. 点击"Canvas转XLSX"按钮
2. 选择要转换的.canvas文件
3. 选择保存XLSX文件的位置
4. 转换完成后会生成包含单词、音标、派生词、含义、例句的Excel文件

[文件格式说明]
- Excel文件：第一行为列标题，数据从第二行开始，支持.xlsx格式
- Canvas文件：可在Obsidian中直接打开，每个单词生成一个文本节点

[Canvas文本格式]
第1行：单词
第2行：[音标]（可选）
第3行：*派生词*（可选）
第4行：**含义：** xxx
第5行：**例句：** xxx（可选）

[提示]
- 可以先预览数据确认格式正确
- 建议使用标准的Excel格式，避免合并单元格
- 生成的Canvas文件可以进一步在Obsidian中编辑和美化
- 支持双向转换：XLSX <-> Canvas"""

        QMessageBox.information(self, "使用帮助", help_text)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    
    window = XLSXToCanvasConverter()
    window.show()
    
    sys.exit(app.exec_())
