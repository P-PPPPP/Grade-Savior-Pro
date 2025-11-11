import sys
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
from PyQt5.QtWidgets import (QApplication, QMainWindow, QVBoxLayout, QHBoxLayout, 
                             QWidget, QPushButton, QLabel, QSlider, QDoubleSpinBox, 
                             QListWidget, QFileDialog, QGroupBox, QFormLayout, 
                             QSplitter, QMessageBox, QProgressDialog)
from PyQt5.QtCore import Qt
import matplotlib
from datetime import datetime

# 设置matplotlib使用支持中文的字体
matplotlib.rcParams['font.sans-serif'] = ['SimHei', 'Microsoft YaHei', 'Arial Unicode MS']
matplotlib.rcParams['axes.unicode_minus'] = False  # 解决负号显示问题

class GradeAdjuster(QMainWindow):
    def __init__(self):
        super().__init__()
        self.df = None
        self.original_scores = None
        self.current_adjusted_scores = None
        self.init_ui()
        
    def init_ui(self):
        self.setWindowTitle("全自动捞人神器 v1.0")
        self.setGeometry(100, 100, 1400, 800)
        
        # 创建中央部件
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # 主布局
        main_layout = QHBoxLayout(central_widget)
        
        # 左侧控制面板
        left_panel = self.create_left_panel()
        main_layout.addWidget(left_panel, 1)
        
        # 右侧图表和数据显示区域
        right_panel = self.create_right_panel()
        main_layout.addWidget(right_panel, 2)
        
        # 初始化状态
        self.update_display()
        
    def create_left_panel(self):
        panel = QWidget()
        layout = QVBoxLayout(panel)
        
        # 文件操作区域
        file_group = QGroupBox("文件操作")
        file_layout = QVBoxLayout(file_group)
        
        self.load_btn = QPushButton("加载Excel文件")
        self.load_btn.clicked.connect(self.load_excel)
        file_layout.addWidget(self.load_btn)
        
        # 导出按钮
        self.export_btn = QPushButton("导出修改后成绩")
        self.export_btn.clicked.connect(self.export_scores)
        self.export_btn.setEnabled(False)  # 初始不可用
        file_layout.addWidget(self.export_btn)
        
        self.file_label = QLabel("未加载文件")
        file_layout.addWidget(self.file_label)
        layout.addWidget(file_group)
        
        # 参数调整区域
        param_group = QGroupBox("成绩调整参数")
        param_layout = QFormLayout(param_group)
        
        # 均值调整滑块
        self.mean_slider = QSlider(Qt.Horizontal)
        self.mean_slider.setMinimum(-20)
        self.mean_slider.setMaximum(20)
        self.mean_slider.setValue(0)
        self.mean_slider.valueChanged.connect(self.update_display)
        param_layout.addRow("整体平移:", self.mean_slider)
        
        self.mean_value = QLabel("0")
        param_layout.addRow("平移量:", self.mean_value)
        
        # 标准差调整滑块
        self.std_slider = QSlider(Qt.Horizontal)
        self.std_slider.setMinimum(-10)
        self.std_slider.setMaximum(10)
        self.std_slider.setValue(0)
        self.std_slider.valueChanged.connect(self.update_display)
        param_layout.addRow("分布调整:", self.std_slider)
        
        self.std_value = QLabel("0")
        param_layout.addRow("调整量:", self.std_value)
        
        # 及格率调整滑块
        self.pass_rate_slider = QSlider(Qt.Horizontal)
        self.pass_rate_slider.setMinimum(-30)
        self.pass_rate_slider.setMaximum(30)
        self.pass_rate_slider.setValue(0)
        self.pass_rate_slider.valueChanged.connect(self.update_display)
        param_layout.addRow("及格率调整:", self.pass_rate_slider)
        
        self.pass_rate_value = QLabel("0")
        param_layout.addRow("调整量:", self.pass_rate_value)
        
        layout.addWidget(param_group)
        
        # 统计信息区域
        stats_group = QGroupBox("统计信息")
        stats_layout = QFormLayout(stats_group)
        
        self.mean_label = QLabel("0")
        self.min_label = QLabel("0")
        self.max_label = QLabel("0")
        self.pass_count_label = QLabel("0")
        self.fail_count_label = QLabel("0")
        self.pass_rate_label = QLabel("0%")
        
        stats_layout.addRow("平均分:", self.mean_label)
        stats_layout.addRow("最低分:", self.min_label)
        stats_layout.addRow("最高分:", self.max_label)
        stats_layout.addRow("及格人数:", self.pass_count_label)
        stats_layout.addRow("不及格人数:", self.fail_count_label)
        stats_layout.addRow("及格率:", self.pass_rate_label)
        
        layout.addWidget(stats_group)
        
        # 不及格学生列表
        fail_group = QGroupBox("不及格学生名单")
        fail_layout = QVBoxLayout(fail_group)
        
        self.fail_list = QListWidget()
        fail_layout.addWidget(self.fail_list)
        
        layout.addWidget(fail_group)
        
        return panel
    
    def create_right_panel(self):
        panel = QWidget()
        layout = QVBoxLayout(panel)
        
        # 创建matplotlib图形
        self.figure = Figure(figsize=(10, 8), dpi=100)
        self.canvas = FigureCanvas(self.figure)
        layout.addWidget(self.canvas)
        
        return panel
    
    def load_excel(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择Excel文件", "", "Excel Files (*.xlsx *.xls)")
        
        if file_path:
            try:
                # 读取Excel文件
                self.df = pd.read_excel(file_path, sheet_name=0)
                
                # 检查必要的列是否存在
                if 'name' not in self.df.columns or 'score' not in self.df.columns:
                    QMessageBox.warning(self, "错误", "Excel文件必须包含'name'和'score'列")
                    return
                
                # 保存原始成绩
                self.original_scores = self.df['score'].copy()
                self.file_label.setText(f"已加载: {file_path.split('/')[-1]}")
                
                # 启用导出按钮
                self.export_btn.setEnabled(True)
                
                # 重置滑块
                self.mean_slider.setValue(0)
                self.std_slider.setValue(0)
                self.pass_rate_slider.setValue(0)
                
                # 更新显示
                self.update_display()
                
            except Exception as e:
                QMessageBox.warning(self, "错误", f"读取文件失败: {str(e)}")
    
    def adjust_scores(self):
        if self.df is None:
            return None
        
        # 获取调整参数
        mean_adjust = self.mean_slider.value()  # 整体平移
        std_adjust = self.std_slider.value() / 10.0  # 分布调整
        pass_rate_adjust = self.pass_rate_slider.value()  # 及格率调整
        
        # 复制原始成绩
        adjusted_scores = self.original_scores.copy()
        
        # 应用调整
        if len(adjusted_scores) > 1:
            # 1. 整体平移调整
            adjusted_scores = adjusted_scores + mean_adjust
            
            # 2. 分布调整（保持排序）
            if std_adjust != 0:
                current_std = adjusted_scores.std()
                if current_std > 0:
                    # 调整标准差，但保持最小值不低于0.1
                    target_std = max(current_std + std_adjust, 0.1)
                    # 转换为Z分数，然后按新标准差缩放
                    z_scores = (adjusted_scores - adjusted_scores.mean()) / current_std
                    adjusted_scores = z_scores * target_std + adjusted_scores.mean()
            
            # 3. 及格率调整 - 新的更有效的方法
            if pass_rate_adjust != 0:
                # 计算当前及格线附近的学生
                current_scores = adjusted_scores.copy()
                
                # 找出50-70分之间的学生（及格线附近）
                near_pass_mask = (current_scores >= 50) & (current_scores < 70)
                near_pass_scores = current_scores[near_pass_mask]
                
                if len(near_pass_scores) > 0:
                    # 根据及格率调整参数，对这些分数进行非线性调整
                    # 正调整：提升接近及格的学生，负调整：降低接近及格的学生
                    adjustment_factor = pass_rate_adjust / 10.0  # 缩放因子
                    
                    # 对接近及格的学生应用调整，离及格线越近调整越大
                    for i in near_pass_scores.index:
                        score = current_scores[i]
                        # 计算与60分的距离，距离越近调整越大
                        distance_to_pass = 60 - score
                        # 使用非线性函数，使调整更平滑
                        if distance_to_pass > 0:  # 低于60分
                            adjustment = adjustment_factor * (1 - distance_to_pass / 10.0)
                        else:  # 高于60分但接近
                            adjustment = adjustment_factor * (1 - abs(distance_to_pass) / 10.0)
                        
                        # 应用调整，但确保不会过度调整
                        current_scores[i] = score + adjustment * 5
                    
                    adjusted_scores = current_scores
        
        # 确保分数在0-100之间
        adjusted_scores = np.clip(adjusted_scores, 0, 100)
        
        # 保存当前调整后的成绩
        self.current_adjusted_scores = adjusted_scores
        
        return adjusted_scores
    
    def update_display(self):
        if self.df is None:
            return
        
        # 更新参数显示
        self.mean_value.setText(str(self.mean_slider.value()))
        self.std_value.setText(f"{self.std_slider.value() / 10.0:.1f}")
        self.pass_rate_value.setText(str(self.pass_rate_slider.value()))
        
        # 调整成绩
        adjusted_scores = self.adjust_scores()
        if adjusted_scores is None:
            return
        
        # 更新统计信息
        mean_score = adjusted_scores.mean()
        min_score = adjusted_scores.min()
        max_score = adjusted_scores.max()
        pass_count = (adjusted_scores >= 60).sum()
        fail_count = (adjusted_scores < 60).sum()
        pass_rate = pass_count / len(adjusted_scores) * 100
        
        self.mean_label.setText(f"{mean_score:.1f}")
        self.min_label.setText(f"{min_score:.1f}")
        self.max_label.setText(f"{max_score:.1f}")
        self.pass_count_label.setText(str(pass_count))
        self.fail_count_label.setText(str(fail_count))
        self.pass_rate_label.setText(f"{pass_rate:.1f}%")
        
        # 更新不及格学生名单
        self.fail_list.clear()
        fail_mask = adjusted_scores < 60
        fail_students = self.df[fail_mask].copy()
        fail_students['adjusted_score'] = adjusted_scores[fail_mask]
        
        for _, student in fail_students.iterrows():
            item_text = f"{student['name']}: {student['adjusted_score']:.1f}分"
            self.fail_list.addItem(item_text)
        
        # 更新图表
        self.update_chart(adjusted_scores)
    
    def update_chart(self, scores):
        self.figure.clear()
        
        # 创建子图布局
        gs = self.figure.add_gridspec(2, 2, width_ratios=[2, 1], height_ratios=[1, 1])
        
        # 主柱状图
        ax1 = self.figure.add_subplot(gs[:, 0])
        
        # 计算分数区间
        bins = [0, 10, 20, 30, 40, 50, 60, 70, 80, 90, 100]
        bin_labels = ['0-10', '10-20', '20-30', '30-40', '40-50', 
                     '50-60', '60-70', '70-80', '80-90', '90-100']
        
        hist, bin_edges = np.histogram(scores, bins=bins)
        
        # 绘制柱状图
        colors = ['lightcoral'] * 6 + ['lightgreen'] * 4  # 不及格红色，及格绿色
        bars = ax1.bar(bin_labels, hist, color=colors, edgecolor='black', alpha=0.7)
        
        # 在柱子上显示数量
        for bar, count in zip(bars, hist):
            height = bar.get_height()
            if height > 0:
                ax1.text(bar.get_x() + bar.get_width()/2., height + 0.1,
                        f'{int(count)}', ha='center', va='bottom')
        
        ax1.set_xlabel('分数区间')
        ax1.set_ylabel('学生数量')
        ax1.set_title('学生成绩分布')
        ax1.grid(True, alpha=0.3)
        
        # 添加及格线标记
        ax1.axvline(x=5.5, color='red', linestyle='--', alpha=0.7)
        ax1.text(5.5, ax1.get_ylim()[1]*0.9, '及格线', rotation=90, 
                ha='right', va='top', color='red')
        
        # 及格/不及格饼图
        ax2 = self.figure.add_subplot(gs[0, 1])
        pass_fail_counts = [len(scores[scores >= 60]), len(scores[scores < 60])]
        colors = ['lightgreen', 'lightcoral']
        labels = ['及格 (≥60)', '不及格 (<60)']
        ax2.pie(pass_fail_counts, labels=labels, colors=colors, autopct='%1.1f%%', startangle=90)
        ax2.set_title('及格率分布')
        
        # 统计信息文本
        ax3 = self.figure.add_subplot(gs[1, 1])
        ax3.axis('off')
        
        stats_text = (
            f"总人数: {len(scores)}\n"
            f"平均分: {scores.mean():.1f}\n"
            f"最高分: {scores.max():.1f}\n"
            f"最低分: {scores.min():.1f}\n"
            f"标准差: {scores.std():.1f}\n"
            f"及格率: {(scores >= 60).sum()/len(scores)*100:.1f}%"
        )
        
        ax3.text(0.1, 0.9, stats_text, transform=ax3.transAxes, fontsize=12,
                verticalalignment='top', bbox=dict(boxstyle='round', facecolor='wheat', alpha=0.5))
        
        self.figure.tight_layout()
        self.canvas.draw()
    
    def export_scores(self):
        if self.df is None or self.current_adjusted_scores is None:
            QMessageBox.warning(self, "警告", "没有可导出的数据，请先加载文件并调整成绩")
            return
        
        # 选择保存路径
        file_path, _ = QFileDialog.getSaveFileName(
            self, "保存修改后的成绩", "", "Excel Files (*.xlsx)")
        
        if file_path:
            try:
                # 创建导出数据的DataFrame
                export_df = self.df.copy()
                export_df['adjusted_score'] = self.current_adjusted_scores
                export_df['score_change'] = self.current_adjusted_scores - self.original_scores
                
                # 添加统计信息作为新工作表
                stats_data = {
                    '统计项': ['平均分', '最低分', '最高分', '及格人数', '不及格人数', '及格率'],
                    '数值': [
                        f"{self.current_adjusted_scores.mean():.2f}",
                        f"{self.current_adjusted_scores.min():.2f}",
                        f"{self.current_adjusted_scores.max():.2f}",
                        f"{(self.current_adjusted_scores >= 60).sum()}",
                        f"{(self.current_adjusted_scores < 60).sum()}",
                        f"{(self.current_adjusted_scores >= 60).sum()/len(self.current_adjusted_scores)*100:.2f}%"
                    ]
                }
                stats_df = pd.DataFrame(stats_data)
                
                # 添加调整参数信息
                params_data = {
                    '调整参数': ['整体平移', '分布调整', '及格率调整'],
                    '参数值': [
                        f"{self.mean_slider.value()}",
                        f"{self.std_slider.value() / 10.0:.1f}",
                        f"{self.pass_rate_slider.value()}"
                    ]
                }
                params_df = pd.DataFrame(params_data)
                
                # 使用ExcelWriter创建多工作表Excel文件
                with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                    export_df.to_excel(writer, sheet_name='调整后成绩', index=False)
                    stats_df.to_excel(writer, sheet_name='统计信息', index=False)
                    params_df.to_excel(writer, sheet_name='调整参数', index=False)
                
                # 显示成功消息
                QMessageBox.information(self, "成功", f"成绩已成功导出到:\n{file_path}")
                
            except Exception as e:
                QMessageBox.warning(self, "错误", f"导出失败: {str(e)}")

def main():
    app = QApplication(sys.argv)
    
    # 设置应用程序样式
    app.setStyle('Fusion')
    
    window = GradeAdjuster()
    window.show()
    
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()