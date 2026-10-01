# ============================================================
# IMAGE CLASSIFICATION USING PYQT5
# MODEL : RESNET18 - CIFAR-10
# ============================================================

import sys
import numpy as np
import time
import json
import os

import torch
import torch.nn as nn
import torchvision.transforms as transforms
import torchvision.models as models

from PIL import Image

import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas

# PyQt5 Imports
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QListWidget, QStackedWidget, QPushButton, 
                             QLabel, QFileDialog, QProgressBar, QGroupBox,
                             QTableWidget, QTableWidgetItem, QHeaderView, 
                             QAbstractItemView, QMessageBox)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QPixmap, QFont, QColor


# ============================================================
# 1. KONFIGURASI HALAMAN (Window Setup)
# ============================================================
# (Handled in MainWindow init)


# ============================================================
# 2. LABEL CIFAR-10
# ============================================================

labels = [
    "airplane", "automobile", "bird", "cat", "deer",
    "dog", "frog", "horse", "ship", "truck"
]


# ============================================================
# 3. LOKASI FILE
# ============================================================

# ============================================================
# 3. LOKASI FILE (PyInstaller Compatible)
# ============================================================

def get_base_dir():
    # If running as a bundled .exe, use the temporary PyInstaller path
    if getattr(sys, 'frozen', False):
        return sys._MEIPASS
    # If running normally in Python, use the script's directory
    else:
        return os.path.dirname(os.path.abspath(__file__))

BASE_DIR = get_base_dir()
MODEL_PATH = os.path.join(BASE_DIR, "cifar10_trained_V3.pth")
HISTORY_PATH = os.path.join(BASE_DIR, "model5_architecture.json")

# ============================================================
# 4. DEVICE
# ============================================================

device = torch.device("cuda" if torch.cuda.is_available() else "cpu")


# ============================================================
# 5. TRANSFORMASI GAMBAR
# ============================================================

transform = transforms.Compose([
    transforms.Resize((32, 32)),
    transforms.ToTensor(),
    transforms.Normalize((0.4914, 0.4822, 0.4465), (0.2023, 0.1994, 0.2010))
])


# ============================================================
# 6. LOAD MODEL
# ============================================================

def load_cifar_model():
    if not os.path.exists(MODEL_PATH):
        print(f"File model tidak ditemukan:\n{MODEL_PATH}")
        sys.exit(1)

    model = models.resnet18()
    num_ftrs = model.fc.in_features
    model.fc = nn.Linear(num_ftrs, 10)

    state_dict = torch.load(MODEL_PATH, map_location=device, weights_only=False)
    model.load_state_dict(state_dict)

    model = model.to(device)
    model.eval()
    return model


# ============================================================
# 7. LOAD HISTORY
# ============================================================

def load_history():
    if not os.path.exists(HISTORY_PATH):
        return {"accuracy": [], "val_accuracy": []}

    with open(HISTORY_PATH, "r") as f:
        history = json.load(f)
    return history


# ============================================================
# 8. LOAD MODEL DAN HISTORY
# ============================================================

model = load_cifar_model()
history = load_history()


# ============================================================
# 9. SESSION STATE (Global Dictionary for PyQt)
# ============================================================

session_state = {
    "class_counts": {label: 0 for label in labels},
    "exit_flag": False
}


# ============================================================
# 10. UPDATE TABLE
# ============================================================

def update_table(class_counts):
    max_label_length = max(len(label) for label in labels)
    max_count_length = max(len(str(count)) for count in class_counts.values())
    table_width = max(max_label_length, max_count_length) + 2
    column_separator_length = 3

    line_length = table_width * len(labels) + column_separator_length * (len(labels) - 1) + 2
    table = "-" * line_length + "\n"
    table += "| " + " | ".join(label.center(table_width) for label in labels) + " |\n"
    table += "-" * line_length + "\n"
    table += "| " + " | ".join(str(class_counts[label]).center(table_width) for label in labels) + " |\n"
    table += "-" * line_length
    return table


def update_table_widget(table_widget, class_counts):
    """Updates the QTableWidget with current counts."""
    for col, label in enumerate(labels):
        item = QTableWidgetItem(str(class_counts[label]))
        item.setTextAlignment(Qt.AlignCenter)
        item.setFont(QFont("Courier New", 14, QFont.Bold))
        table_widget.setItem(0, col, item)


# ============================================================
# 11. RESET TABLE
# ============================================================

def reset_table():
    for label in session_state["class_counts"]:
        session_state["class_counts"][label] = 0
    return update_table(session_state["class_counts"])


# ============================================================
# 12. PROSES PREDIKSI (Adapted for PyQt Widgets)
# ============================================================

def process_image(
    file_path,
    progress_bar,
    prediction_label,
    confidence_label,
    time_label,
    table_widget,
    confidence_table_widget  # Added for the new confidence table
):
    progress_bar.setVisible(True)
    for progress_percent in range(101):
        progress_bar.setValue(progress_percent)
        QApplication.processEvents()
        time.sleep(0.005)

    start_time = time.time()
    img = Image.open(file_path).convert("RGB")

    input_tensor = transform(img)
    input_tensor = input_tensor.unsqueeze(0)
    input_tensor = input_tensor.to(device)

    with torch.no_grad():
        outputs = model(input_tensor)
        probabilities = torch.softmax(outputs, dim=1)
        confidence, prediction = torch.max(probabilities, dim=1)

    predicted_index = prediction.item()
    predicted_class = labels[predicted_index]
    confidence_value = confidence.item() * 100

    end_time = time.time()
    estimation_time = end_time - start_time

    prediction_label.setText(f"Classification: {predicted_class}")
    confidence_label.setText(f"Confidence: {confidence_value:.2f}%")
    time_label.setText(f"Estimation Time: {estimation_time:.2f} seconds")

    session_state["class_counts"][predicted_class] += 1

    update_table_widget(table_widget, session_state["class_counts"])
    
    # NEW: Update the confidence table for all categories
    probs = probabilities[0].cpu().numpy() * 100
    for i, label in enumerate(labels):
        class_item = QTableWidgetItem(label)
        class_item.setTextAlignment(Qt.AlignCenter)
        conf_item = QTableWidgetItem(f"{probs[i]:.2f}%")
        conf_item.setTextAlignment(Qt.AlignCenter)
        # Highlight the predicted class
        if i == predicted_index:
            conf_item.setBackground(QColor(46, 204, 113)) # Green for prediction
            conf_item.setForeground(Qt.white)
            conf_item.setFont(QFont("Arial", 10, QFont.Bold))
            
        confidence_table_widget.setItem(i, 0, class_item)
        confidence_table_widget.setItem(i, 1, conf_item)

    progress_bar.setVisible(False)


# ============================================================
# 13. HISTOGRAM (Matplotlib)
# ============================================================

def plot_histogram():
    counts = [session_state["class_counts"][label] for label in labels]
    fig, ax = plt.subplots(figsize=(10, 6))
    colors = ['#3498db' if c == 0 else '#2ecc71' for c in counts]
    ax.bar(labels, counts, color=colors, edgecolor='black', linewidth=1.2)
    ax.set_title("Class Distribution Histogram", fontsize=16, fontweight='bold')
    ax.set_ylabel("Count", fontsize=12)
    ax.set_xlabel("Class", fontsize=12)
    plt.xticks(rotation=45)
    plt.tight_layout()
    return fig


# ============================================================
# 14. PIE CHART (Matplotlib)
# ============================================================

def plot_pie_chart():
    counts = [session_state["class_counts"][label] for label in labels]
    total_count = sum(counts)

    fig, ax = plt.subplots(figsize=(8, 8))
    if total_count > 0:
        colors = plt.cm.Paired(np.linspace(0, 1, len(labels)))
        wedges, texts, autotexts = ax.pie(counts, labels=labels, autopct='%1.1f%%', colors=colors, startangle=140)
        ax.set_title("Class Distribution Pie Chart", fontsize=16, fontweight='bold')
    else:
        ax.text(0.5, 0.5, 'No data available.\nPlease classify some images first.', 
                horizontalalignment='center', verticalalignment='center', fontsize=14)
        ax.set_title("Class Distribution Pie Chart", fontsize=16, fontweight='bold')
    plt.tight_layout()
    return fig


# ============================================================
# 15. ACCURACY PLOT (Matplotlib)
# ============================================================

def plot_accuracy():
    fig, ax = plt.subplots(figsize=(10, 6))
    if history is None or len(history.get("accuracy", [])) == 0:
        ax.text(0.5, 0.5, 'No training history available.', 
                horizontalalignment='center', verticalalignment='center', fontsize=14)
        ax.set_title("Training and Validation Accuracy", fontsize=16, fontweight='bold')
    else:
        epochs = range(1, len(history["accuracy"]) + 1)
        ax.plot(epochs, history["accuracy"], 'b-o', label='Training Accuracy', linewidth=2, markersize=6)
        if "val_accuracy" in history:
            ax.plot(epochs, history["val_accuracy"], 'r-s', label='Validation Accuracy', linewidth=2, markersize=6)
        ax.set_title("Training and Validation Accuracy", fontsize=16, fontweight='bold')
        ax.set_xlabel("Epochs", fontsize=12)
        ax.set_ylabel("Accuracy", fontsize=12)
        ax.set_ylim(0, 1)
        ax.legend(fontsize=12)
        ax.grid(True, linestyle='--', alpha=0.7)
    plt.tight_layout()
    return fig


# ============================================================
# 16. MODEL ARCHITECTURE (HTML for QTextBrowser)
# ============================================================

def show_model_architecture():
    html_diagram = """
    <style>
        body { font-family: 'Segoe UI', sans-serif; padding: 20px; background-color: #f8f9fa; }
        .diagram-container { display: flex; flex-direction: column; align-items: center; }
        .box { border: 2px solid #007BFF; border-radius: 10px; padding: 15px 25px; margin: 5px 0; text-align: center; font-weight: 600; background-color: #ffffff; width: 70%; max-width: 500px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }
        .arrow { font-size: 24px; color: #007BFF; margin: 2px 0; font-weight: bold; }
        .input-box { background-color: #e7f3ff; border-color: #0056b3; color: #0056b3; }
        .output-box { background-color: #e6ffed; border-color: #28a745; color: #155724; }
        .layer-box { background-color: #fff3cd; border-color: #ffc107; color: #856404; }
        .small-text { font-size: 0.85em; font-weight: normal; color: #6c757d; display: block; margin-top: 4px;}
    </style>
    <div class="diagram-container">
        <div class="box input-box">INPUT IMAGE<span class="small-text">3 × 32 × 32</span></div><div class="arrow">↓</div>
        <div class="box">CONV2D<span class="small-text">3 → 64 channels, 7×7 kernel, stride 2</span></div><div class="arrow">↓</div>
        <div class="box">BATCHNORM2D<span class="small-text">64 features</span></div><div class="arrow">↓</div>
        <div class="box">RELU<span class="small-text">Inplace Activation</span></div><div class="arrow">↓</div>
        <div class="box">MAXPOOL2D<span class="small-text">3×3 kernel, stride 2</span></div><div class="arrow">↓</div>
        <div class="box layer-box">RESNET LAYER 1<span class="small-text">2 × BasicBlock (64 channels)</span></div><div class="arrow">↓</div>
        <div class="box layer-box">RESNET LAYER 2<span class="small-text">2 × BasicBlock (128 channels)</span></div><div class="arrow">↓</div>
        <div class="box layer-box">RESNET LAYER 3<span class="small-text">2 × BasicBlock (256 channels)</span></div><div class="arrow">↓</div>
        <div class="box layer-box">RESNET LAYER 4<span class="small-text">2 × BasicBlock (512 channels)</span></div><div class="arrow">↓</div>
        <div class="box">ADAPTIVE AVG POOL<span class="small-text">Output size: 1 × 1</span></div><div class="arrow">↓</div>
        <div class="box">FLATTEN<span class="small-text">512 features</span></div><div class="arrow">↓</div>
        <div class="box">FULLY CONNECTED (LINEAR)<span class="small-text">512 → 10</span></div><div class="arrow">↓</div>
        <div class="box output-box">OUTPUT CLASSES<span class="small-text">10 CIFAR-10 Classes</span></div>
    </div>
    """
    return html_diagram


# ============================================================
# 17. MAIN WINDOW (PyQt UI Structure)
# ============================================================

class ImageClassifierApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Image Classification Using ResNet18")
        
        # Set a standard resolution that fits all screens
        self.resize(1280, 800)
        
        # Override the layout's aggressive minimum height calculation
        self.setMinimumSize(1000, 600)
        
        self.current_file_path = None

        # Apply Modern Global Stylesheet
        self.setStyleSheet("""
            QMainWindow { background-color: #f4f6f9; }
            QListWidget {
                background-color: #2c3e50; color: white; font-size: 16px; 
                border: none; padding: 10px;
            }
            QListWidget::item {
                padding: 15px; border-bottom: 1px solid #34495e;
            }
            QListWidget::item:selected {
                background-color: #3498db; border-radius: 5px;
            }
            QGroupBox {
                font-size: 16px; font-weight: bold; color: #2c3e50;
                border: 2px solid #bdc3c7; border-radius: 10px; margin-top: 15px; padding-top: 15px;
            }
            QGroupBox::title { subcontrol-origin: margin; left: 20px; padding: 0 5px; }
            QLabel { color: #2c3e50; }
            QProgressBar {
                border: 2px solid #bdc3c7; border-radius: 10px; text-align: center;
                background-color: #ecf0f1; height: 20px;
            }
            QProgressBar::chunk { background-color: #3498db; border-radius: 8px; }
            QTableWidget {
                gridline-color: #bdc3c7; background-color: white;
                font-size: 14px; border: 1px solid #bdc3c7; border-radius: 5px;
            }
            QTableWidget::item { padding: 8px; }
            QHeaderView::section {
                background-color: #34495e; color: white; padding: 10px;
                border: 1px solid #2c3e50; font-weight: bold; font-size: 14px;
            }
        """)

        # Main Layout
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QHBoxLayout(central_widget)
        main_layout.setSpacing(20)
        main_layout.setContentsMargins(10, 10, 10, 10)

        # Sidebar (Added "Exit" to the list)
        self.sidebar = QListWidget()
        self.sidebar.addItems(["Classification", "Histogram", "Pie Chart", "Accuracy Plot", "Model Architecture", "Exit"])
        self.sidebar.setMaximumWidth(220)
        self.sidebar.currentRowChanged.connect(self.change_page)
        main_layout.addWidget(self.sidebar)

        # Stacked Widget for Pages
        self.stacked_widget = QStackedWidget()
        self.stacked_widget.setStyleSheet("background-color: white; border-radius: 10px;")
        main_layout.addWidget(self.stacked_widget, 1)

        # Initialize Pages
        self.init_classification_page()
        self.init_plot_pages()
        self.init_architecture_page()

        # Set default page
        self.sidebar.setCurrentRow(0)

    # Handle Exit click from sidebar
    def change_page(self, index):
        if self.sidebar.item(index).text() == "Exit":
            self.confirm_exit()
            # Revert selection to avoid staying highlighted on "Exit"
            self.sidebar.blockSignals(True)
            self.sidebar.setCurrentRow(0)
            self.sidebar.blockSignals(False)
            return

        self.stacked_widget.setCurrentIndex(index)
        if index == 1: 
            self.histogram_canvas.figure = plot_histogram()
            self.histogram_canvas.draw()
        if index == 2: 
            self.pie_canvas.figure = plot_pie_chart()
            self.pie_canvas.draw()
        if index == 3: 
            self.accuracy_canvas.figure = plot_accuracy()
            self.accuracy_canvas.draw()

    def init_classification_page(self):
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setSpacing(20)
        layout.setContentsMargins(30, 30, 30, 30)

        # Title and Name
        title_label = QLabel("Image Classification Using ResNet18")
        title_label.setAlignment(Qt.AlignCenter)
        title_label.setStyleSheet("font-size: 28px; font-weight: bold; color: #2c3e50; margin-bottom: 5px;")
        layout.addWidget(title_label)

        name_label = QLabel("By Johanes Dian Kurniawan, S.T., M.Si.D.")
        name_label.setAlignment(Qt.AlignCenter)
        name_label.setStyleSheet("font-size: 18px; color: #7f8c8d; margin-bottom: 20px;")
        layout.addWidget(name_label)

        # --- HORIZONTAL MIDDLE SECTION ---
        middle_layout = QHBoxLayout()
        middle_layout.setSpacing(20)

        # LEFT PANEL: Image Upload & Preview
        left_group = QGroupBox("Image Upload & Preview")
        left_layout = QVBoxLayout()
        
        # 1. Image Label (Top) - Made Larger
        self.image_label = QLabel("Upload an image to begin classification")
        self.image_label.setAlignment(Qt.AlignCenter)
        self.image_label.setMinimumHeight(350) 
        self.image_label.setStyleSheet("border: 2px dashed #bdc3c7; border-radius: 10px; background-color: #f8f9fa; font-size: 16px; color: #95a5a6;")
        left_layout.addWidget(self.image_label)

        # 2. Progress Bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        left_layout.addWidget(self.progress_bar)
        
        left_group.setLayout(left_layout)
        middle_layout.addWidget(left_group, 1) # Weight 1 (50% width)

        # RIGHT PANEL: Prediction Results & Counter
        right_group = QGroupBox("Prediction Results & Counter")
        right_layout = QVBoxLayout()
        right_layout.setSpacing(10) # Reduced spacing to bring tables closer together
        
        self.prediction_label = QLabel("Classification: -")
        self.prediction_label.setAlignment(Qt.AlignCenter)
        self.prediction_label.setStyleSheet("font-size: 22px; font-weight: bold; color: #2980b9;")
        right_layout.addWidget(self.prediction_label)
        
        self.confidence_label = QLabel("Confidence: -")
        self.confidence_label.setAlignment(Qt.AlignCenter)
        self.confidence_label.setStyleSheet("font-size: 18px; color: #27ae60; font-weight: bold;")
        right_layout.addWidget(self.confidence_label)
        
        self.time_label = QLabel("Estimation Time: -")
        self.time_label.setAlignment(Qt.AlignCenter)
        self.time_label.setStyleSheet("font-size: 16px; color: #7f8c8d;")
        right_layout.addWidget(self.time_label)

        # Existing Counter Table
        self.table_widget = QTableWidget(1, 10)
        self.table_widget.setHorizontalHeaderLabels(labels)
        self.table_widget.verticalHeader().setVisible(False)
        self.table_widget.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table_widget.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        update_table_widget(self.table_widget, session_state["class_counts"])
        self.table_widget.setMinimumHeight(80)
        right_layout.addWidget(self.table_widget)

        # NEW: Confidence Table for all categories (Made Bigger)
        self.confidence_table_widget = QTableWidget(10, 2)
        self.confidence_table_widget.setHorizontalHeaderLabels(["Class", "Confidence %"])
        self.confidence_table_widget.verticalHeader().setVisible(False)
        self.confidence_table_widget.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.confidence_table_widget.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.confidence_table_widget.setMinimumHeight(350) # Increased height to make it bigger
        
        # Initialize empty confidence table
        for i, label in enumerate(labels):
            self.confidence_table_widget.setItem(i, 0, QTableWidgetItem(label))
            self.confidence_table_widget.setItem(i, 1, QTableWidgetItem("0.00%"))
            
        # Added stretch factor 1 so it expands to fill remaining space
        right_layout.addWidget(self.confidence_table_widget, 1) 
        
        right_group.setLayout(right_layout)
        middle_layout.addWidget(right_group, 1) # Weight 1 (50% width)

        layout.addLayout(middle_layout)

        # --- BOTTOM BUTTONS (Exit button removed) ---
        btn_layout = QHBoxLayout()
        btn_layout.setSpacing(20)
        
        # Upload Button moved to the left side of Classification
        self.upload_btn = QPushButton(" Upload Image")
        self.upload_btn.setStyleSheet("""
            QPushButton {
                background-color: #8e44ad; color: white; border: none;
                padding: 20px 40px; font-size: 20px; font-weight: bold; border-radius: 10px;
            }
            QPushButton:hover { background-color: #9b59b6; }
        """)
        self.upload_btn.clicked.connect(self.upload_image)

        self.classify_btn = QPushButton("🚀 Classification")
        self.classify_btn.setStyleSheet("""
            QPushButton {
                background-color: #27ae60; color: white; border: none;
                padding: 20px 40px; font-size: 20px; font-weight: bold; border-radius: 10px;
            }
            QPushButton:hover { background-color: #2ecc71; }
            QPushButton:pressed { background-color: #219150; }
        """)
        self.classify_btn.clicked.connect(self.run_classification)
        
        self.reset_btn = QPushButton("🔄 Reset Table")
        self.reset_btn.setStyleSheet("""
            QPushButton {
                background-color: #f39c12; color: white; border: none;
                padding: 20px 40px; font-size: 20px; font-weight: bold; border-radius: 10px;
            }
            QPushButton:hover { background-color: #f1c40f; }
            QPushButton:pressed { background-color: #d68910; }
        """)
        self.reset_btn.clicked.connect(self.reset_table_ui)
        
        # New button order: Upload -> Classification -> Reset
        btn_layout.addWidget(self.upload_btn)
        btn_layout.addWidget(self.classify_btn)
        btn_layout.addWidget(self.reset_btn)
        layout.addLayout(btn_layout)

        self.stacked_widget.addWidget(page)

    def init_plot_pages(self):
        # Center the plots using centered layouts
        
        # Histogram - Centered
        fig_hist = plt.figure(figsize=(10, 6))
        self.histogram_canvas = FigureCanvas(fig_hist)
        hist_container = QWidget()
        hist_layout = QHBoxLayout(hist_container)
        hist_layout.setAlignment(Qt.AlignCenter)
        hist_layout.addWidget(self.histogram_canvas)
        hist_layout.setContentsMargins(0, 0, 0, 0)
        self.stacked_widget.addWidget(hist_container)

        # Pie Chart - Centered
        fig_pie = plt.figure(figsize=(8, 8))
        self.pie_canvas = FigureCanvas(fig_pie)
        pie_container = QWidget()
        pie_layout = QHBoxLayout(pie_container)
        pie_layout.setAlignment(Qt.AlignCenter)
        pie_layout.addWidget(self.pie_canvas)
        pie_layout.setContentsMargins(0, 0, 0, 0)
        self.stacked_widget.addWidget(pie_container)

        # Accuracy Plot - Centered
        fig_acc = plt.figure(figsize=(10, 6))
        self.accuracy_canvas = FigureCanvas(fig_acc)
        acc_container = QWidget()
        acc_layout = QHBoxLayout(acc_container)
        acc_layout.setAlignment(Qt.AlignCenter)
        acc_layout.addWidget(self.accuracy_canvas)
        acc_layout.setContentsMargins(0, 0, 0, 0)
        self.stacked_widget.addWidget(acc_container)

    def init_architecture_page(self):
        from PyQt5.QtWidgets import QTextBrowser
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(30, 30, 30, 30)
        
        title_label = QLabel("Model Architecture")
        title_label.setAlignment(Qt.AlignCenter)
        title_label.setStyleSheet("font-size: 28px; font-weight: bold; color: #2c3e50; margin-bottom: 20px;")
        layout.addWidget(title_label)

        self.arch_view = QTextBrowser()
        self.arch_view.setOpenExternalLinks(True)
        self.arch_view.setHtml(show_model_architecture())
        self.arch_view.setStyleSheet("background-color: white; border: 1px solid #bdc3c7; border-radius: 10px;")
        layout.addWidget(self.arch_view)

        self.stacked_widget.addWidget(page)

    def upload_image(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Upload Image", "", "Images (*.jpg *.jpeg *.png)")
        if file_path:
            self.current_file_path = file_path
            pixmap = QPixmap(file_path)
            scaled_pixmap = pixmap.scaled(500, 500, Qt.KeepAspectRatio, Qt.SmoothTransformation)
            self.image_label.setPixmap(scaled_pixmap)
            self.image_label.setText("") 

    def run_classification(self):
        if not self.current_file_path:
            QMessageBox.warning(self, "No Image", "Please upload an image first!")
            return
        process_image(
            self.current_file_path,
            self.progress_bar,
            self.prediction_label,
            self.confidence_label,
            self.time_label,
            self.table_widget,
            self.confidence_table_widget  # Pass the new table
        )

    def reset_table_ui(self):
        reset_table()
        update_table_widget(self.table_widget, session_state["class_counts"])
        self.prediction_label.setText("Classification: -")
        self.confidence_label.setText("Confidence: -")
        self.time_label.setText("Estimation Time: -")
        
        # Reset confidence table
        for i, label in enumerate(labels):
            self.confidence_table_widget.setItem(i, 0, QTableWidgetItem(label))
            self.confidence_table_widget.setItem(i, 1, QTableWidgetItem("0.00%"))

    def confirm_exit(self):
        reply = QMessageBox.question(self, 'Exit Application', 
                                     'Are you sure you want to exit?', 
                                     QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.Yes:
            self.close()


# ============================================================
# 18. RUN APPLICATION
# ============================================================

if __name__ == "__main__":
    plt.style.use('seaborn-v0_8-darkgrid')
    
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    
    window = ImageClassifierApp()
    window.show()
    sys.exit(app.exec_())