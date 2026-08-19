# Automating-Excel-VBA-Based-SPC-Quality-Control-Charts-in-Python
# Python Automated SPC Quality Control System
*(統計製程管制與自動異常檢測系統)*

![Python](https://img.shields.io/badge/Python-3.8%2B-blue.svg)
![License](https://img.shields.io/badge/License-MIT-green.svg)
![Status](https://img.shields.io/badge/Status-Active-brightgreen.svg)
---

## 🌟 主要功能特色 (Key Features)

* **計量型管制圖 (Variables Control Charts)**
  * **$\bar{X}-R$ Chart**：自動計算各子組平均值 (Subgroup Mean) 與全距 (Range)，監控製程中心趨勢與變異程度。
* **計數型管制圖 (Attributes Control Charts)**
  * **$p$-Chart**：監控產品不良率 (Proportion Defective)。
  * **$np$-Chart (d-Chart)**：監控固定抽樣數下的不良數 (Defective Count)。
* **自動異常檢測 (Automatic Anomaly Detection)**
  * 自動比對數據與上下管制界限（UCL / LCL）。
  * 在圖表上以顯眼的紅圈（`Out of Control`）自動圈選並標註超出管制的異常點。
* **資料彈性讀取**
  * 支援直接讀取 Excel (`.xlsx`) / CSV 檔或自動生成模擬數據進行管制測試。

---

## 🛠️ 技術棧 (Tech Stack)

* **Python 3.x**
* **Pandas** / **NumPy**：資料處理與統計量計算
* **Matplotlib**：高品質 SPC 管制圖繪製

---

## 🚀 快速開始 (Quick Start)

### 1. 安裝套件需求
```bash
pip install numpy pandas matplotlib openpyxl
