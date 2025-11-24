import sys
import pymysql
from datetime import datetime
import matplotlib.pyplot as plt
import numpy as np
from matplotlib.figure import Figure
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from PyQt5.QtWidgets import *
from PyQt5.QtCore import Qt, QTimer

# Thiết lập Font cho Matplotlib
plt.rcParams['font.family'] = 'Segoe UI'
plt.rcParams['axes.unicode_minus'] = False


class ExchangeRateDWDashboard(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("ĐỒ ÁN DATA WAREHOUSE - PHÂN TÍCH TỶ GIÁ HỐI ĐOÁI")
        self.setGeometry(100, 50, 1500, 900)
        self.setStyleSheet("background:#f4f6f9;")

        self.conn = None
        # Khởi tạo ComboBox Tháng ở đây để dễ dàng truy cập trong setup UI/Filter
        self.cb_month = QComboBox() 

        # Kết nối CSDL (Đã chuyển sang db_mart)
        try:
            self.conn = pymysql.connect(
                host='localhost',
                user='root',
                password='',       # <-- chỉnh mật khẩu nếu cần
                database='db_mart',
                charset='utf8mb4',
                cursorclass=pymysql.cursors.DictCursor,
                autocommit=True
            )
            print("Kết nối Data Mart Tỷ giá Hối đoái thành công!")
        except Exception as e:
            # Nếu chạy GUI, hiển thị hộp thoại rồi thoát
            app = QApplication.instance() or QApplication(sys.argv)
            QMessageBox.critical(None, "Lỗi", f"Không kết nối được CSDL 'db_mart':\n{e}")
            sys.exit(1)

        # UI + load
        self.init_ui()
        try:
            self.load_initial_filters()
        except Exception as e:
            print("Lỗi khi load filter ban đầu:", e)

        # Connect signals
        self.cb_year.currentIndexChanged.connect(self.load_data)
        self.cb_currency.currentIndexChanged.connect(self.load_data)
        self.cb_month.currentIndexChanged.connect(self.load_data) # <<< KẾT NỐI THÁNG

        # Load lần đầu
        self.load_data()

        # Timer lặp: auto refresh mỗi 10 phút
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.load_data)
        self.timer.start(600000)   # 600_000 ms = 10 phút

    def closeEvent(self, event):
        """Đóng kết nối CSDL khi cửa sổ ứng dụng bị đóng."""
        try:
            if self.conn:
                try:
                    self.conn.close()
                except Exception:
                    pass
                finally:
                    self.conn = None
        except Exception as e:
            print("Lỗi khi đóng kết nối CSDL:", e)
        event.accept()

    def init_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)
        layout.setContentsMargins(15, 15, 15, 15)
        layout.setSpacing(15)

        # Header
        header = QLabel("Exchange Rate Data Mart")
        header.setStyleSheet(
            "font-size:32px; font-weight:900; color:#1565c0; "
            "background:white; padding:20px; border-radius:12px;"
        )
        header.setAlignment(Qt.AlignCenter)
        layout.addWidget(header)

        # Filter bar
        filter_bar = QHBoxLayout()

        # Filter Năm
        lbl_year = QLabel("Năm:")
        lbl_year.setStyleSheet("font-size:14px; font-weight:bold;")
        filter_bar.addWidget(lbl_year)
        self.cb_year = QComboBox()
        self.cb_year.addItem("Tất cả", None)
        self.cb_year.setFixedWidth(120)
        self.cb_year.setStyleSheet("padding: 5px; border-radius: 6px; border: 1px solid #ccc;")
        filter_bar.addWidget(self.cb_year)

        # Filter Tháng <<< THÊM VÀO UI
        lbl_month = QLabel("Tháng:")
        lbl_month.setStyleSheet("font-size:14px; font-weight:bold;")
        filter_bar.addWidget(lbl_month)
        self.cb_month.addItem("Tất cả", None)
        self.cb_month.setFixedWidth(120)
        self.cb_month.setStyleSheet("padding: 5px; border-radius: 6px; border: 1px solid #ccc;")
        filter_bar.addWidget(self.cb_month)

        # Filter Tiền tệ
        lbl_currency = QLabel("Tiền tệ:")
        lbl_currency.setStyleSheet("font-size:14px; font-weight:bold;")
        filter_bar.addWidget(lbl_currency)
        self.cb_currency = QComboBox()
        self.cb_currency.addItem("Tất cả", None)
        self.cb_currency.setFixedWidth(260)
        self.cb_currency.setStyleSheet("padding: 5px; border-radius: 6px; border: 1px solid #ccc;")
        filter_bar.addWidget(self.cb_currency)

        btn_refresh = QPushButton("Làm mới dữ liệu")
        btn_refresh.clicked.connect(self.load_data)
        btn_refresh.setStyleSheet(
            "QPushButton { padding:10px 20px; background:#1E88E5; color:white; border-radius:8px; font-weight:bold; }"
            "QPushButton:hover { background:#1565c0; }"
        )
        filter_bar.addWidget(btn_refresh)
        filter_bar.addStretch()
        layout.addLayout(filter_bar)

        # Tabs
        self.tabs = QTabWidget()
        self.tabs.setStyleSheet("""
            QTabWidget::pane { border: 1px solid #c4c4c4; top:-1px; background: white; border-radius: 12px; } 
            QTabBar::tab { background: #E0E0E0; border: 1px solid #c4c4c4; padding: 10px 18px; font-size: 13px; border-top-left-radius: 8px; border-top-right-radius: 8px; }
            QTabBar::tab:selected { background: white; border-bottom-color: white; color: #1565c0; font-weight: bold; }
        """)
        layout.addWidget(self.tabs, 1)

        # Create tabs
        self.tab_overview = QWidget(); self.tabs.addTab(self.tab_overview, "Tổng quan")
        self.tab_trend = QWidget(); self.tabs.addTab(self.tab_trend, "Xu hướng theo tháng")
        self.tab_compare = QWidget(); self.tabs.addTab(self.tab_compare, "So sánh tiền tệ")
        self.tab_rate_range = QWidget(); self.tabs.addTab(self.tab_rate_range, "Biến động Tỷ giá")
        self.tab_quality = QWidget(); self.tabs.addTab(self.tab_quality, "Chất lượng dữ liệu")

        # Setup subcomponents
        self.setup_tabs()

    def setup_tabs(self):
        self.value_labels = []
        self.setup_overview()
        self.setup_trend()
        self.setup_compare()
        self.setup_rate_range()
        self.setup_quality()

    def setup_overview(self):
        layout = QVBoxLayout(self.tab_overview)
        cards = QHBoxLayout()
        self.cards = [
            self.create_big_card("TỔNG SỐ TIỀN TỆ", "0", "#1E88E5"),
            self.create_big_card("TỔNG SỐ BẢN GHI THÁNG", "0", "#43A047"),
            self.create_big_card("THÁNG CẬP NHẬT GẦN NHẤT", "---", "#FB8C00"),
            self.create_big_card("TỶ GIÁ TB GẦN NHẤT", "---", "#E53935")
        ]
        for card in self.cards:
            cards.addWidget(card)
        layout.addLayout(cards)
        layout.addStretch()

    def create_big_card(self, title, value, color):
        card = QWidget()
        card.setStyleSheet(
            f"QWidget {{ background:white; border-radius:15px; padding:20px; border-left:8px solid {color}; }}"
        )
        vbox = QVBoxLayout(card)
        lbl_title = QLabel(title)
        lbl_title.setStyleSheet("font-size:14px; color:#555; font-weight:500;")
        vbox.addWidget(lbl_title)
        lbl_value = QLabel(value)
        lbl_value.setStyleSheet(f"font-size:36px; font-weight:900; color:{color};")
        self.value_labels.append(lbl_value)
        vbox.addWidget(lbl_value)
        return card

    def setup_trend(self):
        layout = QVBoxLayout(self.tab_trend)
        self.fig_trend = Figure(figsize=(12, 6))
        self.canvas_trend = FigureCanvas(self.fig_trend)
        layout.addWidget(self.canvas_trend)

    def setup_compare(self):
        layout = QVBoxLayout(self.tab_compare)
        self.fig_compare = Figure(figsize=(12, 6))
        self.canvas_compare = FigureCanvas(self.fig_compare)
        layout.addWidget(self.canvas_compare)

    def setup_rate_range(self):
        layout = QVBoxLayout(self.tab_rate_range)
        row = QHBoxLayout()
        self.fig_max_rate = Figure(figsize=(6, 5))
        self.canvas_max = FigureCanvas(self.fig_max_rate)
        self.fig_min_rate = Figure(figsize=(6, 5))
        self.canvas_min = FigureCanvas(self.fig_min_rate)
        row.addWidget(self.canvas_max)
        row.addWidget(self.canvas_min)
        layout.addLayout(row)

    def setup_quality(self):
        layout = QVBoxLayout(self.tab_quality)
        self.table_quality = QTableWidget()
        self.table_quality.setColumnCount(4)
        self.table_quality.setHorizontalHeaderLabels(["Mã Tiền tệ", "Tên Tiền tệ", "Số bản ghi tháng", "Nguồn (SourceId)"])
        self.table_quality.setStyleSheet(
            "QTableWidget { border: 1px solid #ccc; border-radius: 8px; } QHeaderView::section { background-color: #f0f0f0; font-weight: bold; padding: 5px; }"
        )
        self.table_quality.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        layout.addWidget(self.table_quality)

    def load_initial_filters(self):
        cur = self.conn.cursor()
        try:
            # Years
            cur.execute("SELECT DISTINCT year FROM fact_monthly_rate ORDER BY year DESC")
            years = cur.fetchall()
            self.cb_year.clear(); self.cb_year.addItem("Tất cả", None)
            for y in years:
                if y and y.get('year') is not None:
                    self.cb_year.addItem(str(y['year']), y['year'])

            # Currencies
            cur.execute("SELECT id, currencyCode, currencyName FROM mart_dim_currency ORDER BY currencyCode")
            currencies = cur.fetchall()
            self.cb_currency.clear(); self.cb_currency.addItem("Tất cả", None)
            for c in currencies:
                code = c.get('currencyCode') or ''
                name = c.get('currencyName') or ''
                self.cb_currency.addItem(f"{code} - {name}", c.get('id'))

            # Months <<< LOAD THÁNG
            self.cb_month.clear(); self.cb_month.addItem("Tất cả", None)
            for month in range(1, 13):
                self.cb_month.addItem(str(month), month)
                
        finally:
            cur.close()

    def load_data(self):
        # Lấy dữ liệu và vẽ lại toàn bộ dashboard
        try:
            if not self.conn:
                print("Không có kết nối DB.")
                return
            cur = self.conn.cursor()
        except Exception as e:
            print("Lỗi khi tạo cursor:", e)
            return

        try:
            selected_year = self.cb_year.currentData()
            selected_month = self.cb_month.currentData()  # <<< Lấy dữ liệu Tháng
            selected_currency_id = self.cb_currency.currentData()
            
            # Lấy mã tiền tệ đang chọn để hiển thị trên tiêu đề
            currency_code = "Tất cả"
            if selected_currency_id:
                current_text = self.cb_currency.currentText()
                if current_text and ' - ' in current_text:
                    currency_code = current_text.split(' - ')[0]

            # Xây dựng mệnh đề WHERE chung cho các chỉ số có thể lọc (Bao gồm Năm, Tháng và Tiền tệ)
            overview_conditions = []
            if selected_year:
                overview_conditions.append(f"year = {int(selected_year)}")
            if selected_month: # <<< Đã thêm điều kiện Tháng
                overview_conditions.append(f"month = {int(selected_month)}")
            if selected_currency_id:
                overview_conditions.append(f"currencyId = {int(selected_currency_id)}")

            overview_where = " WHERE " + " AND ".join(overview_conditions) if overview_conditions else ""


            # --- 1. Tổng quan ---
            
            is_specific_time = selected_month and selected_year

            # === Cập nhật Tiêu đề Cards tùy theo filter Tiền tệ và Tháng/Năm ===
            if selected_currency_id:
                # Tiêu đề cho 1 loại tiền tệ
                self.cards[0].findChild(QLabel).setText(f"TỶ GIÁ TB {currency_code} (ALL TIME)")
                self.cards[1].findChild(QLabel).setText(f"TỔNG SỐ BẢN GHI (TRONG LỌC)")
                if is_specific_time:
                    self.cards[2].findChild(QLabel).setText(f"THỜI ĐIỂM ĐANG LỌC")
                    self.cards[3].findChild(QLabel).setText(f"TỶ GIÁ TB CỦA {selected_month}/{selected_year}")
                else:
                    self.cards[2].findChild(QLabel).setText(f"THÁNG CẬP NHẬT GẦN NHẤT")
                    self.cards[3].findChild(QLabel).setText(f"TỶ GIÁ TB GẦN NHẤT")
            else:
                # Tiêu đề cho Tất cả tiền tệ (Toàn cục)
                self.cards[0].findChild(QLabel).setText("TỔNG SỐ TIỀN TỆ")
                self.cards[1].findChild(QLabel).setText("TỔNG SỐ BẢN GHI THÁNG (GLOBAL)")
                if is_specific_time:
                    self.cards[2].findChild(QLabel).setText(f"THỜI ĐIỂM ĐANG LỌC")
                    self.cards[3].findChild(QLabel).setText(f"TỶ GIÁ TB CỦA {selected_month}/{selected_year}")
                else:
                    self.cards[2].findChild(QLabel).setText("THÁNG CẬP NHẬT GẦN NHẤT (GLOBAL)")
                    self.cards[3].findChild(QLabel).setText("TỶ GIÁ TB GLOBAL GẦN NHẤT")


            # Chỉ số 1: TỔNG SỐ TIỀN TỆ HOẶC TỶ GIÁ TB TOÀN LỊCH SỬ (ALL TIME)
            if selected_currency_id:
                # Khi lọc theo tiền tệ: Hiển thị TỶ GIÁ TRUNG BÌNH TOÀN LỊCH SỬ của nó
                cur.execute(f"""
                    SELECT ROUND(AVG(avgRate), 8) AS avg_rate_all_time
                    FROM fact_monthly_rate
                    WHERE currencyId = {int(selected_currency_id)}
                """)
                all_time_rate = cur.fetchone().get('avg_rate_all_time')
                self.value_labels[0].setText(f"{all_time_rate:,.4f}" if all_time_rate is not None else "---")
            else:
                # Khi chọn Tất cả: Hiển thị TỔNG SỐ TIỀN TỆ
                cur.execute("SELECT COUNT(*) AS total_currencies FROM mart_dim_currency")
                total_currencies = cur.fetchone().get('total_currencies', 0)
                self.value_labels[0].setText(str(total_currencies))

            # Chỉ số 2: TỔNG SỐ BẢN GHI THÁNG (Áp dụng LỌC)
            cur.execute(f"SELECT COUNT(*) AS total_records FROM fact_monthly_rate {overview_where}")
            total_records = cur.fetchone().get('total_records', 0)
            self.value_labels[1].setText(f"{total_records:,}")

            # Chỉ số 3: THÁNG CẬP NHẬT GẦN NHẤT (Áp dụng LỌC)
            if is_specific_time:
                # Nếu lọc cả tháng và năm, tháng gần nhất chính là tháng/năm đang chọn
                latest_month_str = f"{selected_year}-{str(selected_month).zfill(2)}"
            else:
                # Nếu không, tìm tháng gần nhất trong phạm vi lọc còn lại
                cur.execute(f"SELECT MAX(CONCAT(year,'-',LPAD(month,2,'0'))) AS latest_month FROM fact_monthly_rate {overview_where}")
                latest_month_result = cur.fetchone().get('latest_month')
                latest_month_str = latest_month_result if latest_month_result else None
            
            self.value_labels[2].setText(latest_month_str if latest_month_str else "---")


            # Chỉ số 4: TỶ GIÁ TB GẦN NHẤT (Áp dụng LỌC theo tháng/năm chính xác hoặc gần nhất)
            if latest_month_str:
                avg_rate_query = f"""
                    SELECT ROUND(AVG(avgRate), 8) AS avg_rate
                    FROM fact_monthly_rate
                    WHERE CONCAT(year,'-',LPAD(month,2,'0')) = '{latest_month_str}'
                    """
                
                # Cần thêm điều kiện Tiền tệ nếu có lọc tiền tệ 
                if selected_currency_id:
                    avg_rate_query += f" AND currencyId = {int(selected_currency_id)}"
                
                cur.execute(avg_rate_query)
                ar = cur.fetchone().get('avg_rate')
                self.value_labels[3].setText(f"{ar:,.4f}" if ar is not None else "---")
            else:
                 self.value_labels[3].setText("---")

            # --- 2. Xu hướng theo tháng --- 
            # Giữ logic lọc theo Năm và Tiền tệ (Không lọc theo tháng vì đây là biểu đồ xu hướng theo thời gian)
            trend_conditions = []
            if selected_currency_id:
                trend_conditions.append(f"currencyId = {int(selected_currency_id)}")
            if selected_year:
                trend_conditions.append(f"year = {int(selected_year)}")
            
            trend_where = " WHERE " + " AND ".join(trend_conditions) if trend_conditions else ""
            
            trend_sql = "SELECT CONCAT(year,'-',LPAD(month,2,'0')) AS ym, ROUND(AVG(avgRate),4) AS rate FROM fact_monthly_rate"
            trend_sql += trend_where
            trend_sql += " GROUP BY ym ORDER BY ym ASC"
            cur.execute(trend_sql)
            trend_rows = cur.fetchall()
            
            # ... (Phần vẽ biểu đồ Xu hướng)
            self.fig_trend.clf()
            ax = self.fig_trend.add_subplot(111)
            if trend_rows:
                x = [r['ym'] for r in trend_rows]
                y = [r['rate'] for r in trend_rows]
                ax.plot(x, y, marker='o', linewidth=2)
                ax.set_title(f"XU HƯỚNG TỶ GIÁ TRUNG BÌNH THEO THỜI GIAN ({currency_code})" if not selected_year else f"XU HƯỚNG TỶ GIÁ TRUNG BÌNH ({currency_code} - {selected_year})", fontsize=13, fontweight='bold')
                ax.set_ylabel("Tỷ giá Trung bình")
                if len(x) > 12:
                    step = max(1, len(x)//12)
                    ax.set_xticks(x[::step])
                ax.tick_params(axis='x', rotation=45)
                ax.grid(alpha=0.3)
            else:
                ax.text(0.5, 0.5, "Không có dữ liệu để hiển thị xu hướng", ha='center', va='center')
            self.fig_trend.tight_layout()
            self.canvas_trend.draw()

            # --- 3. So sánh tiền tệ --- (Chỉ lọc theo NĂM)
            year_display = str(selected_year) if selected_year else "Tất cả các năm"
            
            compare_sql = """
                SELECT c.currencyCode, ROUND(AVG(m.avgRate),4) AS avg_rate
                FROM fact_monthly_rate m
                JOIN mart_dim_currency c ON m.currencyId = c.id
            """
            
            compare_conditions = []
            if selected_year:
                compare_conditions.append(f"m.year = {int(selected_year)}")
            # Không lọc theo tháng ở đây vì đây là so sánh các đồng tiền, thường ở cấp độ năm
            compare_where = " WHERE " + " AND ".join(compare_conditions) if compare_conditions else ""
            
            compare_sql += compare_where
            compare_sql += " GROUP BY c.currencyCode ORDER BY avg_rate DESC LIMIT 10"
            cur.execute(compare_sql)
            cmp_rows = cur.fetchall()

            # ... (Phần vẽ biểu đồ So sánh)
            self.fig_compare.clf()
            axc = self.fig_compare.add_subplot(111)
            if cmp_rows:
                labels = [r['currencyCode'] for r in cmp_rows]
                vals = [r['avg_rate'] for r in cmp_rows]
                bars = axc.bar(labels, vals)
                axc.set_title(f"TOP 10 TỶ GIÁ TRUNG BÌNH CAO NHẤT ({year_display})", fontsize=13, fontweight='bold')
                axc.set_ylabel("Tỷ giá Trung bình")
                for b in bars:
                    axc.text(b.get_x() + b.get_width()/2, b.get_height(), f"{b.get_height():.2f}", ha='center', va='bottom', fontsize=9)
            else:
                axc.text(0.5, 0.5, "Không có dữ liệu để so sánh", ha='center', va='center')
            self.fig_compare.tight_layout()
            self.canvas_compare.draw()
            
            # --- 4. Biến động Tỷ giá --- (Giữ nguyên logic chỉ lọc theo NĂM)
            where_year = f" WHERE m.year = {int(selected_year)} " if selected_year else ""
            
            # Biến động LỚN NHẤT (Top 10 Average Range Max - Min)
            cur.execute(f"""
                SELECT 
                    c.currencyCode, 
                    ROUND(AVG(m.maxRate) - AVG(m.minRate), 4) AS avg_range
                FROM fact_monthly_rate m
                JOIN mart_dim_currency c ON m.currencyId = c.id
                {where_year}
                GROUP BY c.currencyCode
                ORDER BY avg_range DESC LIMIT 10
            """)
            max_range_rows = cur.fetchall()
            # ... (Phần vẽ biểu đồ Max Range)
            self.fig_max_rate.clf()
            axm = self.fig_max_rate.add_subplot(111)
            if max_range_rows:
                codes = [r['currencyCode'] for r in max_range_rows[::-1]]
                vals = [r['avg_range'] for r in max_range_rows[::-1]]
                axm.barh(codes, vals, color='#E53935')
                axm.set_title(f"TOP 10 ĐỘ BIẾN ĐỘNG TB LỚN NHẤT ({year_display})", fontsize=12, fontweight='bold')
                axm.set_xlabel("Khoảng Biến động TB (Max - Min)")
                for i, v in enumerate(vals):
                    axm.text(v, i, f"{v:.4f}", color='black', va='center')
            else:
                axm.text(0.5, 0.5, "Không có dữ liệu Biến động Rate", ha='center', va='center')
            self.fig_max_rate.tight_layout()
            self.canvas_max.draw()

            # Biến động NHỎ NHẤT (Top 10 Average Range Max - Min)
            cur.execute(f"""
                SELECT 
                    c.currencyCode, 
                    ROUND(AVG(m.maxRate) - AVG(m.minRate), 4) AS avg_range
                FROM fact_monthly_rate m
                JOIN mart_dim_currency c ON m.currencyId = c.id
                {where_year}
                GROUP BY c.currencyCode
                ORDER BY avg_range ASC LIMIT 10
            """)
            min_range_rows = cur.fetchall()
            # ... (Phần vẽ biểu đồ Min Range)
            self.fig_min_rate.clf()
            axn = self.fig_min_rate.add_subplot(111)
            if min_range_rows:
                codes = [r['currencyCode'] for r in min_range_rows]
                vals = [r['avg_range'] for r in min_range_rows]
                axn.barh(codes, vals, color='#42A5F5')
                axn.set_title(f"TOP 10 ĐỘ BIẾN ĐỘNG TB NHỎ NHẤT ({year_display})", fontsize=12, fontweight='bold')
                axn.set_xlabel("Khoảng Biến động TB (Max - Min)")
                for i, v in enumerate(vals):
                    axn.text(v, i, f"{v:.4f}", color='black', va='center')
            else:
                axn.text(0.5, 0.5, "Không có dữ liệu Biến động Rate", ha='center', va='center')
            self.fig_min_rate.tight_layout()
            self.canvas_min.draw()

            # --- 5. Chất lượng dữ liệu --- (Không cần thay đổi logic)
            cur.execute("""
                SELECT c.currencyCode, c.currencyName, COUNT(*) AS records, c.sourceId
                FROM fact_monthly_rate m
                JOIN mart_dim_currency c ON m.currencyId = c.id
                GROUP BY c.currencyCode, c.currencyName, c.sourceId
                ORDER BY records DESC
            """)
            quality_rows = cur.fetchall()
            self.table_quality.setRowCount(len(quality_rows))
            for i, r in enumerate(quality_rows):
                self.table_quality.setItem(i, 0, QTableWidgetItem(r['currencyCode']))
                self.table_quality.setItem(i, 1, QTableWidgetItem(r['currencyName']))
                self.table_quality.setItem(i, 2, QTableWidgetItem(f"{r['records']:,}"))
                self.table_quality.setItem(i, 3, QTableWidgetItem(str(r.get('sourceId'))))
            self.table_quality.resizeColumnsToContents()

        except Exception as ex:
            print("Lỗi khi load data:", ex)
        finally:
            try:
                cur.close()
            except Exception:
                pass


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ExchangeRateDWDashboard()
    window.show()
    sys.exit(app.exec_())