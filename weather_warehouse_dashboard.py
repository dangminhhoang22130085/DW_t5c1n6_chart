import sys
import pymysql
from datetime import datetime
import matplotlib.pyplot as plt
from matplotlib.figure import Figure
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from PyQt5.QtWidgets import *
from PyQt5.QtCore import Qt, QTimer

# Thiết lập Font cho Matplotlib
plt.rcParams['font.family'] = 'Segoe UI'
plt.rcParams['axes.unicode_minus'] = False


class ExchangeRateDWDashboard(QMainWindow):
    # 2. Khởi tạo ExchangeRateDWDashboard (__init__)
    def __init__(self):
        super().__init__()
        self.setWindowTitle("ĐỒ ÁN DATA WAREHOUSE - PHÂN TÍCH TỶ GIÁ HỐI ĐOÁI")
        self.setGeometry(100, 50, 1500, 900)
        self.setStyleSheet("background:#f4f6f9;")

        self.conn = None
        self.cb_month = QComboBox()  # Khởi tạo ComboBox Tháng
        # 3. Kết nối Database host='localhost', db='db_mart' <<Database>>
        try:
            self.conn = pymysql.connect(
                host='localhost',
                user='root',
                password='',
                database='db_mart',
                charset='utf8mb4',
                cursorclass=pymysql.cursors.DictCursor,
                autocommit=True
            )
            print("Kết nối Data Mart Tỷ giá Hối đoái thành công!")
        except Exception as e:
            # 3a. Hiển thị lỗi & Thoát ứng dụng <<Error>>
            app = QApplication.instance() or QApplication(sys.argv)
            QMessageBox.critical(None, "Lỗi", f"Không kết nối được CSDL 'db_mart':\n{e}")
            sys.exit(1)

        # 4. Khởi tạo Giao diện init_ui()
        self.init_ui()

        # 8. Load dữ liệu Filter (load_initial_filters)
        try:
            self.load_initial_filters()
        except Exception as e:
            print("Lỗi khi load filter ban đầu:", e)

        # 9. Kết nối Signals (ComboBox + Refresh)
        self.cb_year.currentIndexChanged.connect(self.load_data)
        self.cb_currency.currentIndexChanged.connect(self.load_data)
        self.cb_month.currentIndexChanged.connect(self.load_data)

        # 10. Load dữ liệu lần đầu load_data() <<LoadData>>
        self.load_data()

        # 11. Khởi động Timer Auto refresh mỗi 10 phút
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.load_data)
        self.timer.start(600000)   # 600_000 ms = 10 phút

    # 15. Đóng kết nối DB khi cửa sổ bị đóng
    def closeEvent(self, event):
        try:
            if self.conn:
                try:
                    self.conn.close()
                finally:
                    self.conn = None
        except Exception as e:
            print("Lỗi khi đóng kết nối CSDL:", e)
        event.accept()

    # 4. Khởi tạo UI
    def init_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)
        layout.setContentsMargins(15, 15, 15, 15)
        layout.setSpacing(15)

        # 5. Tạo Header (Exchange Rate Data Mart)
        header = QLabel("Exchange Rate Data Mart")
        header.setStyleSheet(
            "font-size:32px; font-weight:900; color:#1565c0; "
            "background:white; padding:20px; border-radius:12px;"
        )
        header.setAlignment(Qt.AlignCenter)
        layout.addWidget(header)

        # 6. Tạo Filter Bar (Năm, Tháng, Tiền tệ, Nút Làm mới)
        filter_bar = QHBoxLayout()
        lbl_year = QLabel("Năm:"); lbl_year.setStyleSheet("font-size:14px; font-weight:bold;")
        filter_bar.addWidget(lbl_year)
        self.cb_year = QComboBox()
        self.cb_year.addItem("Tất cả", None)
        self.cb_year.setFixedWidth(120)
        filter_bar.addWidget(self.cb_year)

        lbl_month = QLabel("Tháng:"); lbl_month.setStyleSheet("font-size:14px; font-weight:bold;")
        filter_bar.addWidget(lbl_month)
        self.cb_month.addItem("Tất cả", None)
        self.cb_month.setFixedWidth(120)
        filter_bar.addWidget(self.cb_month)

        lbl_currency = QLabel("Tiền tệ:"); lbl_currency.setStyleSheet("font-size:14px; font-weight:bold;")
        filter_bar.addWidget(lbl_currency)
        self.cb_currency = QComboBox()
        self.cb_currency.addItem("Tất cả", None)
        self.cb_currency.setFixedWidth(260)
        filter_bar.addWidget(self.cb_currency)

        btn_refresh = QPushButton("Làm mới dữ liệu")
        btn_refresh.clicked.connect(self.load_data)
        filter_bar.addWidget(btn_refresh)
        filter_bar.addStretch()
        layout.addLayout(filter_bar)

        # 7. Tạo 5 Tabs (Tổng quan, Xu hướng, So sánh, Biến động, Chất lượng)
        self.tabs = QTabWidget()
        layout.addWidget(self.tabs, 1)
        self.tab_overview = QWidget(); self.tabs.addTab(self.tab_overview, "Tổng quan")
        self.tab_trend = QWidget(); self.tabs.addTab(self.tab_trend, "Xu hướng theo tháng")
        self.tab_compare = QWidget(); self.tabs.addTab(self.tab_compare, "So sánh tiền tệ")
        self.tab_rate_range = QWidget(); self.tabs.addTab(self.tab_rate_range, "Biến động Tỷ giá")
        self.tab_quality = QWidget(); self.tabs.addTab(self.tab_quality, "Chất lượng dữ liệu")

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
        for card in self.cards: cards.addWidget(card)
        layout.addLayout(cards); layout.addStretch()

    def create_big_card(self, title, value, color):
        card = QWidget()
        card.setStyleSheet(f"QWidget {{ background:white; border-radius:15px; padding:20px; border-left:8px solid {color}; }}")
        vbox = QVBoxLayout(card)
        lbl_title = QLabel(title); lbl_title.setStyleSheet("font-size:14px; color:#555; font-weight:500;")
        vbox.addWidget(lbl_title)
        lbl_value = QLabel(value); lbl_value.setStyleSheet(f"font-size:36px; font-weight:900; color:{color};")
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
        row.addWidget(self.canvas_max); row.addWidget(self.canvas_min)
        layout.addLayout(row)

    def setup_quality(self):
        layout = QVBoxLayout(self.tab_quality)
        self.table_quality = QTableWidget()
        self.table_quality.setColumnCount(4)
        self.table_quality.setHorizontalHeaderLabels(["Mã Tiền tệ", "Tên Tiền tệ", "Số bản ghi tháng", "Nguồn (SourceId)"])
        layout.addWidget(self.table_quality)

    # 8. Load dữ liệu Filter
    def load_initial_filters(self):
        cur = self.conn.cursor()
        try:
            # Years
            cur.execute("SELECT DISTINCT year FROM fact_monthly_rate ORDER BY year DESC")
            years = cur.fetchall()
            self.cb_year.clear(); self.cb_year.addItem("Tất cả", None)
            for y in years: self.cb_year.addItem(str(y['year']), y['year'])
            # Currencies
            cur.execute("SELECT id, currencyCode, currencyName FROM mart_dim_currency ORDER BY currencyCode")
            currencies = cur.fetchall()
            self.cb_currency.clear(); self.cb_currency.addItem("Tất cả", None)
            for c in currencies: self.cb_currency.addItem(f"{c['currencyCode']} - {c['currencyName']}", c['id'])
            # Months
            self.cb_month.clear(); self.cb_month.addItem("Tất cả", None)
            for m in range(1, 13): self.cb_month.addItem(str(m), m)
        finally: cur.close()

    # 14. Load dữ liệu & cập nhật dashboard (14a → 14g)
    def load_data(self):
        try:
            if not self.conn: return
            cur = self.conn.cursor()

            # --- 14a. Lấy Filter & Xây dựng WHERE clause ---
            selected_year = self.cb_year.currentData()
            selected_month = self.cb_month.currentData()
            selected_currency_id = self.cb_currency.currentData()
            conditions = []
            if selected_year: conditions.append(f"year={selected_year}")
            if selected_month: conditions.append(f"month={selected_month}")
            if selected_currency_id: conditions.append(f"currencyId={selected_currency_id}")
            where_clause = " WHERE " + " AND ".join(conditions) if conditions else ""

            # --- 14b. Query Tổng quan & cập nhật 4 Cards ---
            # Card 1: Tổng số tiền tệ hoặc tỷ giá trung bình toàn lịch sử
            if selected_currency_id:
                cur.execute(f"SELECT ROUND(AVG(avgRate), 8) AS avg_rate_all_time FROM fact_monthly_rate WHERE currencyId={selected_currency_id}")
                avg_rate = cur.fetchone()['avg_rate_all_time']
                self.value_labels[0].setText(f"{avg_rate:,.4f}" if avg_rate else "---")
            else:
                cur.execute("SELECT COUNT(*) AS total_currencies FROM mart_dim_currency")
                total_currencies = cur.fetchone()['total_currencies']
                self.value_labels[0].setText(str(total_currencies))

            # Card 2: Tổng số bản ghi tháng (áp dụng filter)
            cur.execute(f"SELECT COUNT(*) AS total_records FROM fact_monthly_rate {where_clause}")
            self.value_labels[1].setText(f"{cur.fetchone()['total_records']:,}")

            # Card 3: Tháng cập nhật gần nhất
            if selected_month and selected_year:
                latest_month_str = f"{selected_year}-{str(selected_month).zfill(2)}"
            else:
                cur.execute(f"SELECT MAX(CONCAT(year,'-',LPAD(month,2,'0'))) AS latest_month FROM fact_monthly_rate {where_clause}")
                latest_month_str = cur.fetchone()['latest_month']
            self.value_labels[2].setText(latest_month_str if latest_month_str else "---")

            # Card 4: Tỷ giá TB gần nhất
            if latest_month_str:
                query = f"SELECT ROUND(AVG(avgRate),8) AS avg_rate FROM fact_monthly_rate WHERE CONCAT(year,'-',LPAD(month,2,'0'))='{latest_month_str}'"
                if selected_currency_id: query += f" AND currencyId={selected_currency_id}"
                cur.execute(query)
                ar = cur.fetchone()['avg_rate']
                self.value_labels[3].setText(f"{ar:,.4f}" if ar else "---")
            else:
                self.value_labels[3].setText("---")

            # --- 14c. Query Xu hướng & vẽ biểu đồ Line ---
            trend_conditions = []
            if selected_currency_id: trend_conditions.append(f"currencyId={selected_currency_id}")
            if selected_year: trend_conditions.append(f"year={selected_year}")
            trend_where = " WHERE " + " AND ".join(trend_conditions) if trend_conditions else ""
            cur.execute(f"SELECT CONCAT(year,'-',LPAD(month,2,'0')) AS ym, ROUND(AVG(avgRate),4) AS rate FROM fact_monthly_rate {trend_where} GROUP BY ym ORDER BY ym ASC")
            rows = cur.fetchall()
            self.fig_trend.clf(); ax = self.fig_trend.add_subplot(111)
            if rows:
                x = [r['ym'] for r in rows]; y = [r['rate'] for r in rows]
                ax.plot(x, y, marker='o'); ax.set_ylabel("Tỷ giá TB")
                ax.set_title("Xu hướng tỷ giá")
                ax.tick_params(axis='x', rotation=45)
            else: ax.text(0.5,0.5,"Không có dữ liệu xu hướng",ha='center',va='center')
            self.fig_trend.tight_layout(); self.canvas_trend.draw()

            # --- 14d. Query So sánh & vẽ biểu đồ Bar ---
            compare_conditions = [];
            if selected_year: compare_conditions.append(f"m.year={selected_year}")
            compare_where = " WHERE " + " AND ".join(compare_conditions) if compare_conditions else ""
            cur.execute(f"SELECT c.currencyCode, ROUND(AVG(m.avgRate),4) AS avg_rate FROM fact_monthly_rate m JOIN mart_dim_currency c ON m.currencyId=c.id {compare_where} GROUP BY c.currencyCode ORDER BY avg_rate DESC LIMIT 10")
            cmp_rows = cur.fetchall()
            self.fig_compare.clf(); axc = self.fig_compare.add_subplot(111)
            if cmp_rows:
                labels = [r['currencyCode'] for r in cmp_rows]; vals = [r['avg_rate'] for r in cmp_rows]
                axc.bar(labels, vals)
            else: axc.text(0.5,0.5,"Không có dữ liệu so sánh",ha='center',va='center')
            self.fig_compare.tight_layout(); self.canvas_compare.draw()

            # --- 14e. Query Biến động & vẽ biểu đồ Horizontal Bar ---
            where_year = f" WHERE m.year={selected_year}" if selected_year else ""
            # Max Range
            cur.execute(f"SELECT c.currencyCode, ROUND(AVG(m.maxRate)-AVG(m.minRate),4) AS avg_range FROM fact_monthly_rate m JOIN mart_dim_currency c ON m.currencyId=c.id {where_year} GROUP BY c.currencyCode ORDER BY avg_range DESC LIMIT 10")
            max_rows = cur.fetchall()
            self.fig_max_rate.clf(); axm = self.fig_max_rate.add_subplot(111)
            if max_rows:
                axm.barh([r['currencyCode'] for r in max_rows[::-1]], [r['avg_range'] for r in max_rows[::-1]], color='#E53935')
            else: axm.text(0.5,0.5,"Không có dữ liệu Max Range",ha='center',va='center')
            self.fig_max_rate.tight_layout(); self.canvas_max.draw()

            # Min Range
            cur.execute(f"SELECT c.currencyCode, ROUND(AVG(m.maxRate)-AVG(m.minRate),4) AS avg_range FROM fact_monthly_rate m JOIN mart_dim_currency c ON m.currencyId=c.id {where_year} GROUP BY c.currencyCode ORDER BY avg_range ASC LIMIT 10")
            min_rows = cur.fetchall()
            self.fig_min_rate.clf(); axn = self.fig_min_rate.add_subplot(111)
            if min_rows:
                axn.barh([r['currencyCode'] for r in min_rows], [r['avg_range'] for r in min_rows], color='#42A5F5')
            else: axn.text(0.5,0.5,"Không có dữ liệu Min Range",ha='center',va='center')
            self.fig_min_rate.tight_layout(); self.canvas_min.draw()

            # --- 14f. Query Chất lượng & cập nhật Table ---
            cur.execute("SELECT c.currencyCode,c.currencyName,COUNT(*) AS records,c.sourceId FROM fact_monthly_rate m JOIN mart_dim_currency c ON m.currencyId=c.id GROUP BY c.currencyCode,c.currencyName,c.sourceId ORDER BY records DESC")
            q_rows = cur.fetchall()
            self.table_quality.setRowCount(len(q_rows))
            for i,r in enumerate(q_rows):
                self.table_quality.setItem(i,0,QTableWidgetItem(r['currencyCode']))
                self.table_quality.setItem(i,1,QTableWidgetItem(r['currencyName']))
                self.table_quality.setItem(i,2,QTableWidgetItem(f"{r['records']:,}"))
                self.table_quality.setItem(i,3,QTableWidgetItem(str(r['sourceId'])))
            self.table_quality.resizeColumnsToContents()

            # 14g. Vẽ lại tất cả Canvas đã thực hiện ở trên

        except Exception as ex:
            print("Lỗi khi load data:", ex)
        finally:
            try: cur.close()
            except Exception: pass


if __name__ == "__main__":
    # 1. Khởi động ứng dụng
    app = QApplication(sys.argv)
    # 2. Khởi tạo ExchangeRateDWDashboard
    window = ExchangeRateDWDashboard()
    window.show()
    # 16. Kết thúc ứng dụng
    sys.exit(app.exec_())
