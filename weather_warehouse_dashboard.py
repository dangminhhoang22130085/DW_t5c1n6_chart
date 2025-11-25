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
    def __init__(self):
        super().__init__()
        self.setWindowTitle("ĐỒ ÁN DATA WAREHOUSE - PHÂN TÍCH TỶ GIÁ HỐI ĐOÁI")
        self.setGeometry(100, 50, 1500, 900)
        self.setStyleSheet("background:#f4f6f9;")

        self.conn = None
        self.cb_month = QComboBox()  # Khởi tạo ComboBox Tháng

        # 1. KHỞI TẠO & KẾT NỐI CSDL
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
            app = QApplication.instance() or QApplication(sys.argv)
            QMessageBox.critical(None, "Lỗi", f"Không kết nối được CSDL 'db_mart':\n{e}")
            sys.exit(1)

        # 2. THIẾT LẬP GIAO DIỆN (UI)
        self.init_ui()

        # 3. TẢI BỘ LỌC BAN ĐẦU
        try:
            self.load_initial_filters()
        except Exception as e:
            print("Lỗi khi load filter ban đầu:", e)

        # 4. GÁN SỰ KIỆN (Triggers load_data)
        self.cb_year.currentIndexChanged.connect(self.load_data)
        self.cb_currency.currentIndexChanged.connect(self.load_data)
        self.cb_month.currentIndexChanged.connect(self.load_data)

        # 5. TẢI DỮ LIỆU CHÍNH LẦN ĐẦU
        self.load_data()

        # 6. CẬP NHẬT TỰ ĐỘNG (QTimer)
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.load_data)
        self.timer.start(600000)  # 600_000 ms = 10 phút

    # 10. ĐÓNG KẾT NỐI (closeEvent)
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

    def init_ui(self):
        # ... (Code thiết lập UI) ...
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)
        layout.setContentsMargins(15, 15, 15, 15)
        layout.setSpacing(15)

        header = QLabel("Exchange Rate Data Mart")
        header.setStyleSheet(
            "font-size:32px; font-weight:900; color:#1565c0; "
            "background:white; padding:20px; border-radius:12px;"
        )
        header.setAlignment(Qt.AlignCenter)
        layout.addWidget(header)

        filter_bar = QHBoxLayout()
        lbl_year = QLabel("Năm:");
        lbl_year.setStyleSheet("font-size:14px; font-weight:bold;")
        filter_bar.addWidget(lbl_year)
        self.cb_year = QComboBox()
        self.cb_year.addItem("Tất cả", None)
        self.cb_year.setFixedWidth(120)
        filter_bar.addWidget(self.cb_year)

        lbl_month = QLabel("Tháng:");
        lbl_month.setStyleSheet("font-size:14px; font-weight:bold;")
        filter_bar.addWidget(lbl_month)
        self.cb_month = QComboBox()  # Dùng self.cb_month đã khởi tạo ở __init__
        self.cb_month.addItem("Tất cả", None)
        self.cb_month.setFixedWidth(120)
        filter_bar.addWidget(self.cb_month)

        lbl_currency = QLabel("Tiền tệ:");
        lbl_currency.setStyleSheet("font-size:14px; font-weight:bold;")
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

        self.tabs = QTabWidget()
        layout.addWidget(self.tabs, 1)
        self.tab_overview = QWidget();
        self.tabs.addTab(self.tab_overview, "Tổng quan")
        self.tab_trend = QWidget();
        self.tabs.addTab(self.tab_trend, "Xu hướng theo tháng")
        self.tab_compare = QWidget();
        self.tabs.addTab(self.tab_compare, "So sánh tiền tệ")
        self.tab_rate_range = QWidget();
        self.tabs.addTab(self.tab_rate_range, "Biến động Tỷ giá")
        self.tab_quality = QWidget();
        self.tabs.addTab(self.tab_quality, "Chất lượng dữ liệu")

        self.setup_tabs()

    def setup_tabs(self):
        self.value_labels = []
        self.setup_overview()
        self.setup_trend()
        self.setup_compare()
        self.setup_rate_range()
        self.setup_quality()

    # ... (Các hàm setup_overview, setup_trend, v.v.) ...

    def setup_overview(self):
        layout = QVBoxLayout(self.tab_overview)
        cards = QHBoxLayout()
        self.cards = [
            # Thêm objectName để dễ dàng tìm kiếm Label Title trong load_data
            self.create_big_card("TỔNG SỐ TIỀN TỆ", "0", "#1E88E5", "title_card_1"),
            self.create_big_card("TỔNG SỐ BẢN GHI THÁNG", "0", "#43A047", "title_card_2"),
            self.create_big_card("THÁNG CẬP NHẬT GẦN NHẤT", "---", "#FB8C00", "title_card_3"),
            self.create_big_card("TỶ GIÁ TB GẦN NHẤT", "---", "#E53935", "title_card_4")
        ]
        for card in self.cards: cards.addWidget(card)
        layout.addLayout(cards);
        layout.addStretch()

    def create_big_card(self, title, value, color, title_object_name):
        card = QWidget()
        card.setStyleSheet(
            f"QWidget {{ background:white; border-radius:15px; padding:20px; border-left:8px solid {color}; }}")
        vbox = QVBoxLayout(card)
        lbl_title = QLabel(title);
        lbl_title.setStyleSheet("font-size:14px; color:#555; font-weight:500;")
        lbl_title.setObjectName(title_object_name)  # Đặt tên Object Name cho Title Label
        vbox.addWidget(lbl_title)
        lbl_value = QLabel(value);
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
        row.addWidget(self.canvas_max);
        row.addWidget(self.canvas_min)
        layout.addLayout(row)

    def setup_quality(self):
        layout = QVBoxLayout(self.tab_quality)
        self.table_quality = QTableWidget()
        self.table_quality.setColumnCount(4)
        self.table_quality.setHorizontalHeaderLabels(
            ["Mã Tiền tệ", "Tên Tiền tệ", "Số bản ghi tháng", "Nguồn (SourceId)"])
        layout.addWidget(self.table_quality)

    # 3. TẢI BỘ LỌC BAN ĐẦU (load_initial_filters)
    def load_initial_filters(self):
        cur = self.conn.cursor()
        try:
            # Years
            cur.execute("SELECT DISTINCT year FROM fact_monthly_rate ORDER BY year DESC")
            years = cur.fetchall()
            self.cb_year.clear();
            self.cb_year.addItem("Tất cả", None)
            for y in years: self.cb_year.addItem(str(y['year']), y['year'])
            # Currencies
            cur.execute("SELECT id, currencyCode, currencyName FROM mart_dim_currency ORDER BY currencyCode")
            currencies = cur.fetchall()
            self.cb_currency.clear();
            self.cb_currency.addItem("Tất cả", None)
            for c in currencies: self.cb_currency.addItem(f"{c['currencyCode']} - {c['currencyName']}", c['id'])
            # Months
            self.cb_month.clear();
            self.cb_month.addItem("Tất cả", None)
            for m in range(1, 13): self.cb_month.addItem(str(m), m)
        finally:
            cur.close()

    # 7. TRUY VẤN DỮ LIỆU & CẬP NHẬT UI (load_data)
    def load_data(self):
        try:
            if not self.conn: return
            cur = self.conn.cursor()

            selected_year = self.cb_year.currentData()
            selected_month = self.cb_month.currentData()
            selected_currency_id = self.cb_currency.currentData()
            conditions = []
            if selected_year: conditions.append(f"year={selected_year}")
            if selected_month: conditions.append(f"month={selected_month}")
            if selected_currency_id: conditions.append(f"currencyId={selected_currency_id}")
            where_clause = " WHERE " + " AND ".join(conditions) if conditions else ""

            # Lấy QLabel Title của Card 1 để cập nhật tên
            lbl_title_card_1 = self.cards[0].findChild(QLabel, "title_card_1")

            # 7a. CẬP NHẬT CARDS TỔNG QUAN
            # Card 1: Tổng số tiền tệ HOẶC Tỷ giá TB toàn lịch sử
            if selected_currency_id:
                # Nếu chọn 1 tiền tệ: hiển thị Tỷ giá TB toàn lịch sử của tiền tệ đó
                cur.execute(
                    f"SELECT ROUND(AVG(avgRate), 8) AS avg_rate_all_time FROM fact_monthly_rate WHERE currencyId={selected_currency_id}")
                avg_rate = cur.fetchone()['avg_rate_all_time']
                self.value_labels[0].setText(f"{avg_rate:,.4f}" if avg_rate else "---")
                if lbl_title_card_1: lbl_title_card_1.setText("TỶ GIÁ TB TỔNG THỂ")
            else:
                # Nếu chọn 'Tất cả': hiển thị Tổng số tiền tệ
                cur.execute("SELECT COUNT(*) AS total_currencies FROM mart_dim_currency")
                total_currencies = cur.fetchone()['total_currencies']
                self.value_labels[0].setText(str(total_currencies))
                if lbl_title_card_1: lbl_title_card_1.setText("TỔNG SỐ TIỀN TỆ")

            # Card 2: Tổng số bản ghi tháng (áp dụng filter)
            cur.execute(f"SELECT COUNT(*) AS total_records FROM fact_monthly_rate {where_clause}")
            self.value_labels[1].setText(f"{cur.fetchone()['total_records']:,}")

            # Card 3 & 4: Tháng/Tỷ giá TB gần nhất
            if selected_month and selected_year:
                latest_month_str = f"{selected_year}-{str(selected_month).zfill(2)}"
            else:
                cur.execute(
                    f"SELECT MAX(CONCAT(year,'-',LPAD(month,2,'0'))) AS latest_month FROM fact_monthly_rate {where_clause}")
                latest_month_str = cur.fetchone()['latest_month']
            self.value_labels[2].setText(latest_month_str if latest_month_str else "---")

            if latest_month_str:
                query = f"SELECT ROUND(AVG(avgRate),8) AS avg_rate FROM fact_monthly_rate WHERE CONCAT(year,'-',LPAD(month,2,'0'))='{latest_month_str}'"
                if selected_currency_id: query += f" AND currencyId={selected_currency_id}"
                cur.execute(query)
                ar = cur.fetchone()['avg_rate']
                self.value_labels[3].setText(f"{ar:,.4f}" if ar else "---")
            else:
                self.value_labels[3].setText("---")

            # 7b. CẬP NHẬT BIỂU ĐỒ XU HƯỚNG
            self.fig_trend.clf();
            ax = self.fig_trend.add_subplot(111)

            if selected_currency_id:
                # Code vẽ biểu đồ Xu hướng cho 1 tiền tệ
                trend_conditions = []
                trend_conditions.append(f"currencyId={selected_currency_id}")
                if selected_year: trend_conditions.append(f"year={selected_year}")
                trend_where = " WHERE " + " AND ".join(trend_conditions) if trend_conditions else ""
                cur.execute(
                    f"SELECT CONCAT(year,'-',LPAD(month,2,'0')) AS ym, ROUND(avgRate,4) AS rate FROM fact_monthly_rate {trend_where} ORDER BY ym ASC")
                rows = cur.fetchall()

                if rows:
                    x = [r['ym'] for r in rows];
                    y = [r['rate'] for r in rows]
                    currency_code_name = self.cb_currency.currentText()
                    ax.plot(x, y, marker='o', linestyle='-', color='#1E88E5')
                    ax.set_ylabel("Tỷ giá TB")
                    ax.set_xlabel("Thời gian (Tháng/Năm)")
                    ax.set_title(f"XU HƯỚNG TỶ GIÁ CỦA {currency_code_name} THEO THỜI GIAN", fontsize=14,
                                 fontweight='bold')
                    ax.tick_params(axis='x', rotation=45)
                    ax.grid(axis='y', linestyle='--')
                else:
                    ax.text(0.5, 0.5, "Không có dữ liệu xu hướng", ha='center', va='center', fontsize=12, color='gray')
            else:
                # Hiển thị thông báo nếu người dùng không chọn tiền tệ
                ax.text(0.5, 0.5, "VUI LÒNG CHỌN MỘT LOẠI TIỀN TỆ ĐỂ XEM XU HƯỚNG", ha='center', va='center',
                        fontsize=14, color='darkred')
                ax.set_xticks([]);
                ax.set_yticks([])  # Ẩn trục

            self.fig_trend.tight_layout();
            self.canvas_trend.draw()  # 8. VẼ LẠI CANVAS

            # 7c. CẬP NHẬT BIỂU ĐỒ SO SÁNH (TOP 10 TỶ GIÁ TB)
            compare_conditions = [];
            if selected_year: compare_conditions.append(f"m.year={selected_year}")
            if selected_month: compare_conditions.append(f"m.month={selected_month}")  # <-- ÁP DỤNG FILTER THÁNG
            compare_where = " WHERE " + " AND ".join(compare_conditions) if compare_conditions else ""

            cur.execute(
                f"SELECT c.currencyCode, ROUND(AVG(m.avgRate),4) AS avg_rate FROM fact_monthly_rate m JOIN mart_dim_currency c ON m.currencyId=c.id {compare_where} GROUP BY c.currencyCode ORDER BY avg_rate DESC LIMIT 10")
            cmp_rows = cur.fetchall()
            self.fig_compare.clf();
            axc = self.fig_compare.add_subplot(111)

            if cmp_rows:
                labels = [r['currencyCode'] for r in cmp_rows];
                vals = [r['avg_rate'] for r in cmp_rows]
                axc.bar(labels, vals, color='#43A047')
                axc.set_ylabel("Tỷ giá TB")

                # Cập nhật tiêu đề biểu đồ so sánh
                title_suffix = ""
                if selected_month and selected_year:
                    title_suffix = f"THÁNG {selected_month}/{selected_year}"
                elif selected_year:
                    title_suffix = f"NĂM {selected_year}"

                axc.set_title(f"TOP 10 TIỀN TỆ CÓ TỶ GIÁ TB CAO NHẤT {title_suffix}", fontsize=14, fontweight='bold')
                axc.grid(axis='y', linestyle='--')
            else:
                axc.text(0.5, 0.5, "Không có dữ liệu so sánh", ha='center', va='center', fontsize=12, color='gray')
            self.fig_compare.tight_layout();
            self.canvas_compare.draw()  # 8. VẼ LẠI CANVAS

            # 7d. CẬP NHẬT BIỂU ĐỒ BIẾN ĐỘNG TỶ GIÁ (RANGE)
            where_year = f" WHERE m.year={selected_year}" if selected_year else ""

            # Max Range
            cur.execute(
                f"SELECT c.currencyCode, ROUND(AVG(m.maxRate)-AVG(m.minRate),4) AS avg_range FROM fact_monthly_rate m JOIN mart_dim_currency c ON m.currencyId=c.id {where_year} GROUP BY c.currencyCode ORDER BY avg_range DESC LIMIT 10")
            max_rows = cur.fetchall()
            self.fig_max_rate.clf();
            axm = self.fig_max_rate.add_subplot(111)
            if max_rows:
                axm.barh([r['currencyCode'] for r in max_rows[::-1]], [r['avg_range'] for r in max_rows[::-1]],
                         color='#E53935')
                axm.set_title(f"TOP 10 BIẾN ĐỘNG (RANGE) TỶ GIÁ LỚN NHẤT {selected_year or ''}", fontsize=12,
                              fontweight='bold')
                axm.set_xlabel("Biến động TB (Max Rate - Min Rate)")
                axm.grid(axis='x', linestyle='--')
            else:
                axm.text(0.5, 0.5, "Không có dữ liệu Max Range", ha='center', va='center', fontsize=12, color='gray')
            self.fig_max_rate.tight_layout();
            self.canvas_max.draw()  # 8. VẼ LẠI CANVAS

            # Min Range
            cur.execute(
                f"SELECT c.currencyCode, ROUND(AVG(m.maxRate)-AVG(m.minRate),4) AS avg_range FROM fact_monthly_rate m JOIN mart_dim_currency c ON m.currencyId=c.id {where_year} GROUP BY c.currencyCode ORDER BY avg_range ASC LIMIT 10")
            min_rows = cur.fetchall()
            self.fig_min_rate.clf();
            axn = self.fig_min_rate.add_subplot(111)
            if min_rows:
                axn.barh([r['currencyCode'] for r in min_rows], [r['avg_range'] for r in min_rows], color='#42A5F5')
                axn.set_title(f"TOP 10 BIẾN ĐỘNG (RANGE) TỶ GIÁ NHỎ NHẤT {selected_year or ''}", fontsize=12,
                              fontweight='bold')
                axn.set_xlabel("Biến động TB (Max Rate - Min Rate)")
                axn.grid(axis='x', linestyle='--')
            else:
                axn.text(0.5, 0.5, "Không có dữ liệu Min Range", ha='center', va='center', fontsize=12, color='gray')
            self.fig_min_rate.tight_layout();
            self.canvas_min.draw()  # 8. VẼ LẠI CANVAS

            # 7e. CẬP NHẬT BẢNG CHẤT LƯỢNG DỮ LIỆU
            cur.execute(
                "SELECT c.currencyCode,c.currencyName,COUNT(*) AS records,c.sourceId FROM fact_monthly_rate m JOIN mart_dim_currency c ON m.currencyId=c.id GROUP BY c.currencyCode,c.currencyName,c.sourceId ORDER BY records DESC")
            q_rows = cur.fetchall()
            self.table_quality.setRowCount(len(q_rows))
            for i, r in enumerate(q_rows):
                self.table_quality.setItem(i, 0, QTableWidgetItem(r['currencyCode']))
                self.table_quality.setItem(i, 1, QTableWidgetItem(r['currencyName']))
                self.table_quality.setItem(i, 2, QTableWidgetItem(f"{r['records']:,}"))
                self.table_quality.setItem(i, 3, QTableWidgetItem(str(r['sourceId'])))
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
    # 9. kiểm tra đóng ứng dụng chưa
    sys.exit(app.exec_())