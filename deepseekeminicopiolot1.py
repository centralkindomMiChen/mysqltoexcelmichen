import sys
import pandas as pd
import pymysql
from sqlalchemy import create_engine, text
import sqlalchemy.types
import hashlib
from PyQt5.QtWidgets import (QApplication, QMainWindow, QVBoxLayout, QHBoxLayout,
                             QWidget, QPushButton, QLabel, QLineEdit, QTextEdit,
                             QFileDialog, QGroupBox)
from PyQt5.QtCore import QTimer, Qt, QDateTime
from PyQt5.QtGui import QFont, QColor, QPalette

class ExcelToMySQLApp(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Excel导入MySQL工具 (宝石天蓝半透明 - UTF8兼容)")
        self.setGeometry(300, 300, 820, 650)
        self.set_lightblue_style()

        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.main_layout = QVBoxLayout(self.central_widget)
        self.main_layout.setSpacing(10)

        self.init_ui()

        self.timer = QTimer(self)
        self.timer.timeout.connect(self.update_time)
        self.timer.start(1000)

        self.imported_rows = 0
        self.status = "等待操作"
        self.update_status()

    def set_lightblue_style(self):
        # 半透明天蓝、宝石蓝渐变风格
        self.setStyleSheet("""
            QMainWindow {
                background-color: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                                                stop:0 rgba(180,222,255,190),
                                                stop:1 rgba(90,180,255,170));
                border: 1px solid rgba(120, 180, 240, 180);
                border-radius: 8px;
            }
            QGroupBox {
                background-color: rgba(210, 235, 255, 160);
                border: 1px solid rgba(100,180,240,200);
                border-radius: 6px;
                margin-top: 12px;
                padding: 12px;
                padding-top: 25px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                subcontrol-position: top left;
                left: 12px;
                padding: 3px 6px;
                background-color: rgba(120, 180, 240, 180);
                border-radius: 4px;
                color: #2B4C77;
                font-weight: bold;
            }
            QPushButton {
                background-color: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                                                stop:0 rgba(120,180,240,220),
                                                stop:1 rgba(90,160,220,220));
                border: 1px solid rgba(100,180,240,200);
                border-radius: 5px;
                padding: 7px 14px;
                min-width: 110px;
                color: #255A8A;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                                                stop:0 rgba(140,200,255,230),
                                                stop:1 rgba(100,180,240,230));
                border: 1.5px solid rgba(90,140,200,220);
                color: #1366bb;
            }
            QPushButton:pressed {
                background-color: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                                                stop:0 rgba(90,160,220,230),
                                                stop:1 rgba(80,130,180,230));
                border: 1px solid rgba(60,110,170,220);
            }
            QPushButton:disabled {
                background-color: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                                                stop:0 rgba(180,222,255,100),
                                                stop:1 rgba(120,180,240,100));
                border: 1px solid rgba(120,180,240,100);
                color: rgba(25, 90, 160, 100);
            }
            QLineEdit {
                background-color: rgba(255, 255, 255, 210);
                border: 1px solid rgba(120,180,240,180);
                border-radius: 4px;
                padding: 6px;
                color: #1865A0;
            }
            QLineEdit[readOnly="true"] {
                background-color: rgba(200, 230, 255, 150);
                color: #2B4C77;
                border: 1px solid rgba(120,180,240,140);
                border-radius: 4px;
                padding: 6px;
            }
            QLabel {
                color: #2B4C77;
                background-color: transparent;
                padding: 3px;
            }
            QTextEdit#ConsoleOutput {
                background-color: rgba(190, 230, 255, 140);
                border: 1px solid rgba(100,180,240,150);
                color: #255A8A;
                font-family: 'Consolas', 'Courier New', monospace;
                border-radius: 4px;
            }
        """)
        self.setWindowOpacity(0.95)
        font = QFont("微软雅黑", 10)
        self.setFont(font)

    def init_ui(self):
        self.create_control_group()
        self.create_db_info_group()
        self.create_table_info_group()
        self.create_console_group()

        self.status_bar = QLabel()
        self.status_bar.setAlignment(Qt.AlignCenter)
        self.status_bar.setStyleSheet(
            "padding: 4px; color: #255A8A; "
            "background-color: rgba(180,222,255,190); "
            "border-radius: 3px; font-weight:bold;"
        )
        self.main_layout.addWidget(self.status_bar)
        self.update_time()

    def create_control_group(self):
        group = QGroupBox("控制面板")
        layout = QHBoxLayout()
        layout.setSpacing(10)

        self.btn_select = QPushButton("选择Excel文件")
        self.btn_select.clicked.connect(self.select_excel_file)
        layout.addWidget(self.btn_select)

        self.file_path = QLineEdit()
        self.file_path.setPlaceholderText("未选择文件")
        self.file_path.setReadOnly(True)
        layout.addWidget(self.file_path, 1)

        self.btn_import = QPushButton("导入数据库")
        self.btn_import.clicked.connect(self.import_to_mysql)
        self.btn_import.setEnabled(False)
        layout.addWidget(self.btn_import)

        group.setLayout(layout)
        self.main_layout.addWidget(group)

    def create_db_info_group(self):
        group = QGroupBox("数据库信息")
        layout = QHBoxLayout()
        layout.setSpacing(8)

        lbl_db_name = QLabel("数据库名称:")
        lbl_db_name.setMinimumWidth(80)
        layout.addWidget(lbl_db_name)
        self.db_name = QLineEdit("michentestdb2")
        layout.addWidget(self.db_name, 1)

        lbl_db_user = QLabel("用户名:")
        lbl_db_user.setMinimumWidth(60)
        layout.addWidget(lbl_db_user)
        self.db_user = QLineEdit("root")
        layout.addWidget(self.db_user, 1)

        lbl_db_password = QLabel("密码:")
        lbl_db_password.setMinimumWidth(50)
        layout.addWidget(lbl_db_password)
        self.db_password = QLineEdit("123")
        self.db_password.setEchoMode(QLineEdit.Password)
        layout.addWidget(self.db_password, 1)

        group.setLayout(layout)
        self.main_layout.addWidget(group)

    def create_table_info_group(self):
        group = QGroupBox("表信息")
        layout = QHBoxLayout()
        layout.setSpacing(8)

        lbl_table_name = QLabel("表名称:")
        lbl_table_name.setMinimumWidth(65)
        layout.addWidget(lbl_table_name)
        self.table_name = QLineEdit("report_data")
        layout.addWidget(self.table_name, 1)

        lbl_rows_imported = QLabel("导入行数:")
        lbl_rows_imported.setMinimumWidth(70)
        layout.addWidget(lbl_rows_imported)
        self.rows_imported = QLineEdit("0")
        self.rows_imported.setReadOnly(True)
        self.rows_imported.setMaximumWidth(100)
        layout.addWidget(self.rows_imported, 0)

        lbl_op_status = QLabel("状态:")
        lbl_op_status.setMinimumWidth(45)
        layout.addWidget(lbl_op_status)
        self.operation_status = QLineEdit("等待操作")
        self.operation_status.setReadOnly(True)
        layout.addWidget(self.operation_status, 1)

        group.setLayout(layout)
        self.main_layout.addWidget(group)

    def create_console_group(self):
        group = QGroupBox("控制台输出")
        layout = QVBoxLayout()
        self.console_output = QTextEdit()
        self.console_output.setObjectName("ConsoleOutput")
        self.console_output.setReadOnly(True)
        layout.addWidget(self.console_output)
        group.setLayout(layout)
        self.main_layout.addWidget(group)
        self.main_layout.setStretchFactor(group, 1)

    def select_excel_file(self):
        options = QFileDialog.Options()
        file_name, _ = QFileDialog.getOpenFileName(
            self, "选择Excel文件", "",
            "Excel Files (*.xlsx *.xls);;All Files (*)",
            options=options
        )
        if file_name:
            self.file_path.setText(file_name)
            self.btn_import.setEnabled(True)
            self.log_message(f"已选择文件: {file_name}")
        else:
            self.btn_import.setEnabled(False)

    def import_to_mysql(self):
        excel_file = self.file_path.text()
        if not excel_file:
            self.log_message("错误: 请先选择Excel文件")
            return

        self.btn_import.setEnabled(False)
        self.status = "处理中..."
        self.update_status()
        QApplication.processEvents()

        conn = None
        cursor = None
        engine = None

        try:
            self.log_message("正在读取Excel文件...")
            df = pd.read_excel(excel_file)
            if df.empty:
                self.log_message("警告: Excel文件为空，跳过导入")
                self.status = "完成 (无数据)"
                return
            self.log_message(f"成功读取Excel文件，共 {len(df)} 行数据")

            config = {
                'user': self.db_user.text(),
                'password': self.db_password.text(),
                'host': 'localhost',
                'charset': 'utf8',
                'use_unicode': True
            }
            self.log_message("正在连接MySQL数据库...")
            conn = pymysql.connect(**config)
            cursor = conn.cursor()
            db_name = self.db_name.text()
            self.log_message(f"正在创建/使用数据库: {db_name}")
            cursor.execute(f"CREATE DATABASE IF NOT EXISTS `{db_name}` CHARACTER SET utf8 COLLATE utf8_general_ci")
            cursor.execute(f"USE `{db_name}`")
            table_name = self.table_name.text()

            def generate_data_fingerprint(dataframe):
                sample_data = dataframe.head(min(5, len(dataframe))).to_string(index=False)
                return hashlib.md5(sample_data.encode('utf-8')).hexdigest()

            current_fingerprint = generate_data_fingerprint(df)
            self.log_message("已生成数据指纹用于重复检查")
            cursor.execute(f"SHOW TABLES LIKE '{table_name}'")
            table_exists = cursor.fetchone() is not None

            if table_exists:
                self.log_message(f"表 {table_name} 已存在，检查重复数据...")
                try:
                    cursor.execute(f"SHOW COLUMNS FROM `{table_name}` LIKE 'data_fingerprint'")
                    has_fingerprint_column = cursor.fetchone() is not None
                    if has_fingerprint_column:
                        cursor.execute(f"SELECT 1 FROM `{table_name}` WHERE `data_fingerprint` = %s LIMIT 1", (current_fingerprint,))
                        if cursor.fetchone() is not None:
                            self.log_message("检测到相似数据批次 (基于指纹)，跳过导入")
                            self.status = "完成 (跳过重复数据)"
                            return
                    else:
                        self.log_message(f"警告: 表 {table_name} 不存在 'data_fingerprint' 列。")
                except pymysql.Error as e:
                    self.log_message(f"检查表结构时出错: {e}. 继续尝试导入...")

            engine_url = f"mysql+pymysql://{config['user']}:{config['password']}@{config['host']}/{db_name}?charset=utf8"
            engine = create_engine(engine_url)
            df['data_fingerprint'] = current_fingerprint
            custom_dtype = {'data_fingerprint': sqlalchemy.types.VARCHAR(32)}

            self.log_message("正在导入数据到MySQL...")
            df.to_sql(
                name=table_name,
                con=engine,
                if_exists='append',
                index=False,
                chunksize=1000,
                dtype=custom_dtype if not table_exists else None
            )

            if not table_exists:
                with engine.connect() as connection:
                    trans = connection.begin()
                    try:
                        inspector = sqlalchemy.inspect(engine)
                        columns_in_table = [col['name'] for col in inspector.get_columns(table_name)]
                        if 'id' not in columns_in_table:
                            connection.execute(text(f"ALTER TABLE `{table_name}` ADD COLUMN `id` INT AUTO_INCREMENT PRIMARY KEY FIRST;"))
                            self.log_message(f"Added AUTO_INCREMENT PRIMARY KEY 'id' to new table '{table_name}'.")
                        if 'import_time' not in columns_in_table:
                            connection.execute(text(f"ALTER TABLE `{table_name}` ADD COLUMN `import_time` TIMESTAMP DEFAULT CURRENT_TIMESTAMP;"))
                            self.log_message(f"Added 'import_time' column to new table '{table_name}'.")
                        trans.commit()
                    except Exception as alter_e:
                        trans.rollback()
                        self.log_message(f"添加列时出错 (可能已存在或其它问题): {alter_e}")

            self.imported_rows = len(df)
            self.log_message(f"成功导入 {self.imported_rows} 行数据到表 '{table_name}'")
            self.log_message("正在从数据库导出当前批次数据到CSV和Excel...")
            export_df = pd.read_sql(f"SELECT * FROM `{table_name}` WHERE `data_fingerprint` = '{current_fingerprint}'", engine)

            if not export_df.empty:
                csv_output = 'output_current_import.csv'
                export_df.to_csv(csv_output, index=False, encoding='utf-8-sig')
                self.log_message(f"当前导入批次数据已导出到: {csv_output}")
                excel_output = 'output_current_import.xlsx'
                export_df.to_excel(excel_output, index=False)
                self.log_message(f"当前导入批次数据已导出到: {excel_output}")
            else:
                self.log_message("没有数据导出（当前批次未找到或为空）。")
            self.status = "完成 (成功)"

        except Exception as e_outer:
            self.log_message(f"发生错误: {str(e_outer)}")
            import traceback
            self.log_message(traceback.format_exc())
            self.status = "完成 (失败)"
        finally:
            if cursor:
                cursor.close()
            if conn:
                conn.close()
            if engine:
                engine.dispose()

            self.btn_import.setEnabled(True)
            self.update_status()
            QApplication.processEvents()

    def log_message(self, message):
        timestamp = QDateTime.currentDateTime().toString("hh:mm:ss")
        self.console_output.append(f"[{timestamp}] {message}")
        self.console_output.ensureCursorVisible()

    def update_time(self):
        current_time = QDateTime.currentDateTime().toString("yyyy-MM-dd hh:mm:ss")
        self.status_bar.setText(f"系统时间: {current_time}")

    def update_status(self):
        self.rows_imported.setText(str(self.imported_rows))
        self.operation_status.setText(self.status)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    font = QFont("微软雅黑", 10)
    app.setFont(font)
    window = ExcelToMySQLApp()
    window.show()
    sys.exit(app.exec_())
