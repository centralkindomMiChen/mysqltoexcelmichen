import sys
import pandas as pd
import pymysql
from sqlalchemy import create_engine, text
import sqlalchemy.types # Not strictly needed for query/export but was part of original imports for table creation
import hashlib
from PyQt5.QtWidgets import (QApplication, QMainWindow, QVBoxLayout, QHBoxLayout,
                             QWidget, QPushButton, QLabel, QLineEdit, QTextEdit,
                             QFileDialog, QGroupBox, QDateEdit, QMessageBox, QCheckBox) 
from PyQt5.QtCore import QTimer, Qt, QDateTime, QDate
from PyQt5.QtGui import QFont, QColor, QPalette # QColor, QPalette not actively used but good for consistency
import traceback 

class ExcelToMySQLApp(QMainWindow):
    """
    Main application window for importing Excel data to MySQL and querying data.

    Provides UI for:
    - Selecting an Excel file.
    - Inputting MySQL database credentials and table name.
    - Importing data from Excel to the specified MySQL table.
    - Querying data from the MySQL table with optional date range filtering.
    - Displaying query results.
    - Exporting query results to Excel or CSV files.
    - Viewing logs of operations.
    """
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Excel导入MySQL工具 (宝石天蓝半透明 - UTF8兼容)")
        self.setGeometry(300, 300, 820, 750) 
        self.set_lightblue_style()

        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.main_layout = QVBoxLayout(self.central_widget)
        self.main_layout.setSpacing(10)

        # Attributes for data query and export
        # self.date_filter_active was removed, replaced by self.date_filter_checkbox.isChecked()
        self.current_query_results = [] # Stores raw data rows from the last successful query for export.
        self.current_query_headers = [] # Stores column headers from the last successful query for export.

        self.init_ui() 

        self.timer = QTimer(self)
        self.timer.timeout.connect(self.update_time)
        self.timer.start(1000) # For updating the status bar time

        self.imported_rows = 0 # Counter for imported rows in the current session
        self.status = "等待操作" # General status message for UI

    def set_lightblue_style(self):
        """Sets the application's visual style using QSS (Qt Style Sheets)."""
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
                margin-top: 12px; /* Provides space for the title to sit above the border */
                padding: 12px;
                padding-top: 25px; /* Ensures content padding below the title */
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
                min-width: 100px; 
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
            QLineEdit, QDateEdit { /* Shared style for QLineEdit and QDateEdit input field */
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
            QTextEdit#ConsoleOutput, QTextEdit#QueryPreviewArea { /* Shared style for console and query preview */
                background-color: rgba(190, 230, 255, 140);
                border: 1px solid rgba(100,180,240,150);
                color: #255A8A;
                font-family: 'Consolas', 'Courier New', monospace; /* Monospaced font for tabular data */
                border-radius: 4px;
            }
            QDateEdit::drop-down { /* Styling for the QDateEdit dropdown button */
                subcontrol-origin: padding;
                subcontrol-position: top right;
                width: 20px;
                border-left-width: 1px;
                border-left-color: darkgray;
                border-left-style: solid;
                border-top-right-radius: 3px;
                border-bottom-right-radius: 3px;
            }
            QDateEdit::down-arrow { /* Standard arrow icon for QDateEdit dropdown */
                image: url(:/qt-project.org/styles/commonstyle/images/standardbutton-down-arrow.png); 
            }
            QCheckBox { /* Styling for QCheckBox text */
                color: #2B4C77; 
                background-color: transparent;
                padding: 3px;
                spacing: 5px; /* Space between indicator and text */
            }
            QCheckBox::indicator { /* Styling for QCheckBox indicator box */
                width: 13px;
                height: 13px;
                border: 1px solid rgba(120,180,240,180);
                border-radius: 3px;
                background-color: rgba(255, 255, 255, 210);
            }
            QCheckBox::indicator:checked { /* Styling for checked QCheckBox indicator */
                background-color: rgba(90,180,255,200);
                image: url(:/qt-project.org/styles/commonstyle/images/standardbutton-checkbox-checked.png); /* Standard checkmark icon */
            }
            QCheckBox::indicator:disabled { /* Styling for disabled QCheckBox indicator */
                border-color: rgba(120,180,240,100);
                background-color: rgba(220,220,220,150);
            }
        """)
        self.setWindowOpacity(0.95) # Apply semi-transparency to the main window
        font = QFont("微软雅黑", 10) # Standard application font
        self.setFont(font)

    def init_ui(self):
        """
        Initializes the main user interface of the application.
        Sets up the central widget, main layout, panes for controls and output,
        and all functional groups (control, DB info, table info, data query, console).
        """
        # Main horizontal layout for left and right panes
        body_layout = QHBoxLayout()
        body_layout.setSpacing(10)

        # --- Left Pane ---
        # Contains control groups for file selection, DB info, table info, and data query.
        left_pane_widget = QWidget()
        left_layout = QVBoxLayout(left_pane_widget)
        left_layout.setSpacing(10)

        self.control_group = self.create_control_group()
        left_layout.addWidget(self.control_group)

        self.db_info_group = self.create_db_info_group()
        left_layout.addWidget(self.db_info_group)

        self.table_info_group = self.create_table_info_group()
        left_layout.addWidget(self.table_info_group)

        # Data Query Functionality Group
        self.data_query_group = QGroupBox("数据查询功能")
        query_group_main_layout = QVBoxLayout() 
        query_group_main_layout.setSpacing(10) 

        # Checkbox for enabling/disabling date range filter
        checkbox_layout = QHBoxLayout()
        self.date_filter_checkbox = QCheckBox("按日期范围查询")
        self.date_filter_checkbox.setChecked(False) # Unchecked by default
        checkbox_layout.addWidget(self.date_filter_checkbox)
        checkbox_layout.addStretch(1) 
        query_group_main_layout.addLayout(checkbox_layout)

        # Row 1: Date editors and Query button
        row1_layout = QHBoxLayout()
        row1_layout.setSpacing(8) 
        row1_layout.addWidget(QLabel("开始日期:"))
        self.start_date_edit = QDateEdit(QDate.currentDate())
        self.start_date_edit.setCalendarPopup(True)
        self.start_date_edit.setDisplayFormat("yyyy-MM-dd")
        row1_layout.addWidget(self.start_date_edit)
        
        row1_layout.addSpacing(15) 
        
        row1_layout.addWidget(QLabel("结束日期:"))
        self.end_date_edit = QDateEdit(QDate.currentDate())
        self.end_date_edit.setCalendarPopup(True)
        self.end_date_edit.setDisplayFormat("yyyy-MM-dd")
        row1_layout.addWidget(self.end_date_edit)
        row1_layout.addStretch(1) 
        self.query_button = QPushButton("查询数据")
        self.query_button.clicked.connect(self.query_data)
        row1_layout.addWidget(self.query_button)
        query_group_main_layout.addLayout(row1_layout)

        # Date editors are initially disabled, linked to the checkbox state
        self.start_date_edit.setEnabled(False)
        self.end_date_edit.setEnabled(False)
        self.date_filter_checkbox.toggled.connect(self.toggle_date_editors_enabled_state)

        # Row 2: Export path and related buttons
        row2_layout = QHBoxLayout()
        row2_layout.setSpacing(8)
        row2_layout.addWidget(QLabel("导出路径:"))
        self.export_path_edit = QLineEdit()
        self.export_path_edit.setPlaceholderText("选择导出文件路径...")
        row2_layout.addWidget(self.export_path_edit, 1) 
        self.browse_export_button = QPushButton("浏览...")
        self.browse_export_button.clicked.connect(self.browse_export_file_path)
        row2_layout.addWidget(self.browse_export_button)
        self.export_button = QPushButton("导出数据")
        self.export_button.setEnabled(False) # Initially disabled, enabled after successful query
        self.export_button.clicked.connect(self.export_queried_data)
        row2_layout.addWidget(self.export_button)
        query_group_main_layout.addLayout(row2_layout)
        
        # Label to display the number of rows fetched by the query
        self.query_row_count_label = QLabel("查询结果: 0 行")
        self.query_row_count_label.setAlignment(Qt.AlignCenter)
        query_group_main_layout.addWidget(self.query_row_count_label)
        query_group_main_layout.addStretch(1) # Pushes controls to top if space available

        self.data_query_group.setLayout(query_group_main_layout)
        left_layout.addWidget(self.data_query_group)
        
        left_layout.addStretch(1) # Pushes all groups in left pane to the top
        body_layout.addWidget(left_pane_widget, 2) # Left pane takes 2/3 of width

        # --- Right Pane ---
        # Contains data query preview area and console output.
        right_pane_widget = QWidget()
        right_layout = QVBoxLayout(right_pane_widget)
        right_layout.setSpacing(10)

        # Group for Data Query Preview
        query_preview_group = QGroupBox("数据查询预览")
        query_preview_layout = QVBoxLayout()
        self.query_preview_area = QTextEdit()
        self.query_preview_area.setObjectName("QueryPreviewArea") # For styling (monospaced font)
        self.query_preview_area.setPlaceholderText("Data Query Preview Area (待实现)")
        self.query_preview_area.setReadOnly(True)
        query_preview_layout.addWidget(self.query_preview_area)
        query_preview_group.setLayout(query_preview_layout)
        right_layout.addWidget(query_preview_group, 1) # Takes 1 part of stretch factor

        # Console Output Group
        self.console_group = self.create_console_group()
        right_layout.addWidget(self.console_group, 1) # Takes 1 part of stretch factor
        
        body_layout.addWidget(right_pane_widget, 1) # Right pane takes 1/3 of width

        # Add the body_layout (containing left and right panes) to the main_layout
        self.main_layout.addLayout(body_layout)

        # Status Bar at the bottom of the window
        self.status_bar = QLabel()
        self.status_bar.setAlignment(Qt.AlignCenter)
        self.status_bar.setStyleSheet(
            "padding: 4px; color: #255A8A; "
            "background-color: rgba(180,222,255,190); "
            "border-radius: 3px; font-weight:bold;"
        )
        self.main_layout.addWidget(self.status_bar)
        
        self.update_time() # Initialize status bar time
        self.update_status() # Initialize status fields

    def toggle_date_editors_enabled_state(self, checked):
        """
        Enables or disables the start and end date QDateEdit widgets.

        This method is connected to the `toggled` signal of the date filter QCheckBox.
        
        Args:
            checked (bool): The current checked state of the QCheckBox.
        """
        self.start_date_edit.setEnabled(checked)
        self.end_date_edit.setEnabled(checked)
        if checked:
            self.log_message("日期范围查询已启用。")
        else:
            self.log_message("日期范围查询已停用。默认查询前10条数据。")

    def query_data(self):
        """
        Queries data from the MySQL database based on UI inputs.

        If the '按日期范围查询' checkbox is checked, it queries data within the
        selected date range (inclusive) from a column named 'Date'. Otherwise,
        it fetches the top 10 rows from the table.
        Results are displayed in the query preview area, and the row count label
        is updated. The export button is enabled if results are found.
        Handles database connection errors and query execution errors.
        """
        self.log_message("开始执行数据查询...")
        self.query_preview_area.clear()
        self.current_query_results = [] # Clear previous results
        self.current_query_headers = [] # Clear previous headers
        self.export_button.setEnabled(False) # Disable export button initially

        # Retrieve database and table information from UI fields
        db_name_val = self.db_name.text()
        db_user_val = self.db_user.text()
        db_password_val = self.db_password.text()
        db_host_val = 'localhost' 
        table_name_val = self.table_name.text()

        if not table_name_val: # Validate that table name is provided
            self.log_message("错误: 表名称不能为空。")
            self.query_preview_area.setText("错误: 表名称不能为空。") 
            self.query_row_count_label.setText("查询结果: 失败")
            return

        sql_query = ""
        params = None

        # Determine query type based on checkbox state
        if self.date_filter_checkbox.isChecked():
            # Date range query
            start_date_qdate = self.start_date_edit.date()
            end_date_qdate = self.end_date_edit.date()
            
            # Validate that start date is not after end date
            if start_date_qdate > end_date_qdate:
                self.log_message("错误: 开始日期不能晚于结束日期。")
                self.query_preview_area.setText("错误: 开始日期不能晚于结束日期。")
                self.query_row_count_label.setText("查询结果: 失败")
                return

            start_date_str = start_date_qdate.toString("yyyy-MM-dd")
            end_date_str = end_date_qdate.toString("yyyy-MM-dd")
            # Assuming the date column in the SQL table is named 'Date'. Backticks are used for safety.
            sql_query = f"SELECT * FROM `{table_name_val}` WHERE `Date` BETWEEN %s AND %s"
            params = (start_date_str, end_date_str)
            self.log_message(f"日期筛选已激活。查询日期范围: {start_date_str} 到 {end_date_str}")
        else:
            # Default query: top 10 rows
            sql_query = f"SELECT * FROM `{table_name_val}` LIMIT 10"
            self.log_message("日期筛选未激活。默认查询前10条数据。")
        
        conn = None
        cursor = None
        
        try:
            # Establish database connection
            self.log_message(f"连接数据库: {db_name_val}@{db_host_val}...")
            conn = pymysql.connect(host=db_host_val, 
                                   user=db_user_val, 
                                   password=db_password_val, 
                                   database=db_name_val, 
                                   charset='utf8')
            cursor = conn.cursor()
            
            # Execute the query
            self.log_message(f"执行查询: {sql_query}" + (f" WITH params {params}" if params else ""))
            cursor.execute(sql_query, params)
            results = cursor.fetchall() # Fetch all rows
            num_rows = len(results)
            self.current_query_results = results # Store for potential export

            if num_rows > 0:
                # Get column headers from cursor description
                self.current_query_headers = [desc[0] for desc in cursor.description]
                header_str = "\t".join(self.current_query_headers) # Tab-separated headers
                self.query_preview_area.append(header_str)
                # Format and display each row
                for row in results:
                    # Convert all cell values to string, handling None values
                    row_str = "\t".join(map(lambda x: str(x) if x is not None else "", row))
                    self.query_preview_area.append(row_str)
                self.export_button.setEnabled(True) # Enable export if data is found
                self.log_message(f"查询成功，检索到 {num_rows} 行数据。")
            else:
                self.query_preview_area.setText("没有找到符合条件的数据。")
                self.log_message("查询成功，但没有找到符合条件的数据。")
                self.export_button.setEnabled(False)
            
            self.query_row_count_label.setText(f"查询结果: {num_rows} 行")

        except pymysql.Error as e: # Handle database-specific errors
            self.log_message(f"数据库查询错误: {e}")
            self.query_preview_area.setText(f"查询失败: {e}")
            self.query_row_count_label.setText("查询结果: 失败")
            self.export_button.setEnabled(False)
            self.current_query_results = [] # Clear any partial data
            self.current_query_headers = []
        except Exception as e_general: # Handle other unexpected errors
            self.log_message(f"查询过程中发生一般错误: {e_general}")
            self.log_message(traceback.format_exc()) # Log full traceback
            self.query_preview_area.setText(f"查询失败: {e_general}")
            self.query_row_count_label.setText("查询结果: 失败")
            self.export_button.setEnabled(False)
            self.current_query_results = []
            self.current_query_headers = []
        finally:
            # Ensure database resources are closed
            if cursor:
                cursor.close()
            if conn:
                conn.close()
            self.log_message("查询操作完成。")
            QApplication.processEvents() # Keep UI responsive

    def browse_export_file_path(self):
        """
        Opens a QFileDialog for the user to select or enter a file path
        for exporting queried data. Supports .xlsx, .xls, and .csv formats.
        The selected path is set to the `export_path_edit` QLineEdit.
        """
        options = QFileDialog.Options()
        # Generate a default filename based on table name and current date filter
        default_filename = f"{self.table_name.text()}_export"
        if self.date_filter_checkbox.isChecked(): 
            start_date_str = self.start_date_edit.date().toString("yyyyMMdd")
            end_date_str = self.end_date_edit.date().toString("yyyyMMdd")
            if start_date_str == end_date_str:
                default_filename += f"_{start_date_str}"
            else:
                default_filename += f"_{start_date_str}_to_{end_date_str}"
        
        file_path, selected_filter = QFileDialog.getSaveFileName(self,
                                                               "选择导出文件路径和类型",
                                                               default_filename, # Default filename suggestion
                                                               "Excel Files (*.xlsx);;Excel 97-2003 Files (*.xls);;CSV Files (*.csv);;All Files (*)",
                                                               options=options)
        if file_path:
            # Ensure the file path has an appropriate extension based on the selected filter,
            # if the user didn't type one. QFileDialog might handle this with DefaultSuffix,
            # but this provides an explicit fallback.
            if selected_filter == "Excel Files (*.xlsx)" and not file_path.lower().endswith(".xlsx"):
                file_path += ".xlsx"
            elif selected_filter == "Excel 97-2003 Files (*.xls)" and not file_path.lower().endswith(".xls"):
                file_path += ".xls"
            elif selected_filter == "CSV Files (*.csv)" and not file_path.lower().endswith(".csv"):
                file_path += ".csv"
                
            self.export_path_edit.setText(file_path)
            self.log_message(f"导出路径已选择: {file_path}")

    def export_queried_data(self):
        """
        Exports the data currently stored in `self.current_query_results`
        and `self.current_query_headers` to a file (Excel or CSV).

        The file path is taken from `self.export_path_edit`. The file type
        is determined by the file extension. Uses pandas DataFrames for export.
        Provides user feedback via QMessageBox and logs operations.
        """
        # Check if there is data to export
        if not self.current_query_results or not self.current_query_headers:
            self.log_message("没有查询结果可供导出。请先执行查询。")
            QMessageBox.warning(self, "导出错误", "没有查询结果可供导出。\n请先执行查询。")
            return

        export_file_path = self.export_path_edit.text()
        if not export_file_path: # Check if a file path has been selected
            self.log_message("错误: 请先选择导出文件的路径和名称。")
            QMessageBox.warning(self, "导出错误", "请先通过“浏览...”按钮选择导出文件的路径和名称。")
            return

        # Create a DataFrame from the stored query results
        df = pd.DataFrame(self.current_query_results, columns=self.current_query_headers)
        if df.empty: # Double-check, though prior check should cover this
            self.log_message("查询结果为空，没有数据可导出。")
            QMessageBox.information(self, "导出提示", "查询结果为空，没有数据可导出。")
            return

        try:
            # Determine file type from extension for export
            file_parts = export_file_path.lower().split('.')
            file_ext = file_parts[-1] if len(file_parts) > 1 else ""

            if file_ext == 'xlsx': # Modern Excel format
                df.to_excel(export_file_path, index=False)
                self.log_message(f"数据成功导出到 Excel (.xlsx) 文件: {export_file_path}, 共 {len(df)} 行。")
                QMessageBox.information(self, "导出成功", f"数据已成功导出到:\n{export_file_path}")
            elif file_ext == 'xls': # Older Excel format
                try:
                    df.to_excel(export_file_path, index=False, engine='xlwt') # Requires xlwt library
                    self.log_message(f"数据成功导出到 Excel (.xls) 文件: {export_file_path}, 共 {len(df)} 行。")
                    QMessageBox.information(self, "导出成功", f"数据已成功导出到:\n{export_file_path}")
                except ImportError: # Handle missing xlwt library
                    self.log_message("错误: 导出为 .xls 需要 'xlwt' 库。请安装 (pip install xlwt) 后重试。")
                    QMessageBox.critical(self, "导出错误", "导出为 .xls 文件格式需要 'xlwt' 库。\n请安装后重试。")
                    return
            elif file_ext == 'csv': # CSV format
                df.to_csv(export_file_path, index=False, encoding='utf-8-sig') # utf-8-sig for BOM
                self.log_message(f"数据成功导出到 CSV 文件: {export_file_path}, 共 {len(df)} 行。")
                QMessageBox.information(self, "导出成功", f"数据已成功导出到:\n{export_file_path}")
            else: # Unsupported or unknown extension
                self.log_message(f"错误: 不支持的文件扩展名 '{file_ext}'. 文件路径: {export_file_path}。请使用 .xlsx, .xls, 或 .csv。")
                QMessageBox.warning(self, "导出错误", f"不支持的文件扩展名 '{file_ext}'。\n请确保文件路径以 .xlsx, .xls, 或 .csv 结尾。")
                return
        except Exception as e: # Catch any other errors during file export
            self.log_message(f"导出数据时发生错误: {e}")
            self.log_message(traceback.format_exc()) # Log full traceback
            QMessageBox.critical(self, "导出失败", f"导出数据时发生错误:\n{e}")
    
    def create_control_group(self):
        """Creates and returns the QGroupBox for file selection and import controls."""
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
        return group

    def create_db_info_group(self):
        """Creates and returns the QGroupBox for database connection information inputs."""
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
        return group

    def create_table_info_group(self):
        """Creates and returns the QGroupBox for table name input and import status display."""
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
        self.rows_imported_display = QLineEdit("0") 
        self.rows_imported_display.setReadOnly(True)
        self.rows_imported_display.setMaximumWidth(100)
        layout.addWidget(self.rows_imported_display, 0)

        lbl_op_status = QLabel("状态:")
        lbl_op_status.setMinimumWidth(45)
        layout.addWidget(lbl_op_status)
        self.operation_status_display = QLineEdit("等待操作") 
        self.operation_status_display.setReadOnly(True)
        layout.addWidget(self.operation_status_display, 1)

        group.setLayout(layout)
        return group

    def create_console_group(self):
        """Creates and returns the QGroupBox for displaying log messages."""
        group = QGroupBox("控制台输出")
        layout = QVBoxLayout()
        self.console_output = QTextEdit()
        self.console_output.setObjectName("ConsoleOutput") 
        self.console_output.setReadOnly(True)
        layout.addWidget(self.console_output)
        group.setLayout(layout)
        return group

    def select_excel_file(self):
        """Opens a file dialog for selecting an Excel file and updates the UI."""
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
        """
        Imports data from the selected Excel file into the specified MySQL database and table.
        Handles database creation, table creation (with 'id' and 'import_time' columns),
        data fingerprinting for duplication checks, and export of the imported batch.
        """
        excel_file = self.file_path.text()
        if not excel_file:
            self.log_message("错误: 请先选择Excel文件")
            return

        self.btn_import.setEnabled(False)
        self.status = "处理中..."
        self.update_status()
        QApplication.processEvents() # Ensure UI updates during processing

        conn_pymysql = None 
        cursor_pymysql = None 
        engine = None # SQLAlchemy engine

        try:
            self.log_message("正在读取Excel文件...")
            df = pd.read_excel(excel_file)
            if df.empty:
                self.log_message("警告: Excel文件为空，跳过导入")
                self.status = "完成 (无数据)"
                self.update_status() 
                return
            self.log_message(f"成功读取Excel文件，共 {len(df)} 行数据")

            # Database connection configuration
            config = {
                'user': self.db_user.text(),
                'password': self.db_password.text(),
                'host': 'localhost', 
                'charset': 'utf8', 
                'use_unicode': True
            }
            self.log_message("正在连接MySQL数据库 (for import)...")
            # Initial connection using pymysql for database/table checks and creation
            conn_pymysql = pymysql.connect(user=config['user'], password=config['password'], host=config['host'], charset=config['charset'])
            cursor_pymysql = conn_pymysql.cursor()
            
            db_name_val = self.db_name.text()
            self.log_message(f"正在创建/使用数据库: {db_name_val}")
            cursor_pymysql.execute(f"CREATE DATABASE IF NOT EXISTS `{db_name_val}` CHARACTER SET utf8 COLLATE utf8_general_ci")
            cursor_pymysql.execute(f"USE `{db_name_val}`")
            table_name_val = self.table_name.text()

            # Data fingerprinting to check for duplicate batch imports
            def generate_data_fingerprint(dataframe):
                sample_data = dataframe.head(min(5, len(dataframe))).to_string(index=False)
                return hashlib.md5(sample_data.encode('utf-8')).hexdigest()

            current_fingerprint = generate_data_fingerprint(df)
            self.log_message("已生成数据指纹用于重复检查")
            cursor_pymysql.execute(f"SHOW TABLES LIKE '{table_name_val}'")
            table_exists = cursor_pymysql.fetchone() is not None

            if table_exists:
                self.log_message(f"表 {table_name_val} 已存在，检查重复数据...")
                try:
                    cursor_pymysql.execute(f"SHOW COLUMNS FROM `{table_name_val}` LIKE 'data_fingerprint'")
                    has_fingerprint_column = cursor_pymysql.fetchone() is not None
                    if has_fingerprint_column:
                        cursor_pymysql.execute(f"SELECT 1 FROM `{table_name_val}` WHERE `data_fingerprint` = %s LIMIT 1", (current_fingerprint,))
                        if cursor_pymysql.fetchone() is not None:
                            self.log_message("检测到相似数据批次 (基于指纹)，跳过导入")
                            self.status = "完成 (跳过重复数据)"
                            self.update_status() 
                            # Ensure early closure before returning
                            if cursor_pymysql and not cursor_pymysql.closed: cursor_pymysql.close()
                            if conn_pymysql and conn_pymysql.open: conn_pymysql.close()
                            return
                    else:
                        self.log_message(f"警告: 表 {table_name_val} 不存在 'data_fingerprint' 列。")
                except pymysql.Error as e:
                    self.log_message(f"检查表结构时出错: {e}. 继续尝试导入...")
            
            # Close pymysql connection, as SQLAlchemy will manage its own via engine
            if cursor_pymysql and not cursor_pymysql.closed: cursor_pymysql.close()
            if conn_pymysql and conn_pymysql.open: conn_pymysql.close()

            # Use SQLAlchemy engine for pandas to_sql and subsequent read_sql for consistency
            engine_url = f"mysql+pymysql://{config['user']}:{config['password']}@{config['host']}/{db_name_val}?charset={config['charset']}"
            engine = create_engine(engine_url)
            
            df['data_fingerprint'] = current_fingerprint # Add fingerprint column to DataFrame
            custom_dtype = {'data_fingerprint': sqlalchemy.types.VARCHAR(32)}

            self.log_message("正在导入数据到MySQL (using SQLAlchemy)...")
            df.to_sql(
                name=table_name_val,
                con=engine,
                if_exists='append', # Append data if table exists
                index=False,        # Do not write DataFrame index as a column
                chunksize=1000,     # Import data in chunks
                dtype=custom_dtype if not table_exists else None # Apply dtype only for new table creation
            )

            if not table_exists: # Add 'id' and 'import_time' if it's a new table
                with engine.connect() as connection:
                    trans = connection.begin()
                    try:
                        inspector = sqlalchemy.inspect(engine)
                        columns_in_table = [col['name'] for col in inspector.get_columns(table_name_val)]
                        if 'id' not in columns_in_table:
                            connection.execute(text(f"ALTER TABLE `{table_name_val}` ADD COLUMN `id` INT AUTO_INCREMENT PRIMARY KEY FIRST;"))
                            self.log_message(f"Added AUTO_INCREMENT PRIMARY KEY 'id' to new table '{table_name_val}'.")
                        if 'import_time' not in columns_in_table:
                            connection.execute(text(f"ALTER TABLE `{table_name_val}` ADD COLUMN `import_time` TIMESTAMP DEFAULT CURRENT_TIMESTAMP;"))
                            self.log_message(f"Added 'import_time' column to new table '{table_name_val}'.")
                        trans.commit()
                    except Exception as alter_e:
                        trans.rollback()
                        self.log_message(f"添加列时出错 (可能已存在或其它问题): {alter_e}")
            
            self.imported_rows = len(df) 
            self.log_message(f"成功导入 {self.imported_rows} 行数据到表 '{table_name_val}'")
            
            # Export the just-imported batch for verification
            self.log_message("正在从数据库导出当前批次数据到CSV和Excel (using SQLAlchemy)...")
            export_df = pd.read_sql(f"SELECT * FROM `{table_name_val}` WHERE `data_fingerprint` = '{current_fingerprint}'", engine)

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
            self.log_message(traceback.format_exc()) 
            self.status = "完成 (失败)"
        finally:
            # Clean up any remaining database resources
            if 'cursor_pymysql' in locals() and cursor_pymysql and not cursor_pymysql.closed:
                cursor_pymysql.close()
            if 'conn_pymysql' in locals() and conn_pymysql and conn_pymysql.open:
                conn_pymysql.close()
            if engine: 
                engine.dispose() 

            self.btn_import.setEnabled(True) # Re-enable import button
            self.update_status() 
            QApplication.processEvents() 

    def log_message(self, message):
        """Appends a timestamped message to the console output area."""
        timestamp = QDateTime.currentDateTime().toString("hh:mm:ss")
        self.console_output.append(f"[{timestamp}] {message}")
        self.console_output.ensureCursorVisible() # Auto-scroll to the latest message

    def update_time(self):
        """Updates the system time display in the status bar."""
        current_time = QDateTime.currentDateTime().toString("yyyy-MM-dd hh:mm:ss")
        self.status_bar.setText(f"系统时间: {current_time}")

    def update_status(self):
        """Updates the display fields for imported rows and operation status."""
        if hasattr(self, 'rows_imported_display'): 
             self.rows_imported_display.setText(str(self.imported_rows))
        if hasattr(self, 'operation_status_display'): 
             self.operation_status_display.setText(self.status)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    font = QFont("微软雅黑", 10) 
    app.setFont(font)
    window = ExcelToMySQLApp()
    window.show()
    sys.exit(app.exec_())
