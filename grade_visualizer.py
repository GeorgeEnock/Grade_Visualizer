import sys
import os
import re
import requests
import webbrowser
import datetime
import pandas as pd
import numpy as np
import matplotlib
import tempfile  # Added for safe temp file handling

# Use Agg backend to prevent Matplotlib from interfering with PySide6 event loop
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.colors
from matplotlib.backends.backend_pdf import PdfPages

from PySide6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel,
                               QLineEdit, QPushButton, QMessageBox, QFileDialog, QComboBox, QCheckBox,
                               QTableWidget, QTableWidgetItem, QHeaderView, QAbstractItemView, QProgressBar, QScrollArea)
from PySide6.QtCore import Qt, QThread, Signal, QTimer, QRectF, QSettings
from PySide6.QtGui import QFont, QColor, QPainter, QPen, QPalette, QIcon

class DataWorker(QThread):
    """
    Worker thread to handle file downloading and data processing
    to keep the GUI responsive.
    """
    status_update = Signal(str)
    progress_update = Signal(int)
    finished = Signal(bool, str)
    data_ready = Signal(list)

    def __init__(self, input_source, chart_type, pie_bins_text, sheet_num, chart_color, sort_option, process_all_sheets, raw_data=None, task_type="file_to_pdf", output_filename="grades_report.pdf"):
        super().__init__()
        self.input_source = input_source
        self.chart_type = chart_type
        self.pie_bins_text = pie_bins_text
        self.sheet_num = sheet_num
        self.chart_color = chart_color
        self.sort_option = sort_option
        self.process_all_sheets = process_all_sheets
        self.raw_data = raw_data
        self.task_type = task_type
        
        # Create Grade Visualizer folder in Documents
        docs_dir = os.path.join(os.path.expanduser('~'), 'Documents')
        save_dir = os.path.join(docs_dir, 'Grade Visualizer')
        os.makedirs(save_dir, exist_ok=True)
        
        # Add timestamp to filename to ensure uniqueness
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        name, ext = os.path.splitext(output_filename)
        self.output_filename = os.path.join(save_dir, f"{name}_{timestamp}{ext}")

    def run(self):
        try:
            self.progress_update.emit(5)
            # --- MODE 1: Generate PDF from UI Data (Table) ---
            if self.task_type == "data_to_pdf":
                self.status_update.emit("Generating PDF from Table Data...")
                self.progress_update.emit(10)
                df = pd.DataFrame(self.raw_data)
                
                # Clean and Validate
                df['score'] = pd.to_numeric(df['score'], errors='coerce')
                if 'sn' in df.columns:
                    df['sn'] = pd.to_numeric(df['sn'], errors='coerce')
                
                df = df.dropna(subset=['score'])
                self.progress_update.emit(40)
                
                # Apply Sorting
                df = self.apply_sorting(df)
                self.progress_update.emit(60)
                
                self.create_single_pdf(df, self.output_filename, "Custom Data")
                self.progress_update.emit(90)
                self.status_update.emit("PDF saved successfully")
                self.finished.emit(True, f"Report saved as: {self.output_filename}")
                self.progress_update.emit(100)
                return

            # --- MODE 2 & 3: File Operations ---
            # Check if input is a local file
            if os.path.isfile(self.input_source):
                temp_filename = self.input_source
                is_temp = False
                self.progress_update.emit(10)
            else:
                self.status_update.emit("Extracting File ID...")
                self.progress_update.emit(10)
                file_id = self.extract_file_id(self.input_source)
                
                if not file_id:
                    raise ValueError("Could not parse Google Drive File ID from link.")

                # Construct export URL for Excel format
                download_url = f"https://docs.google.com/spreadsheets/d/{file_id}/export?format=xlsx"
                # Use system temp directory to avoid permission issues and conflicts
                temp_filename = os.path.join(tempfile.gettempdir(), f"temp_grades_{file_id}.xlsx")
                is_temp = True

                self.status_update.emit("Downloading Excel file...")
                # Add headers and timeout to prevent firewall blocks and hanging
                headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'}
                self.progress_update.emit(20)
                response = requests.get(download_url, headers=headers, timeout=30)
                
                if response.status_code != 200:
                    raise ConnectionError(f"Download failed (Status: {response.status_code}). Check link permissions.")

                with open(temp_filename, 'wb') as f:
                    f.write(response.content)
                self.progress_update.emit(40)

            # --- MODE 2: Load Data to UI ---
            if self.task_type == "file_to_data":
                self.status_update.emit("Parsing Data...")
                self.progress_update.emit(50)
                data = self.extract_single_sheet_data(temp_filename)
                self.progress_update.emit(90)
                self.data_ready.emit(data)
                self.finished.emit(True, "Data loaded successfully.")
                self.progress_update.emit(100)
            
            # --- MODE 3: File to PDF (Direct) ---
            else:
                self.status_update.emit("Processing Data & Generating PDF...")
                self.progress_update.emit(50)
                self.generate_pdf_from_file(temp_filename, self.output_filename)
                self.progress_update.emit(90)
                self.status_update.emit("PDF saved successfully")
                self.finished.emit(True, f"Report saved as: {self.output_filename}")
                self.progress_update.emit(100)

            # Clean up temporary file
            if is_temp and os.path.exists(temp_filename):
                try:
                    os.remove(temp_filename)
                except OSError:
                    pass  # File might be locked, ignore

        except PermissionError:
            self.progress_update.emit(0)
            self.finished.emit(False, f"Permission Denied: Could not write to file.\nPlease close '{os.path.basename(self.output_filename)}' if it is open.")
        except requests.exceptions.RequestException:
            self.progress_update.emit(0)
            self.finished.emit(False, "Network Error: Connection timed out or failed. Check internet/firewall.")
        except Exception as e:
            self.progress_update.emit(0)
            self.finished.emit(False, str(e))

    def extract_file_id(self, url):
        """Extracts the file ID from a standard Google Drive URL."""
        # Regex to find the ID between /d/ and /
        match = re.search(r'/d/([a-zA-Z0-9-_]+)', url)
        return match.group(1) if match else None

    def apply_sorting(self, df):
        if "S/N" in self.sort_option:
            if 'sn' in df.columns:
                ascending = "Ascending" in self.sort_option
                df = df.sort_values(by='sn', ascending=ascending, na_position='last')
        elif "Score" in self.sort_option:
            ascending = "Ascending" in self.sort_option
            df = df.sort_values(by='score', ascending=ascending, na_position='last')
        return df

    def extract_single_sheet_data(self, input_file):
        """Extracts data from the selected sheet for the UI table."""
        # Use context manager to ensure file handle is closed (prevents file locking on Windows)
        with pd.ExcelFile(input_file) as xl:
            # Handle Sheet Selection
            try:
                sheet_idx = int(self.sheet_num) - 1
            except ValueError:
                sheet_idx = 0
            if sheet_idx < 0 or sheet_idx >= len(xl.sheet_names):
                sheet_idx = 0
                
            df_raw = xl.parse(sheet_idx, header=None)
            return self.parse_raw_sheet_content(df_raw)

    def parse_raw_sheet_content(self, df_raw):
        """Parses raw dataframe to find S/N, Name, Marks, Reg No, and Exam No."""
        grade_names = ['score', 'marks', 'grade', 'points', 'credit']
        sn_keywords = ['s/n', 's/no', 'no.', 'no', 'serial', 'seq', 'id']
        name_keywords = ['name', 'student', 'candidate', 'full name']
        reg_keywords = ['reg', 'registration', 'reg no', 'reg. no', 'admission']
        exam_keywords = ['exam', 'examination', 'exam no', 'exam. no']
        
        header_configs = []
        
        # Scan for headers
        for idx, row in df_raw.iterrows():
            row_vals = [str(v).lower().strip() for v in row.values]
            sn_idx = -1
            grade_idx = -1
            name_idx = -1
            reg_idx = -1
            exam_idx = -1
            
            for c_i, val in enumerate(row_vals):
                if 0 < len(val) < 30:
                    if sn_idx == -1 and (val in sn_keywords or any(val.startswith(k) for k in sn_keywords)):
                        sn_idx = c_i
                    elif grade_idx == -1 and any(k in val for k in grade_names):
                        grade_idx = c_i
                    elif name_idx == -1 and any(k in val for k in name_keywords):
                        name_idx = c_i
                    elif reg_idx == -1 and any(k in val for k in reg_keywords):
                        reg_idx = c_i
                    elif exam_idx == -1 and any(k in val for k in exam_keywords):
                        exam_idx = c_i
            
            if sn_idx != -1 and grade_idx != -1:
                header_configs.append((idx, sn_idx, grade_idx, name_idx, reg_idx, exam_idx))
        
        collected_data = []
        if header_configs:
            for k in range(len(header_configs)):
                start_row, sn_c, grade_c, name_c, reg_c, exam_c = header_configs[k]
                if k < len(header_configs) - 1:
                    end_row = header_configs[k+1][0]
                else:
                    end_row = len(df_raw)
                
                for r in range(start_row + 1, end_row):
                    if r >= len(df_raw): break
                    row = df_raw.iloc[r]
                    try:
                        s_val = row.iloc[sn_c]
                        g_val = row.iloc[grade_c]
                        
                        # Validate numeric S/N and Grade
                        s_num = float(s_val)
                        g_num = float(g_val)
                        
                        name_val = ""
                        if name_c != -1:
                            name_val = str(row.iloc[name_c]).strip()
                            if name_val.lower() == 'nan': name_val = ""
                        
                        reg_val = ""
                        if reg_c != -1:
                            reg_val = str(row.iloc[reg_c]).strip()
                            if reg_val.lower() == 'nan': reg_val = ""

                        exam_val = ""
                        if exam_c != -1:
                            exam_val = str(row.iloc[exam_c]).strip()
                            if exam_val.lower() == 'nan': exam_val = ""
                        
                        collected_data.append({
                            'sn': s_num, 
                            'name': name_val, 
                            'score': g_num,
                            'reg_no': reg_val,
                            'exam_no': exam_val
                        })
                    except (ValueError, TypeError):
                        continue
        
        # Fallback to simple parsing if multi-block failed
        if not collected_data:
            # Try standard read
            df = df_raw.copy()
            # Simple heuristic: find header row
            for i in range(min(10, len(df))):
                row_vals = [str(v).lower() for v in df.iloc[i]]
                if any(k in v for v in row_vals for k in grade_names):
                    df.columns = df.iloc[i]
                    df = df[i+1:]
                    break
            
            # Find cols
            g_col = next((c for c in df.columns if any(k in str(c).lower() for k in grade_names)), None)
            s_col = next((c for c in df.columns if any(k in str(c).lower() for k in sn_keywords)), None)
            n_col = next((c for c in df.columns if any(k in str(c).lower() for k in name_keywords)), None)
            r_col = next((c for c in df.columns if any(k in str(c).lower() for k in reg_keywords)), None)
            e_col = next((c for c in df.columns if any(k in str(c).lower() for k in exam_keywords)), None)
            
            if g_col:
                for _, row in df.iterrows():
                    try:
                        g_num = float(row[g_col])
                        s_num = float(row[s_col]) if s_col else 0
                        n_val = str(row[n_col]) if n_col else ""
                        r_val = str(row[r_col]) if r_col else ""
                        e_val = str(row[e_col]) if e_col else ""
                        
                        if n_val.lower() == 'nan': n_val = ""
                        if r_val.lower() == 'nan': r_val = ""
                        if e_val.lower() == 'nan': e_val = ""

                        collected_data.append({
                            'sn': s_num, 
                            'name': n_val, 
                            'score': g_num,
                            'reg_no': r_val,
                            'exam_no': e_val
                        })
                    except: pass

        return collected_data

    def create_single_pdf(self, df, output_file, sheet_name):
        with PdfPages(output_file) as pdf:
            self.add_report_page(pdf, df['score'], sheet_name, len(df))

    def generate_pdf_from_file(self, input_file, output_file):
        """Reads the Excel file and creates the visualization PDF."""
        
        def find_col(dataframe, candidates):
            # Normalize column names
            dataframe.columns = [str(c).strip() for c in dataframe.columns]
            # Exact match
            for col in dataframe.columns:
                if col.lower() in candidates:
                    return col
            # Substring match
            for col in dataframe.columns:
                if any(name in col.lower() for name in candidates):
                    return col
            return None

        xl = pd.ExcelFile(input_file)
        try:
            if self.process_all_sheets:
                sheet_indices = range(len(xl.sheet_names))
            else:
                # Handle Sheet Selection (User input is 1-based)
                try:
                    sheet_idx = int(self.sheet_num) - 1
                except ValueError:
                    sheet_idx = 0
                
                if sheet_idx < 0:
                    sheet_idx = 0

                total_sheets = len(xl.sheet_names)
                if sheet_idx >= total_sheets:
                    raise ValueError(f"Sheet not found. The only sheet found is sheet number {total_sheets}")
                sheet_indices = [sheet_idx]

            # Plotting
            with PdfPages(output_file) as pdf:
                for sheet_idx in sheet_indices:
                    sheet_name = xl.sheet_names[sheet_idx]
                    try:
                        # Read raw data for multi-block detection
                        df_raw = xl.parse(sheet_idx, header=None)
                        
                        grade_names = ['score', 'marks', 'grade', 'points', 'credit']
                        sn_keywords = ['s/n', 's/no', 'no.', 'no', 'serial', 'seq', 'id']
                        
                        scores = pd.Series(dtype=float)
                        total_students_count = 0
                        
                        # 1. Try Multi-Block Extraction (S/N + Grade)
                        header_configs = []
                        for idx, row in df_raw.iterrows():
                            row_vals = [str(v).lower().strip() for v in row.values]
                            sn_idx = -1
                            grade_idx = -1
                            
                            # Find S/N
                            for c_i, val in enumerate(row_vals):
                                if val in sn_keywords or any(val.startswith(k) for k in sn_keywords):
                                    # Avoid matching empty strings or very long text
                                    if 0 < len(val) < 20:
                                        sn_idx = c_i
                                        break
                            
                            # Find Grade
                            for c_i, val in enumerate(row_vals):
                                if any(k in val for k in grade_names):
                                    if len(val) < 20:
                                        grade_idx = c_i
                                        break
                            
                            if sn_idx != -1 and grade_idx != -1:
                                header_configs.append((idx, sn_idx, grade_idx))
                        
                        if header_configs:
                            collected_data = []
                            for k in range(len(header_configs)):
                                start_row, sn_c, grade_c = header_configs[k]
                                if k < len(header_configs) - 1:
                                    end_row = header_configs[k+1][0]
                                else:
                                    end_row = len(df_raw)
                                
                                for r in range(start_row + 1, end_row):
                                    if r >= len(df_raw): break
                                    row = df_raw.iloc[r]
                                    try:
                                        s_val = row.iloc[sn_c]
                                        g_val = row.iloc[grade_c]
                                        
                                        # Validate numeric S/N and Grade
                                        s_num = float(s_val)
                                        g_num = float(g_val)
                                        
                                        collected_data.append({'sn': s_num, 'score': g_num})
                                    except (ValueError, TypeError):
                                        continue
                            
                            if collected_data:
                                df_clean = pd.DataFrame(collected_data)
                                total_students_count = len(df_clean)
                                
                                df_clean = self.apply_sorting(df_clean)
                                
                                scores = df_clean['score']

                        # 2. Fallback to Single Block logic if no scores found
                        if scores.empty:
                            df = xl.parse(sheet_idx)
                            target_col = find_col(df, grade_names)

                            # If not found, try to find header row in first 10 rows
                            if not target_col:
                                for i in range(min(10, len(df_raw))):
                                    row_values = [str(v).lower() for v in df_raw.iloc[i].values]
                                    if any(any(name in val for name in grade_names) for val in row_values):
                                        df = xl.parse(sheet_idx, header=i)
                                        target_col = find_col(df, grade_names)
                                        if target_col:
                                            break
                            
                            if not target_col:
                                raise ValueError(f"Could not find a grade column. Expected one of: {', '.join(grade_names)}")

                            # Ensure numeric data
                            scores = pd.to_numeric(df[target_col], errors='coerce').dropna()
                            total_students_count = scores.count()
                            
                            # Handle S/N Column for Total Students and Sorting
                            sn_col = find_col(df, sn_keywords)
                            
                            temp_data = {'score': scores}
                            if sn_col:
                                # Align sn series with scores series by index
                                sn_series = pd.to_numeric(df.loc[scores.index, sn_col], errors='coerce')
                                temp_data['sn'] = sn_series
                                
                                # For total count
                                full_sn_series = pd.to_numeric(df[sn_col], errors='coerce').dropna()
                                if not full_sn_series.empty:
                                    total_students_count = int(full_sn_series.max())
                                    
                            temp_df = pd.DataFrame(temp_data)
                            sorted_df = self.apply_sorting(temp_df)
                            scores = sorted_df['score']

                        if scores.empty:
                            raise ValueError(f"Column found but contained no valid numeric data.")
                        
                        self.add_report_page(pdf, scores, sheet_name, total_students_count)

                    except Exception as e:
                        if self.process_all_sheets:
                            # Create an error page for this sheet
                            fig_err = plt.figure(figsize=(8.5, 11))
                            ax_err = fig_err.add_subplot(111)
                            ax_err.axis('off')
                            ax_err.text(0.5, 0.5, f"Sheet: {sheet_name}\n\nSkipped due to error:\n{str(e)}", 
                                        ha='center', va='center', fontsize=12, color='red')
                            pdf.savefig(fig_err)
                            plt.close(fig_err)
                        else:
                            # If single sheet mode, raise the error to be caught by the worker
                            raise e
        finally:
            xl.close()

    def add_report_page(self, pdf, scores, sheet_name, total_students_count):
        """Generates a single report page (Stats + Chart) and adds it to the PDF."""
        # Statistics
        mean_val = scores.mean()
        median_val = scores.median()
        mode_val = scores.mode()
        mode_str = ', '.join(map(lambda x: f'{x:.2f}', mode_val.tolist())) if not mode_val.empty else 'N/A'
        std_dev = scores.std()
        min_val = scores.min()
        max_val = scores.max()

        # --- Page 1: Statistics Table ---
        fig_table = plt.figure(figsize=(8.5, 11))
        ax_table = fig_table.add_subplot(111)
        ax_table.axis('off')

        stats_data = [
            ['Total Students', f'{total_students_count}'],
            ['Mean Score', f'{mean_val:.2f}'],
            ['Median Score', f'{median_val:.2f}'],
            ['Mode Score(s)', mode_str],
            ['Standard Deviation', f'{std_dev:.2f}'],
            ['Minimum Score', f'{min_val:.2f}'],
            ['Maximum Score', f'{max_val:.2f}']
        ]
        
        table = ax_table.table(cellText=stats_data, colLabels=['Statistic', 'Value'], loc='center', cellLoc='left')
        table.auto_set_font_size(False)
        table.set_fontsize(12)
        table.scale(1.2, 2.5)

        # Style table
        for (row, col), cell in table.get_celld().items():
            if row == 0 or col == 0:
                cell.set_text_props(fontweight='bold')

        ax_table.set_title(f'Summary Statistics - {sheet_name}', fontsize=18, fontweight='bold', pad=40)
        pdf.savefig(fig_table, bbox_inches='tight')
        plt.close(fig_table)

        # --- Page 2: Chart ---
        fig_chart = plt.figure(figsize=(8.5, 11))
        ax_chart = fig_chart.add_subplot(111)
        
        if self.chart_type == "Histogram":
            ax_chart.hist(scores, bins=10, color=self.chart_color, edgecolor='black', alpha=0.7, label='Scores')
            ax_chart.set_title(f'Grades Distribution (Histogram) - {sheet_name}', fontsize=16, fontweight='bold')
            ax_chart.set_xlabel('Score', fontsize=12)
            ax_chart.set_ylabel('Frequency', fontsize=12)

        elif self.chart_type == "Bar Chart":
            score_counts = scores.value_counts().sort_index()
            ax_chart.bar(score_counts.index, score_counts.values, color=self.chart_color, edgecolor='black')
            ax_chart.set_title(f'Grades Distribution (Bar Chart) - {sheet_name}', fontsize=16, fontweight='bold')
            ax_chart.set_xlabel('Score', fontsize=12)
            ax_chart.set_ylabel('Number of Students', fontsize=12)

        elif self.chart_type == "Pie Chart":
            try:
                if self.pie_bins_text:
                    bins = sorted([float(b.strip()) for b in self.pie_bins_text.split(',')])
                    if len(bins) < 2:
                        raise ValueError("Pie Bins must have at least two values.")
                else: # Default bins for a 0-100 scale
                    bins = [0, 60, 70, 80, 90, 101]
            except Exception as e:
                raise ValueError(f"Invalid Pie Bins. Use comma-separated numbers (e.g., 0,5,10,20). Error: {e}")

            labels = [f'{bins[i]} to {bins[i+1]}' for i in range(len(bins) - 1)]
            grade_counts = pd.cut(scores, bins=bins, labels=labels, right=False, include_lowest=True).value_counts().sort_index()
            
            grade_counts = grade_counts[grade_counts > 0]

            # Generate shades of the selected chart_color
            alphas = np.linspace(0.5, 1.0, len(grade_counts)) if len(grade_counts) > 1 else [1.0]
            colors = [matplotlib.colors.to_rgba(self.chart_color, alpha=a) for a in alphas]

            # Explode the largest slice
            explode = [0.1 if x == grade_counts.max() else 0 for x in grade_counts]

            ax_chart.pie(grade_counts, labels=grade_counts.index, autopct='%1.1f%%', startangle=140,
                colors=colors, explode=explode, wedgeprops={'edgecolor': 'white', 'linewidth': 2})
            ax_chart.set_title(f'Grade Distribution by Category (Pie Chart) - {sheet_name}', fontsize=16, fontweight='bold')
            ax_chart.axis('equal')

        elif self.chart_type == "Ogive (Frequency Curve)":
            counts, bin_edges = np.histogram(scores, bins=10)
            cum_freq = np.cumsum(counts)
            
            ogive_x = np.insert(bin_edges[1:], 0, bin_edges[0])
            ogive_y = np.insert(cum_freq, 0, 0)
            ax_chart.plot(ogive_x, ogive_y, marker='o', linestyle='-', color=self.chart_color)

            ax_chart.set_title(f'Ogive (Cumulative Frequency Curve) - {sheet_name}', fontsize=16, fontweight='bold')
            ax_chart.set_xlabel('Scores', fontsize=12)
            ax_chart.set_ylabel('Cumulative Number of Students', fontsize=12)
            ax_chart.grid(True, which='both', linestyle='--', linewidth=0.5)

        if self.chart_type in ["Histogram", "Bar Chart"]:
            ax_chart.axvline(mean_val, color='#e74c3c', linestyle='--', linewidth=2, label=f'Mean: {mean_val:.2f}')
            ax_chart.axvline(median_val, color='#2ecc71', linestyle='-', linewidth=2, label=f'Median: {median_val:.2f}')
            ax_chart.legend()
            ax_chart.grid(axis='y', alpha=0.3)

        fig_chart.subplots_adjust(bottom=0.15)
        fig_chart.text(0.1, 0.05, f"Generated by GradeVisualizer on {datetime.datetime.now().strftime('%Y-%m-%d %H:%M')}", 
                    fontsize=10, color='gray')

        pdf.savefig(fig_chart)
        plt.close(fig_chart)

class LoadingSpinner(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.angle = 0
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.rotate)
        self.timer.start(50)
        self.setFixedSize(100, 100)

    def rotate(self):
        self.angle = (self.angle + 30) % 360
        self.update()

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        painter.translate(self.width() / 2, self.height() / 2)
        painter.rotate(self.angle)
        
        pen = QPen(QColor("#3498db"))
        pen.setWidth(6)
        pen.setCapStyle(Qt.RoundCap)
        painter.setPen(pen)
        painter.drawArc(QRectF(-30, -30, 60, 60), 0, 270 * 16)

class IntroWindow(QWidget):
    finished = Signal()

    def __init__(self):
        super().__init__()
        self.setWindowFlags(Qt.Window | Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint | Qt.SplashScreen)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.resize(700, 500)
        self.center()
        
        self.layout = QVBoxLayout(self)
        self.container = QWidget()
        self.container.setObjectName("Container")
        self.container.setStyleSheet("""
            QWidget#Container { 
                background: qlineargradient(spread:pad, x1:0, y1:0, x2:1, y2:1, stop:0 #fdfefe, stop:1 #eaf2f8);
                border: 2px solid #3498db; 
                border-radius: 15px; 
            }
            QLabel { color: #2c3e50; background-color: transparent; padding: 10px; }
            QTableWidget { 
                background-color: #ffffff; 
                color: #34495e;
                border: 1px solid #d0d3d4; 
                gridline-color: #e5e8e8; 
            }
            QHeaderView::section {
                background-color: #eaf2f8;
                color: #17202a;
                font-weight: bold;
                padding: 5px;
                border: 1px solid #d0d3d4;
            }
        """)
        self.container_layout = QVBoxLayout(self.container)
        self.layout.addWidget(self.container)
        
        self.content_label = QLabel()
        self.content_label.setAlignment(Qt.AlignCenter)
        self.content_label.setWordWrap(True)
        self.container_layout.addWidget(self.content_label)
        
        self.table = QTableWidget()
        self.table.hide()
        self.container_layout.addWidget(self.table)
        
        self.spinner = LoadingSpinner()
        self.spinner.hide()
        self.container_layout.addWidget(self.spinner, 0, Qt.AlignCenter)

        # Navigation Buttons
        self.prev_btn = QPushButton("Previous")
        self.prev_btn.setCursor(Qt.PointingHandCursor)
        self.prev_btn.setStyleSheet("""
            QPushButton { background-color: #7f8c8d; color: white; border: none; padding: 8px 15px; border-radius: 4px; font-weight: bold; }
            QPushButton:hover { background-color: #95a5a6; }
            QPushButton:disabled { background-color: #bdc3c7; }
        """)
        self.prev_btn.clicked.connect(self.prev_step)

        self.skip_btn = QPushButton("Next")
        self.skip_btn.setCursor(Qt.PointingHandCursor)
        self.skip_btn.setStyleSheet("""
            QPushButton { background-color: #e74c3c; color: white; border: none; padding: 8px 15px; border-radius: 4px; font-weight: bold; }
            QPushButton:hover { background-color: #c0392b; }
        """)
        self.skip_btn.clicked.connect(self.next_step)
        
        btn_layout = QHBoxLayout()
        btn_layout.addWidget(self.prev_btn)
        btn_layout.addStretch()
        btn_layout.addWidget(self.skip_btn)
        self.container_layout.addLayout(btn_layout)

        self.step = 0
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.next_step)
        self.timer.start(60000) # Start with 60s for first message
        
        self.update_display()

    def center(self):
        qr = self.frameGeometry()
        cp = QApplication.primaryScreen().availableGeometry().center()
        qr.moveCenter(cp)
        self.move(qr.topLeft())

    def prev_step(self):
        if self.step > 0:
            self.step -= 1
            self.update_display()
            self.restart_timer()

    def restart_timer(self):
        self.timer.stop()
        if self.step < 3:
            self.timer.start(60000)
        elif self.step == 3:
            self.timer.start(3000)

    def next_step(self):
        self.step += 1
        self.update_display()
        self.restart_timer()

    def update_display(self):
        self.prev_btn.setEnabled(self.step > 0)
        
        if self.step == 0:
            self.content_label.setText("<h1 style='color:#2980b9'>Hello Madam Eloise</h1><h2>Welcome to Grade Visualizer Tool (G.V.T)</h2>")
            self.table.hide()
            self.spinner.hide()
            self.prev_btn.show()
            self.skip_btn.show()
        elif self.step == 1:
            self.content_label.setText(
                "<div style='font-size: 16px; padding: 20px;'>"
                "This software has been developed by a Tanzanian software developer<br>"
                "<b style='font-size: 18px; color: #e67e22'>George Julius Enock</b><br><br>"
                "A student from D.I.T taking Communication System Technology (C.S.T) course<br>"
                "and passionate in computer programming.<br><br>"
                "<b>Email:</b> georgeenockjulius@gmail.com<br>"
                "<b>Phone:</b> +255 786 169 481</div>"
            )
            self.table.hide()
            self.spinner.hide()
            self.prev_btn.show()
            self.skip_btn.show()
        elif self.step == 2:
            self.content_label.setText("<h3>Accepted Excel Format</h3><p>Please ensure your data matches the columns below:</p>")
            self.table.setColumnCount(6)
            self.table.setHorizontalHeaderLabels(["S/No", "Full Name", "Reg Number", "EXAM NO", "Marks", "Sign"])
            self.table.setRowCount(3)
            # Add demo data
            demo_data = [
                ["1", "John Doe", "DIT/001", "EX001", "85", ""],
                ["2", "Jane Smith", "DIT/002", "EX002", "92", ""],
                ["3", "Albert K.", "DIT/003", "EX003", "78", ""]
            ]
            for r, row_data in enumerate(demo_data):
                for c, val in enumerate(row_data):
                    self.table.setItem(r, c, QTableWidgetItem(val))
            
            self.table.verticalHeader().setVisible(False)
            self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
            self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
            self.table.show()
            self.spinner.hide()
            self.prev_btn.show()
            self.skip_btn.show()
            
        elif self.step == 3:
            self.content_label.setText("<h3>Loading Application...</h3>")
            self.table.hide()
            self.spinner.show()
            self.skip_btn.hide()
            self.prev_btn.hide()
            
        elif self.step == 4:
            self.timer.stop()
            self.close()
            self.finished.emit()

class GradeVisualizer(QMainWindow):
    """
    Main GUI Application Window.
    """
    def __init__(self):
        super().__init__()
        self.setWindowTitle("GradeVisualizer")
        self.resize(1100, 700)
        self.last_table_state = None
        self.setup_ui()
        self.setup_theme()

    def setup_theme(self):
        self.themes = {
            "light": {
                "palette": self.create_light_palette(),
                "stylesheet": self.get_light_stylesheet(),
                "icon": "🌙"
            },
            "dark": {
                "palette": self.create_dark_palette(),
                "stylesheet": self.get_dark_stylesheet(),
                "icon": "☀️"
            }
        }
        self.settings = QSettings("GJE", "GradeVisualizer")
        saved_theme = self.settings.value("theme", None)

        if saved_theme is None:
            # No theme saved, try to detect from OS. Default to light.
            system_theme = "light"
            try:
                # This is the modern way for Qt6.2+
                color_scheme_enum = QApplication.styleHints().colorScheme()
                if color_scheme_enum == Qt.ColorScheme.Dark:
                    system_theme = "dark"
            except AttributeError:
                pass # styleHints().colorScheme() not available, will use default
            saved_theme = system_theme
        
        # Set the opposite so the first toggle sets the correct saved theme
        self.current_theme_name = "dark" if saved_theme == "light" else "light"
        self.toggle_theme()

    def toggle_theme(self):
        # Switch to the other theme
        self.current_theme_name = "dark" if self.current_theme_name == "light" else "light"
        theme_data = self.themes[self.current_theme_name]
        
        app = QApplication.instance()
        app.setPalette(theme_data["palette"])
        app.setStyleSheet(theme_data["stylesheet"])
        
        self.theme_toggle_btn.setText(theme_data["icon"])
        
        # Save preference
        self.settings.setValue("theme", self.current_theme_name)

    def create_light_palette(self):
        palette = QPalette()
        palette.setColor(QPalette.Window, QColor(244, 246, 249))
        palette.setColor(QPalette.WindowText, QColor(52, 58, 64))
        palette.setColor(QPalette.Base, QColor(255, 255, 255))
        palette.setColor(QPalette.AlternateBase, QColor(248, 249, 250))
        palette.setColor(QPalette.ToolTipBase, QColor(255, 255, 255))
        palette.setColor(QPalette.ToolTipText, QColor(52, 58, 64))
        palette.setColor(QPalette.Text, QColor(52, 58, 64))
        palette.setColor(QPalette.Button, QColor(255, 255, 255))
        palette.setColor(QPalette.ButtonText, QColor(52, 58, 64))
        palette.setColor(QPalette.Link, QColor(42, 130, 218))
        palette.setColor(QPalette.Highlight, QColor(52, 152, 219))
        palette.setColor(QPalette.HighlightedText, Qt.white)
        return palette

    def create_dark_palette(self):
        palette = QPalette()
        palette.setColor(QPalette.Window, QColor(44, 62, 80))
        palette.setColor(QPalette.WindowText, QColor(236, 240, 241))
        palette.setColor(QPalette.Base, QColor(52, 73, 94))
        palette.setColor(QPalette.AlternateBase, QColor(74, 98, 122))
        palette.setColor(QPalette.ToolTipBase, QColor(52, 73, 94))
        palette.setColor(QPalette.ToolTipText, QColor(236, 240, 241))
        palette.setColor(QPalette.Text, QColor(236, 240, 241))
        palette.setColor(QPalette.Button, QColor(86, 101, 115))
        palette.setColor(QPalette.ButtonText, QColor(236, 240, 241))
        palette.setColor(QPalette.Link, QColor(52, 152, 219))
        palette.setColor(QPalette.Highlight, QColor(52, 152, 219))
        palette.setColor(QPalette.HighlightedText, QColor(255, 255, 255))
        return palette

    def get_light_stylesheet(self):
        return """
            QMainWindow { background-color: #f4f6f9; }
            QWidget { font-family: 'Segoe UI', 'Roboto', 'Arial', sans-serif; font-size: 14px; color: #343a40; }
            QLabel { color: #495057; }
            QLabel#TitleLabel { font-size: 26px; font-weight: bold; color: #2c3e50; margin-bottom: 15px; }
            QLabel#StatsLabel { font-size: 16px; color: #2c3e50; background-color: #e9ecef; padding: 8px 15px; border-radius: 6px; border: 1px solid #dee2e6; }
            QLineEdit, QComboBox { padding: 10px; border: 1px solid #ced4da; border-radius: 6px; background-color: #ffffff; color: #343a40; selection-background-color: #3498db; selection-color: #ffffff; }
            QLineEdit:focus, QComboBox:focus { border: 1px solid #3498db; background-color: #fff; }
            QComboBox QAbstractItemView { background-color: #ffffff; color: #343a40; selection-background-color: #3498db; selection-color: #ffffff; border: 1px solid #ced4da; }
            QCheckBox { spacing: 8px; }
            QPushButton { background-color: #ffffff; border: 1px solid #ced4da; border-radius: 6px; padding: 8px 16px; color: #495057; font-weight: 600; }
            QPushButton:hover { background-color: #f8f9fa; border-color: #adb5bd; color: #212529; }
            QPushButton:pressed { background-color: #e9ecef; }
            QPushButton:disabled { background-color: #e9ecef; color: #adb5bd; border-color: #dee2e6; }
            QPushButton#GenerateBtn { background-color: #3498db; color: white; border: 1px solid #2980b9; font-size: 16px; padding: 12px 20px; }
            QPushButton#GenerateBtn:hover { background-color: #2980b9; }
            QPushButton#GenerateBtn:pressed { background-color: #2573a7; }
            QPushButton#DangerBtn { color: #e74c3c; border: 1px solid #e74c3c; background-color: #fff; }
            QPushButton#DangerBtn:hover { background-color: #e74c3c; color: white; }
            QTableWidget { background-color: white; border: 1px solid #dee2e6; border-radius: 6px; gridline-color: #f1f3f5; selection-background-color: #e7f5ff; selection-color: #000; alternate-background-color: #f8f9fa; }
            QHeaderView::section { background-color: #ffffff; padding: 10px; border: none; border-bottom: 2px solid #dee2e6; font-weight: bold; color: #495057; }
        """

    def get_dark_stylesheet(self):
        return """
            QMainWindow { background-color: #2c3e50; }
            QWidget { font-family: 'Segoe UI', 'Roboto', 'Arial', sans-serif; font-size: 14px; color: #ecf0f1; }
            QLabel { color: #bdc3c7; }
            QLabel#TitleLabel { font-size: 26px; font-weight: bold; color: #ffffff; margin-bottom: 15px; }
            QLabel#StatsLabel { font-size: 16px; color: #ecf0f1; background-color: #34495e; padding: 8px 15px; border-radius: 6px; border: 1px solid #4a627a; }
            QLineEdit, QComboBox { padding: 10px; border: 1px solid #566573; border-radius: 6px; background-color: #34495e; color: #ecf0f1; selection-background-color: #3498db; selection-color: #ffffff; }
            QLineEdit:focus, QComboBox:focus { border: 1px solid #3498db; background-color: #4a627a; }
            QComboBox QAbstractItemView { background-color: #34495e; color: #ecf0f1; selection-background-color: #3498db; selection-color: #ffffff; border: 1px solid #566573; }
            QCheckBox { spacing: 8px; }
            QPushButton { background-color: #566573; border: 1px solid #7f8c8d; border-radius: 6px; padding: 8px 16px; color: #ecf0f1; font-weight: 600; }
            QPushButton:hover { background-color: #7f8c8d; border-color: #95a5a6; }
            QPushButton:pressed { background-color: #4a627a; }
            QPushButton:disabled { background-color: #34495e; color: #7f8c8d; border-color: #4a627a; }
            QPushButton#GenerateBtn { background-color: #3498db; color: white; border: 1px solid #2980b9; font-size: 16px; padding: 12px 20px; }
            QPushButton#GenerateBtn:hover { background-color: #2980b9; }
            QPushButton#GenerateBtn:pressed { background-color: #2573a7; }
            QPushButton#DangerBtn { color: #e74c3c; border: 1px solid #e74c3c; background-color: transparent; }
            QPushButton#DangerBtn:hover { background-color: #e74c3c; color: white; }
            QTableWidget { background-color: #34495e; border: 1px solid #4a627a; border-radius: 6px; gridline-color: #2c3e50; selection-background-color: #528baf; selection-color: #ffffff; alternate-background-color: #4a627a; }
            QHeaderView::section { background-color: #566573; padding: 10px; border: none; border-bottom: 2px solid #4a627a; font-weight: bold; color: #ecf0f1; }
        """

    def setup_ui(self):
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        self.setCentralWidget(scroll_area)

        content_widget = QWidget()
        scroll_area.setWidget(content_widget)
        layout = QVBoxLayout(content_widget)
        layout.setSpacing(10)
        layout.setContentsMargins(20, 20, 20, 20)

        # Title Label
        title_label = QLabel("Grade Visualizer Tool")
        title_label.setObjectName("TitleLabel")
        title_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(title_label)

        # Input Layout
        input_layout = QHBoxLayout()
        
        self.url_input = QLineEdit()
        self.url_input.setPlaceholderText("Paste Link or Browse File...")
        input_layout.addWidget(self.url_input)

        self.browse_btn = QPushButton("Browse")
        self.browse_btn.clicked.connect(self.browse_file)
        input_layout.addWidget(self.browse_btn)

        self.load_data_btn = QPushButton("Load Data")
        self.load_data_btn.clicked.connect(self.load_data)
        input_layout.addWidget(self.load_data_btn)

        self.open_link_btn = QPushButton("Open File/Link")
        self.open_link_btn.clicked.connect(self.open_drive_link)
        input_layout.addWidget(self.open_link_btn)

        self.reset_btn = QPushButton("Reset System")
        self.reset_btn.clicked.connect(self.reset_system)
        self.reset_btn.setObjectName("DangerBtn")
        input_layout.addWidget(self.reset_btn)

        self.theme_toggle_btn = QPushButton()
        self.theme_toggle_btn.setToolTip("Toggle Dark/Light Mode")
        self.theme_toggle_btn.setFixedWidth(45)
        self.theme_toggle_btn.clicked.connect(self.toggle_theme)
        input_layout.addWidget(self.theme_toggle_btn)

        layout.addLayout(input_layout)

        # Options Layout (Chart Type, Color, Sheet)
        options_layout = QHBoxLayout()
        
        # Chart Type
        options_layout.addWidget(QLabel("Type:"))
        self.chart_type_combo = QComboBox()
        self.chart_type_combo.addItems(["Histogram", "Bar Chart", "Pie Chart", "Ogive (Frequency Curve)"])
        self.chart_type_combo.setToolTip("Select the type of chart to generate in the PDF report.")
        self.chart_type_combo.currentTextChanged.connect(self.update_chart_options)
        options_layout.addWidget(self.chart_type_combo)

        # Color
        options_layout.addWidget(QLabel("Color:"))
        self.color_combo = QComboBox()
        self.color_combo.addItems(["Blue", "Red", "Green", "Purple", "Orange", "Teal"])
        self.color_combo.setToolTip("Select the primary color theme for the generated chart.")
        options_layout.addWidget(self.color_combo)

        # Sheet Number
        options_layout.addWidget(QLabel("Sheet No:"))
        self.sheet_input = QLineEdit("1")
        self.sheet_input.setToolTip("Enter the sheet number to process (1-based index).")
        self.sheet_input.setFixedWidth(40)
        self.sheet_input.setAlignment(Qt.AlignCenter)
        options_layout.addWidget(self.sheet_input)

        # All Sheets Checkbox
        self.all_sheets_cb = QCheckBox("All")
        self.all_sheets_cb.toggled.connect(self.toggle_sheet_input)
        self.all_sheets_cb.setToolTip("Check to process all sheets in the workbook.")
        options_layout.addWidget(self.all_sheets_cb)

        # Sort Option
        options_layout.addWidget(QLabel("Sort:"))
        self.sort_combo = QComboBox()
        self.sort_combo.addItems(["S/N (Ascending)", "S/N (Descending)", "Score (Ascending)", "Score (Descending)"])
        self.sort_combo.setToolTip("Select how to sort the data before generating statistics.")
        self.sort_combo.currentTextChanged.connect(self.sort_table)
        options_layout.addWidget(self.sort_combo)

        layout.addLayout(options_layout)

        # Pie Chart Bins (initially hidden)
        self.pie_bins_widget = QWidget()
        pie_bins_layout = QHBoxLayout(self.pie_bins_widget)
        pie_bins_layout.setContentsMargins(0, 0, 0, 0)
        self.pie_bins_label = QLabel("Pie Bins (e.g., 0,5,10,20):")
        self.pie_bins_input = QLineEdit()
        self.pie_bins_input.setToolTip("Define custom boundaries for Pie Chart slices, separated by commas.")
        self.pie_bins_input.setPlaceholderText("Defaults to 0-100 scale if empty")
        pie_bins_layout.addWidget(self.pie_bins_label)
        pie_bins_layout.addWidget(self.pie_bins_input)
        layout.addWidget(self.pie_bins_widget)

        self.pie_bins_widget.setVisible(False) # Hide by default

        # --- Search Bar ---
        search_layout = QHBoxLayout()
        search_layout.addWidget(QLabel("Search Name:"))
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Search by Name, Exam No, or Reg No...")
        self.search_input.textChanged.connect(self.filter_table_by_name)
        search_layout.addWidget(self.search_input)
        layout.addLayout(search_layout)

        # --- Data Table Section ---
        self.data_table = QTableWidget()
        self.data_table.setColumnCount(6)
        self.data_table.setHorizontalHeaderLabels(["S/N", "Name", "Reg No", "Exam No", "Score", "Sign"])
        self.data_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.data_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.data_table.itemChanged.connect(self.update_class_average)
        self.data_table.setAlternatingRowColors(True)
        self.data_table.setShowGrid(False)
        layout.addWidget(self.data_table)

        # Table Controls
        table_controls = QHBoxLayout()
        self.add_row_btn = QPushButton("Add Row")
        self.add_row_btn.setToolTip("Add a new empty row to the end of the table.")
        self.add_row_btn.clicked.connect(self.add_row)
        table_controls.addWidget(self.add_row_btn)

        self.del_row_btn = QPushButton("Delete Selected")
        self.del_row_btn.clicked.connect(self.delete_row)
        self.del_row_btn.setToolTip("Delete the currently selected row(s) from the table.")
        self.del_row_btn.setObjectName("DangerBtn")
        table_controls.addWidget(self.del_row_btn)

        self.undo_btn = QPushButton("Undo Delete")
        self.undo_btn.clicked.connect(self.undo_delete)
        self.undo_btn.setToolTip("Restore the rows from the last deletion action.")
        self.undo_btn.setEnabled(False)
        table_controls.addWidget(self.undo_btn)

        self.renumber_btn = QPushButton("Auto-Arrange S/N")
        self.renumber_btn.clicked.connect(self.renumber_sn)
        self.renumber_btn.setToolTip("Renumbers the S/N column sequentially (1, 2, 3...).")
        table_controls.addWidget(self.renumber_btn)

        self.check_dup_btn = QPushButton("Mark Duplicates")
        self.check_dup_btn.clicked.connect(self.mark_duplicates)
        self.check_dup_btn.setToolTip("Find and highlight rows with duplicate names.")
        table_controls.addWidget(self.check_dup_btn)

        self.clean_btn = QPushButton("Remove Empty Scores")
        self.clean_btn.clicked.connect(self.remove_empty_scores)
        self.clean_btn.setToolTip("Find and remove rows that have no score.")
        table_controls.addWidget(self.clean_btn)

        layout.addLayout(table_controls)
        # ---------------------------

        # --- Stats Display ---
        stats_layout = QHBoxLayout()
        stats_layout.addWidget(QLabel("<b>Class Average:</b>"))
        self.average_label = QLabel("N/A")
        self.average_label.setObjectName("StatsLabel")
        stats_layout.addWidget(self.average_label)
        stats_layout.addStretch()
        layout.addLayout(stats_layout)
        # ---------------------

        # Generate Button
        self.generate_btn = QPushButton("Generate PDF")
        self.generate_btn.setCursor(Qt.PointingHandCursor)
        self.generate_btn.setToolTip("Generate the PDF report from the data in the table or the specified file.")
        self.generate_btn.setObjectName("GenerateBtn")
        self.generate_btn.clicked.connect(self.start_generation)
        layout.addWidget(self.generate_btn)

        # Open Folder Button
        self.open_folder_btn = QPushButton("Open Output Folder")
        self.open_folder_btn.setToolTip("Open the folder where PDF reports are saved.")
        self.open_folder_btn.clicked.connect(self.open_output_folder)
        layout.addWidget(self.open_folder_btn)

        # Progress Bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.progress_bar.setTextVisible(False)
        layout.addWidget(self.progress_bar)

        # Status Label
        self.status_label = QLabel("Ready")
        self.status_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.status_label)

    def browse_file(self):
        file_name, _ = QFileDialog.getOpenFileName(self, "Select Excel File", "", "Excel Files (*.xlsx *.xls)")
        if file_name:
            self.url_input.setText(file_name)

    def open_drive_link(self):
        url = self.url_input.text().strip()
        if not url:
            QMessageBox.warning(self, "Input Error", "Please enter a link or file path first.")
            return
        
        try:
            if os.path.isfile(url):
                os.startfile(url)
            else:
                webbrowser.open(url)
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Could not open link or file:\n{str(e)}")

    def open_output_folder(self):
        docs_dir = os.path.join(os.path.expanduser('~'), 'Documents')
        save_dir = os.path.join(docs_dir, 'Grade Visualizer')
        os.makedirs(save_dir, exist_ok=True)
        try:
            os.startfile(save_dir)
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Could not open output folder:\n{str(e)}")

    def toggle_sheet_input(self, checked):
        self.sheet_input.setEnabled(not checked)

    def update_chart_options(self, chart_type):
        """Show or hide UI elements based on the selected chart type."""
        self.pie_bins_widget.setVisible(chart_type == "Pie Chart")

    def load_data(self):
        input_source = self.url_input.text().strip()
        if not input_source:
            QMessageBox.warning(self, "Input Error", "Please enter a valid link or file path.")
            return
        
        sheet_num = self.sheet_input.text().strip()
        
        self.load_data_btn.setEnabled(False)
        self.status_label.setText("Loading Data...")
        
        self.worker = DataWorker(input_source, "", "", sheet_num, "", "", False, task_type="file_to_data")
        self.worker.data_ready.connect(self.populate_table)
        self.worker.finished.connect(self.load_finished)
        self.worker.start()

    def load_finished(self, success, message):
        self.load_data_btn.setEnabled(True)
        self.status_label.setText(message if success else "Error Loading Data")
        if not success:
            QMessageBox.critical(self, "Error", message)

    def populate_table(self, data):
        self.data_table.blockSignals(True)
        self.data_table.setRowCount(len(data))
        for i, row in enumerate(data):
            self.data_table.setItem(i, 0, QTableWidgetItem(str(row.get('sn', ''))))
            self.data_table.setItem(i, 1, QTableWidgetItem(str(row.get('name', ''))))
            self.data_table.setItem(i, 2, QTableWidgetItem(str(row.get('reg_no', ''))))
            self.data_table.setItem(i, 3, QTableWidgetItem(str(row.get('exam_no', ''))))
            self.data_table.setItem(i, 4, QTableWidgetItem(str(row.get('score', ''))))
            self.data_table.setItem(i, 5, QTableWidgetItem(str(row.get('sign', ''))))
        self.data_table.blockSignals(False)
        self.update_class_average()
        # Re-apply filter after populating
        self.filter_table_by_name(self.search_input.text())

    def filter_table_by_name(self, text):
        """Hides or shows rows based on the search text in Name, Exam No, or Reg No columns."""
        search_text = text.lower()
        for row in range(self.data_table.rowCount()):
            name_item = self.data_table.item(row, 1) # Column 1 is 'Name'
            reg_item = self.data_table.item(row, 2)  # Column 2 is 'Reg No'
            exam_item = self.data_table.item(row, 3) # Column 3 is 'Exam No'
            
            row_text = ""
            if name_item: row_text += name_item.text().lower() + " "
            if reg_item: row_text += reg_item.text().lower() + " "
            if exam_item: row_text += exam_item.text().lower()
            
            if row_text:
                self.data_table.setRowHidden(row, search_text not in row_text)
            else:
                self.data_table.setRowHidden(row, bool(search_text))

    def add_row(self):
        row_idx = self.data_table.rowCount()
        self.data_table.insertRow(row_idx)

    def delete_row(self):
        # Use selectionModel to correctly get selected rows, even if they are empty
        selected_indexes = self.data_table.selectionModel().selectedRows()
        selected_rows = sorted([index.row() for index in selected_indexes], reverse=True)
        if not selected_rows:
            return

        reply = QMessageBox.question(self, "Confirm Deletion", f"Are you sure you want to delete {len(selected_rows)} row(s)?",
                                     QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        
        if reply == QMessageBox.Yes:
            # Save state for undo
            self.last_table_state = self.get_table_data()
            self.undo_btn.setEnabled(True)
            
            for row in selected_rows:
                self.data_table.removeRow(row)
            self.update_class_average()

    def get_table_data(self):
        """Captures current table data for undo functionality."""
        data = []
        for r in range(self.data_table.rowCount()):
            sn = self.data_table.item(r, 0).text() if self.data_table.item(r, 0) else ""
            name = self.data_table.item(r, 1).text() if self.data_table.item(r, 1) else ""
            reg_no = self.data_table.item(r, 2).text() if self.data_table.item(r, 2) else ""
            exam_no = self.data_table.item(r, 3).text() if self.data_table.item(r, 3) else ""
            score = self.data_table.item(r, 4).text() if self.data_table.item(r, 4) else ""
            sign = self.data_table.item(r, 5).text() if self.data_table.item(r, 5) else ""
            data.append({'sn': sn, 'name': name, 'reg_no': reg_no, 'exam_no': exam_no, 'score': score, 'sign': sign})
        return data

    def undo_delete(self):
        if self.last_table_state is not None:
            self.populate_table(self.last_table_state)
            self.last_table_state = None
            self.undo_btn.setEnabled(False)
            QMessageBox.information(self, "Undo", "Last deletion undone.")

    def sort_table(self):
        """Sorts the data in the table based on the dropdown selection."""
        if self.data_table.rowCount() == 0:
            return

        table_data = self.get_table_data()
        df = pd.DataFrame(table_data)

        # Coerce types for proper sorting
        df['score'] = pd.to_numeric(df['score'], errors='coerce')
        df['sn'] = pd.to_numeric(df['sn'], errors='coerce')
        
        sort_option = self.sort_combo.currentText()

        # Apply sorting logic
        if "S/N" in sort_option and 'sn' in df.columns:
            ascending = "Ascending" in sort_option
            df = df.sort_values(by='sn', ascending=ascending, na_position='last')
        elif "Score" in sort_option and 'score' in df.columns:
            ascending = "Ascending" in sort_option
            df = df.sort_values(by='score', ascending=ascending, na_position='last')
        
        # Convert back to list of dicts, handling NaNs for display
        sorted_data = df.where(pd.notnull(df), "").to_dict('records')
        self.populate_table(sorted_data)

    def update_class_average(self):
        """Calculates and displays the average score from the table."""
        scores = []
        for row in range(self.data_table.rowCount()):
            score_item = self.data_table.item(row, 4)
            if score_item and score_item.text().strip():
                try:
                    score = float(score_item.text())
                    scores.append(score)
                except ValueError:
                    continue
        
        if scores:
            average = np.mean(scores)
            self.average_label.setText(f"<b>{average:.2f}</b>")
        else:
            self.average_label.setText("N/A")

    def renumber_sn(self):
        """Auto-arranges S/N column sequentially."""
        for row in range(self.data_table.rowCount()):
            self.data_table.setItem(row, 0, QTableWidgetItem(str(row + 1)))

    def mark_duplicates(self):
        """Highlights rows with duplicate Names in red."""
        # Reset background
        for row in range(self.data_table.rowCount()):
            for col in range(self.data_table.columnCount()):
                item = self.data_table.item(row, col)
                if item:
                    item.setData(Qt.BackgroundRole, None)
                    item.setData(Qt.ForegroundRole, None)

        seen = {}
        duplicates = set()
        
        # Identify duplicates based on Name (Column 1)
        for row in range(self.data_table.rowCount()):
            name_item = self.data_table.item(row, 1)
            name = name_item.text().strip().lower() if name_item else ""
            if name:
                if name in seen:
                    duplicates.add(name)
                seen[name] = True
        
        # Highlight
        if self.current_theme_name == "dark":
            highlight_bg = QColor(118, 42, 42) # Dark, saturated red
            highlight_fg = QColor(240, 240, 240)
        else:
            highlight_bg = QColor(255, 200, 200) # Light red
            highlight_fg = QColor("black")

        count = 0
        for row in range(self.data_table.rowCount()):
            name_item = self.data_table.item(row, 1)
            name = name_item.text().strip().lower() if name_item else ""
            
            if name in duplicates:
                count += 1
                for col in range(self.data_table.columnCount()):
                    item = self.data_table.item(row, col)
                    if not item:
                        item = QTableWidgetItem("")
                        self.data_table.setItem(row, col, item)
                    item.setBackground(highlight_bg)
                    item.setForeground(highlight_fg)
        
        if count > 0:
            QMessageBox.information(self, "Duplicates", f"Marked {count} rows with repeated names.")
        else:
            QMessageBox.information(self, "Duplicates", "No duplicate names found.")

    def remove_empty_scores(self):
        """Removes rows where the score is empty or NaN."""
        rows_to_delete = []
        for row in range(self.data_table.rowCount()):
            score_item = self.data_table.item(row, 4)
            score_text = score_item.text().strip().lower() if score_item else ""
            
            if not score_text or score_text == 'nan':
                rows_to_delete.append(row)
        
        if not rows_to_delete:
            QMessageBox.information(self, "Clean Data", "No empty scores found.")
            return

        reply = QMessageBox.question(self, "Confirm Clean", f"Delete {len(rows_to_delete)} rows with empty/NaN scores?",
                                     QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        
        if reply == QMessageBox.Yes:
            self.last_table_state = self.get_table_data()
            self.undo_btn.setEnabled(True)
            
            for row in sorted(rows_to_delete, reverse=True):
                self.data_table.removeRow(row)
            self.update_class_average()
            
            QMessageBox.information(self, "Success", "Empty rows removed.")

    def reset_system(self):
        """Clears all inputs and data to start afresh."""
        reply = QMessageBox.question(self, "Reset System", "Are you sure you want to clear all data?", 
                                     QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.Yes:
            self.url_input.clear()
            self.sheet_input.setText("1")
            self.data_table.setRowCount(0)
            self.pie_bins_input.clear()
            self.all_sheets_cb.setChecked(False)
            self.status_label.setText("Ready")
            self.chart_type_combo.setCurrentIndex(0)
            self.color_combo.setCurrentIndex(0)
            self.sort_combo.setCurrentIndex(0)
            self.last_table_state = None
            self.undo_btn.setEnabled(False)
            self.update_class_average()
            self.search_input.clear()

    def start_generation(self):
        input_source = self.url_input.text().strip()
        has_table_data = self.data_table.rowCount() > 0

        if not input_source and not has_table_data:
            QMessageBox.warning(self, "Input Error", "Please enter a valid link or file path.")
            return
        
        chart_type = self.chart_type_combo.currentText()
        pie_bins_text = self.pie_bins_input.text().strip()
        sheet_num = self.sheet_input.text().strip()
        process_all_sheets = self.all_sheets_cb.isChecked()
        sort_option = self.sort_combo.currentText()
        
        # Map Color Name to Hex
        color_map = {
            "Blue": "#3498db", "Red": "#e74c3c", "Green": "#2ecc71",
            "Purple": "#9b59b6", "Orange": "#e67e22", "Teal": "#1abc9c"
        }
        chart_color = color_map.get(self.color_combo.currentText(), "#3498db")

        self.progress_bar.setValue(0)
        self.progress_bar.setVisible(True)

        self.generate_btn.setEnabled(False)
        self.url_input.setEnabled(False)
        self.browse_btn.setEnabled(False)
        self.open_link_btn.setEnabled(False)
        self.chart_type_combo.setEnabled(False)
        self.color_combo.setEnabled(False)
        self.sheet_input.setEnabled(False)
        self.all_sheets_cb.setEnabled(False)
        self.sort_combo.setEnabled(False)
        self.pie_bins_input.setEnabled(False)
        self.load_data_btn.setEnabled(False)
        self.add_row_btn.setEnabled(False)
        self.del_row_btn.setEnabled(False)
        self.undo_btn.setEnabled(False)
        self.reset_btn.setEnabled(False)
        self.renumber_btn.setEnabled(False)
        self.check_dup_btn.setEnabled(False)
        self.clean_btn.setEnabled(False)
        
        # Determine if we are generating from Table or File
        if has_table_data:
            # Extract data from table
            raw_data = []
            for i in range(self.data_table.rowCount()):
                sn = self.data_table.item(i, 0).text() if self.data_table.item(i, 0) else ""
                name = self.data_table.item(i, 1).text() if self.data_table.item(i, 1) else ""
                reg_no = self.data_table.item(i, 2).text() if self.data_table.item(i, 2) else ""
                exam_no = self.data_table.item(i, 3).text() if self.data_table.item(i, 3) else ""
                score = self.data_table.item(i, 4).text() if self.data_table.item(i, 4) else ""
                sign = self.data_table.item(i, 5).text() if self.data_table.item(i, 5) else ""
                raw_data.append({'sn': sn, 'name': name, 'reg_no': reg_no, 'exam_no': exam_no, 'score': score, 'sign': sign})
            
            source_arg = input_source if input_source else "User Data"
            self.worker = DataWorker(source_arg, chart_type, pie_bins_text, sheet_num, chart_color, sort_option, process_all_sheets, raw_data=raw_data, task_type="data_to_pdf")
        else:
            # Standard File Generation
            self.worker = DataWorker(input_source, chart_type, pie_bins_text, sheet_num, chart_color, sort_option, process_all_sheets, task_type="file_to_pdf")

        self.worker.status_update.connect(self.status_label.setText)
        self.worker.progress_update.connect(self.progress_bar.setValue)
        self.worker.finished.connect(self.process_finished)
        self.worker.start()

    def process_finished(self, success, message):
        self.progress_bar.setVisible(False)
        self.generate_btn.setEnabled(True)
        self.url_input.setEnabled(True)
        self.browse_btn.setEnabled(True)
        self.open_link_btn.setEnabled(True)
        self.chart_type_combo.setEnabled(True)
        self.color_combo.setEnabled(True)
        self.sheet_input.setEnabled(True)
        self.all_sheets_cb.setEnabled(True)
        self.sort_combo.setEnabled(True)
        self.pie_bins_input.setEnabled(True)
        self.load_data_btn.setEnabled(True)
        self.add_row_btn.setEnabled(True)
        self.del_row_btn.setEnabled(True)
        self.undo_btn.setEnabled(True if self.last_table_state else False)
        self.reset_btn.setEnabled(True)
        self.renumber_btn.setEnabled(True)
        self.check_dup_btn.setEnabled(True)
        self.clean_btn.setEnabled(True)
        
        if success:
            QMessageBox.information(self, "Success", message)
            self.status_label.setText("Ready")
            self.url_input.clear()
        else:
            QMessageBox.critical(self, "Error", f"An error occurred:\n{message}")
            self.status_label.setText("Error occurred")

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.dirname(os.path.abspath(__file__))

    return os.path.join(base_path, relative_path)

if __name__ == "__main__":
    # Ensure High DPI scaling is handled correctly on Windows 10/11
    os.environ["QT_AUTO_SCREEN_SCALE_FACTOR"] = "1"
    app = QApplication(sys.argv)
    
    # Set App Icon
    icon_path = resource_path("grade_visualizer_tool.ico")
    app.setWindowIcon(QIcon(icon_path))
    
    # Set application style
    app.setStyle('Fusion')

    # Show Intro Splash
    intro = IntroWindow()
    window = GradeVisualizer()
    
    intro.finished.connect(window.show)
    intro.show()
    
    sys.exit(app.exec())