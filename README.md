# Grade Visualizer Tool (G.V.T)

## Overview
Grade Visualizer Tool is a user-friendly desktop application built with Python and PySide6. It empowers educators to analyze student grades effectively by generating statistical summaries and visual charts. The tool supports data import from Excel files (including Google Drive links) and offers manual data entry capabilities.

## Features

### Data Management
- **Flexible Input**: Load data from local Excel files (`.xlsx`) or directly via Google Drive links.
- **Manual Entry**: A built-in spreadsheet-like interface allows for manual data entry and modification.
- **Data Cleaning Tools**:
  - **Auto-Arrange S/N**: Automatically renumbers the serial number column.
  - **Mark Duplicates**: Highlights rows with duplicate student names for easy identification.
  - **Remove Empty Scores**: Cleans the dataset by removing rows without valid scores.
- **Search**: Filter the data table by Name, Registration Number, or Exam Number.

### Visualization & Reporting
- **PDF Reports**: Generates detailed PDF reports containing both statistical summaries and charts.
- **Chart Types**:
  - Histogram
  - Bar Chart
  - Pie Chart (with customizable bins)
  - Ogive (Cumulative Frequency Curve)
- **Statistics**: Automatically calculates Total Students, Mean, Median, Mode, Standard Deviation, Minimum, and Maximum scores.

### User Interface
- **Theming**: Switch between Light and Dark modes.
- **Responsive Design**: Includes a splash screen and progress indicators for long-running tasks.

## Requirements

- Python 3.8+
- Required Python packages:
  - `PySide6`
  - `pandas`
  - `matplotlib`
  - `numpy`
  - `requests`
  - `openpyxl`

## Installation

1.  **Clone or Download** the repository to your local machine.
2.  **Install Dependencies**:
    Open your terminal or command prompt and run:
    ```bash
    pip install PySide6 pandas matplotlib numpy requests openpyxl
    ```
3.  **Icon File**: Ensure the file `grade_visualizer_tool.ico` is present in the project directory for the application icon to load correctly.

## Usage

1.  **Start the Application**:
    ```bash
    python grade_visualizer.py
    ```
2.  **Import Data**:
    - **From File**: Click **Browse** to select a local Excel file or paste a Google Drive link into the input field, then click **Load Data**.
    - **Manual Input**: Use the **Add Row** button to start entering data manually into the table.
3.  **Configure Settings**:
    - Choose the desired **Chart Type** and **Color**.
    - If loading from a file, specify the **Sheet No** (default is 1).
    - Select a **Sort** order (e.g., Score Descending).
4.  **Generate Report**:
    - Click the **Generate PDF** button.
    - The application will process the data and save the report to:
      `Documents/Grade Visualizer/`
    - Use the **Open Output Folder** button to quickly access your reports.

## Building an Executable

To distribute this application as a standalone executable, you can use PyInstaller. Run the following command in your terminal:

```bash
pyinstaller --noconfirm --onefile --windowed --icon "grade_visualizer_tool.ico" --add-data "grade_visualizer_tool.ico;." --name "GradeVisualizer" grade_visualizer.py
```

*Note: The `--add-data` flag uses a semicolon `;` separator on Windows. On Linux/macOS, use a colon `:` instead.*

## Credits
Developed by George Julius Enock.
