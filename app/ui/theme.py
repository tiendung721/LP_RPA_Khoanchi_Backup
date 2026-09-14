"""Bảng màu và stylesheet dùng chung cho giao diện sáng trên Windows."""

from __future__ import annotations

from PySide6.QtGui import QColor, QPalette
from PySide6.QtWidgets import QApplication

from .feedback import ButtonPressFeedback

PRIMARY = "#2563EB"
PRIMARY_DARK = "#1D4ED8"
PRIMARY_PRESSED = "#1E40AF"
PRIMARY_LIGHT = "#EFF6FF"
TEXT = "#162033"
MUTED_TEXT = "#667085"
BORDER = "#D9E2EC"
SURFACE = "#FFFFFF"
SURFACE_ALT = "#F4F7FB"
SUCCESS = "#15803D"
SUCCESS_BG = "#ECFDF3"
WARNING = "#A16207"
WARNING_BG = "#FFF8DB"
ERROR = "#B42318"
ERROR_BG = "#FFF0F0"


APP_STYLESHEET = f"""
QWidget {{
    color: {TEXT};
    font-family: "Segoe UI";
    font-size: 10pt;
}}
QMainWindow, QDialog, QWidget#applicationRoot {{
    background: {SURFACE_ALT};
}}
QLabel#pageTitle {{
    font-size: 22pt;
    font-weight: 700;
    color: #101828;
}}
QLabel#pageSubtitle, QLabel[muted="true"] {{
    color: {MUTED_TEXT};
}}
QLabel[sectionTitle="true"] {{
    color: #172B4D;
    font-size: 13pt;
    font-weight: 700;
}}
QLabel[cardCategory="true"] {{
    color: {PRIMARY};
    font-size: 8.5pt;
    font-weight: 700;
}}
QLabel[actionTitle="true"] {{
    color: #243B5A;
    font-size: 10.5pt;
    font-weight: 700;
}}
QLabel[status="error"] {{
    color: {ERROR};
}}
QLabel[status="warning"] {{
    color: {WARNING};
}}
QLabel[status="success"] {{
    color: {SUCCESS};
}}
QFrame#excelSummaryHeader {{
    background: {SURFACE};
    border: 1px solid {BORDER};
    border-radius: 12px;
}}
QFrame#excelSummaryHeader[state="success"] {{
    background: {SUCCESS_BG};
    border-color: #B7E4C7;
}}
QFrame#excelSummaryHeader[state="warning"] {{
    background: {WARNING_BG};
    border-color: #E8D58B;
}}
QFrame#excelSummaryHeader[state="error"] {{
    background: {ERROR_BG};
    border-color: #F1B4AE;
}}
QLabel#excelSummaryIcon {{
    color: {PRIMARY};
    background: {PRIMARY_LIGHT};
    border-radius: 21px;
    font-size: 17pt;
    font-weight: 800;
}}
QLabel#excelSummaryIcon[state="success"],
QLabel#excelSummaryIcon[state="no_changes"] {{
    color: {SUCCESS};
    background: #DDF7E8;
}}
QLabel#excelSummaryIcon[state="warning"] {{
    color: {WARNING};
    background: #FDEFB0;
}}
QLabel#excelSummaryIcon[state="error"] {{
    color: {ERROR};
    background: #FFDAD6;
}}
QLabel#excelSummaryTitle {{
    color: #101828;
    font-size: 15pt;
    font-weight: 700;
}}
QLabel#excelSummarySubtitle {{
    color: {MUTED_TEXT};
}}
QScrollArea#excelSummaryScroll,
QWidget#excelSummaryBody {{
    background: transparent;
}}
QFrame#excelSummaryMetric {{
    background: {SURFACE};
    border: 1px solid {BORDER};
    border-radius: 10px;
}}
QFrame#excelSummaryMetric[tone="primary"] {{
    border-top: 3px solid {PRIMARY};
}}
QFrame#excelSummaryMetric[tone="success"] {{
    border-top: 3px solid {SUCCESS};
}}
QFrame#excelSummaryMetric[tone="warning"] {{
    border-top: 3px solid #D69E2E;
}}
QFrame#excelSummaryMetric[tone="neutral"] {{
    border-top: 3px solid #94A3B8;
}}
QLabel#excelSummaryMetricLabel {{
    color: {MUTED_TEXT};
    font-size: 8.5pt;
    font-weight: 700;
}}
QLabel#excelSummaryMetricValue {{
    color: #101828;
    font-size: 20pt;
    font-weight: 750;
}}
QLabel#excelSummaryMetricValue[tone="primary"] {{ color: {PRIMARY_DARK}; }}
QLabel#excelSummaryMetricValue[tone="success"] {{ color: {SUCCESS}; }}
QLabel#excelSummaryMetricValue[tone="warning"] {{ color: {WARNING}; }}
QLabel#excelSummaryMetricNote {{
    color: {MUTED_TEXT};
    font-size: 8.5pt;
}}
QFrame#excelSummaryNotice {{
    background: {PRIMARY_LIGHT};
    border: 1px solid #BFDBFE;
    border-radius: 9px;
}}
QFrame#excelSummaryNotice[tone="success"] {{
    background: {SUCCESS_BG};
    border-color: #B7E4C7;
}}
QFrame#excelSummaryNotice[tone="warning"] {{
    background: {WARNING_BG};
    border-color: #E8D58B;
}}
QFrame#excelSummaryNotice[tone="error"] {{
    background: {ERROR_BG};
    border-color: #F1B4AE;
}}
QLabel#excelSummaryNoticeIcon {{
    color: {PRIMARY};
    font-weight: 800;
}}
QLabel#excelSummaryNoticeIcon[tone="success"] {{ color: {SUCCESS}; }}
QLabel#excelSummaryNoticeIcon[tone="warning"] {{ color: {WARNING}; }}
QLabel#excelSummaryNoticeIcon[tone="error"] {{ color: {ERROR}; }}
QLabel#excelSummarySectionTitle {{
    color: #475467;
    font-size: 8.5pt;
    font-weight: 700;
}}
QToolButton#excelSummaryDetailsToggle {{
    min-height: 36px;
    padding: 0 10px;
    text-align: left;
    color: #344054;
    background: #F8FAFC;
    border: 1px solid {BORDER};
    border-radius: 8px;
    font-weight: 650;
}}
QToolButton#excelSummaryDetailsToggle[tone="warning"] {{
    color: {WARNING};
    background: {WARNING_BG};
    border-color: #E8D58B;
}}
QFrame#excelSummaryDetails {{
    background: {SURFACE};
    border: 1px solid {BORDER};
    border-radius: 8px;
}}
QFrame#excelSummaryDetails[tone="warning"] {{
    background: #FFFCED;
    border-color: #E8D58B;
}}
QFrame#conflictProblemDetails {{
    background: {SURFACE};
    border: 1px solid #BFDBFE;
    border-left: 4px solid {PRIMARY};
    border-radius: 8px;
}}
QLabel#conflictProblemDetailsTitle {{
    color: #172B4D;
    font-size: 10.5pt;
    font-weight: 700;
}}
QLabel#conflictProblemDetailsContext {{
    color: {MUTED_TEXT};
    font-size: 9pt;
}}
QPlainTextEdit#conflictProblemDetailsText {{
    min-height: 0;
    padding: 7px 9px;
    color: {TEXT};
    background: #F8FAFC;
    border: 1px solid #E2E8F0;
    border-radius: 6px;
}}
QToolButton#conflictProblemDetailsCloseButton {{
    min-width: 28px;
    min-height: 28px;
    padding: 0;
    color: {MUTED_TEXT};
    background: transparent;
    border: none;
    border-radius: 5px;
    font-size: 14pt;
}}
QToolButton#conflictProblemDetailsCloseButton:hover {{
    color: {PRIMARY_DARK};
    background: {PRIMARY_LIGHT};
    border: none;
}}
QFrame[card="true"], QGroupBox {{
    background: {SURFACE};
    border: 1px solid {BORDER};
    border-radius: 12px;
}}
QFrame[card="true"] {{
    padding: 1px;
}}
QFrame[actionRow="true"] {{
    background: #F8FAFD;
    border: 1px solid #E1E8F0;
    border-radius: 9px;
}}
QFrame[actionRow="true"]:hover {{
    background: #F4F8FE;
    border-color: #BED1EC;
}}
QFrame[taskRow="true"] {{
    background: transparent;
    border: none;
    border-bottom: 1px solid #E7EDF5;
}}
QFrame[taskRow="true"]:hover {{
    background: #F7FAFE;
}}
QFrame[upcoming="true"] {{
    background: #F1F5F9;
    border: 1px dashed #B9C4D2;
    border-radius: 12px;
}}
QGroupBox {{
    margin-top: 14px;
    padding: 16px 12px 12px 12px;
    font-weight: 600;
}}
QGroupBox::title {{
    subcontrol-origin: margin;
    left: 12px;
    padding: 0 5px;
    color: #334155;
}}
QPushButton, QToolButton {{
    min-height: 34px;
    padding: 1px 15px;
    border: 1px solid #C8D2DF;
    border-radius: 8px;
    background: {SURFACE};
    font-weight: 600;
}}
QPushButton:hover, QToolButton:hover {{
    color: #1D4ED8;
    border-color: #93B4E8;
    background: #F5F9FF;
}}
QPushButton:pressed, QToolButton:pressed {{
    background: #E8EFF8;
    border-color: #7E9FCB;
    padding-top: 3px;
    padding-bottom: 0px;
}}
QPushButton:disabled, QToolButton:disabled {{
    color: #98A2B3;
    background: #F2F4F7;
    border-color: #E1E6ED;
}}
QPushButton[loading="true"], QToolButton[loading="true"],
QPushButton[loading="true"]:disabled, QToolButton[loading="true"]:disabled {{
    color: #1D4ED8;
    background: {PRIMARY_LIGHT};
    border-color: #93B4E8;
}}
QPushButton[primary="true"] {{
    color: white;
    background: {PRIMARY};
    border-color: {PRIMARY};
    font-weight: 700;
}}
QPushButton[primary="true"]:hover {{
    color: white;
    background: {PRIMARY_DARK};
    border-color: {PRIMARY_DARK};
}}
QPushButton[primary="true"]:pressed {{
    color: white;
    background: {PRIMARY_PRESSED};
    border-color: {PRIMARY_PRESSED};
}}
QPushButton[primary="true"][loading="true"],
QPushButton[primary="true"][loading="true"]:disabled {{
    color: white;
    background: #3B82F6;
    border-color: #3B82F6;
}}
QPushButton[link="true"] {{
    min-height: 26px;
    padding: 0px 4px;
    color: #2563EB;
    background: transparent;
    border: none;
    border-radius: 4px;
    font-weight: 600;
}}
QPushButton[link="true"]:hover {{
    color: #1D4ED8;
    background: #EEF4FF;
    border: none;
}}
QPushButton[link="true"]:pressed {{
    color: #1E40AF;
    background: #E4EDFC;
    border: none;
    padding: 1px 4px 0px 4px;
}}
QPushButton[link="true"]:disabled {{
    color: #98A2B3;
    background: transparent;
    border: none;
}}
QPushButton[danger="true"] {{
    color: {ERROR};
    border-color: #F3B3AE;
}}
QPushButton[danger="true"]:hover {{
    color: #912018;
    background: #FFF6F5;
    border-color: #E99892;
}}
QLineEdit, QComboBox, QSpinBox, QDoubleSpinBox, QPlainTextEdit {{
    min-height: 34px;
    padding: 1px 9px;
    border: 1px solid #C8D2DF;
    border-radius: 8px;
    background: {SURFACE};
    selection-background-color: #BFDBFE;
}}
QPlainTextEdit {{
    padding: 8px;
}}
QLineEdit:focus, QComboBox:focus, QSpinBox:focus,
QDoubleSpinBox:focus, QPlainTextEdit:focus {{
    border: 1px solid {PRIMARY};
}}
QLineEdit[invalid="true"], QComboBox[invalid="true"] {{
    border: 1px solid {ERROR};
    background: {ERROR_BG};
}}
QComboBox::drop-down {{
    border: 0;
    width: 24px;
}}
QTableView {{
    background: {SURFACE};
    alternate-background-color: #F8FAFC;
    border: 1px solid {BORDER};
    border-radius: 9px;
    gridline-color: #E8EDF3;
    selection-background-color: #DBEAFE;
    selection-color: {TEXT};
}}
QTableView::item {{
    padding: 5px;
}}
QHeaderView::section {{
    background: #EDF2F7;
    color: #334155;
    padding: 7px 5px;
    border: 0;
    border-right: 1px solid #D7DFE9;
    border-bottom: 1px solid #D7DFE9;
    font-weight: 600;
}}
QListWidget#navigation {{
    border: 0;
    background: transparent;
    color: #DCE6F5;
    outline: 0;
    padding: 10px 7px;
}}
QListWidget#navigation::item {{
    min-height: 40px;
    border-radius: 9px;
    padding: 3px 11px;
    margin: 3px 0;
}}
QListWidget#navigation::item:hover {{
    background: #203A60;
}}
QListWidget#navigation::item:selected {{
    background: {PRIMARY};
    color: white;
    font-weight: 600;
}}
QStatusBar {{
    background: {SURFACE};
    border-top: 1px solid {BORDER};
    color: {MUTED_TEXT};
}}
QMenu {{
    background: {SURFACE};
    border: 1px solid {BORDER};
    border-radius: 8px;
    padding: 6px;
}}
QMenu::item {{
    min-height: 28px;
    border-radius: 6px;
    padding: 2px 18px 2px 10px;
}}
QMenu::item:selected {{
    color: #1D4ED8;
    background: {PRIMARY_LIGHT};
}}
QToolTip {{
    color: white;
    background: #26364D;
    border: 0;
    border-radius: 5px;
    padding: 5px 8px;
}}
QScrollBar:vertical {{
    width: 11px;
    background: transparent;
}}
QScrollBar::handle:vertical {{
    min-height: 24px;
    border-radius: 5px;
    background: #C5CFDB;
}}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{
    height: 0;
}}
"""


def apply_application_theme(application: QApplication) -> None:
    """Áp dụng font, palette và stylesheet nhất quán cho toàn ứng dụng."""

    application.setStyle("Fusion")
    palette = QPalette()
    palette.setColor(QPalette.ColorRole.Window, QColor(SURFACE_ALT))
    palette.setColor(QPalette.ColorRole.WindowText, QColor(TEXT))
    palette.setColor(QPalette.ColorRole.Base, QColor(SURFACE))
    palette.setColor(QPalette.ColorRole.AlternateBase, QColor(SURFACE_ALT))
    palette.setColor(QPalette.ColorRole.Text, QColor(TEXT))
    palette.setColor(QPalette.ColorRole.Button, QColor(SURFACE))
    palette.setColor(QPalette.ColorRole.ButtonText, QColor(TEXT))
    palette.setColor(QPalette.ColorRole.Highlight, QColor(PRIMARY))
    palette.setColor(QPalette.ColorRole.HighlightedText, QColor("#FFFFFF"))
    palette.setColor(QPalette.ColorGroup.Disabled, QPalette.ColorRole.Text, QColor("#98A2B3"))
    application.setPalette(palette)
    application.setStyleSheet(APP_STYLESHEET)
    # Giữ tham chiếu trên QApplication để event filter tồn tại suốt vòng đời app.
    previous_feedback = getattr(application, "_button_press_feedback", None)
    if previous_feedback is not None:
        application.removeEventFilter(previous_feedback)
        previous_feedback.deleteLater()
    feedback = ButtonPressFeedback(application)
    application.installEventFilter(feedback)
    application._button_press_feedback = feedback  # type: ignore[attr-defined]


# Tên ngắn thuận tiện cho mã khởi động và tương thích với các bản tích hợp cũ.
apply_theme = apply_application_theme
