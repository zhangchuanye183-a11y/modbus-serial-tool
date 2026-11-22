# -*- coding: utf-8 -*-
"""
串口调试助手 · Modbus v3.0（单文件版）
===============================================================================
Python 3.9+ / PyQt6 / pyserial（Excel 点表需要 openpyxl）
纯本机运行，无任何网络服务。

能力概览：
  * 串口独立工作线程 SerialWorker(QThread)，UI 线程只做组包/入队，永不阻塞
  * 四页 Tab：通讯监视 / 报文解析 / 地址扫描 / 数据监测
  * Modbus RTU（CRC16）与 Modbus TCP（MBAP）双协议，全页面同步
  * 深浅两套主题实时切换；串口参数、协议、主题、快捷指令、点表本地持久化
  * 串口热插拔自动检测并提示；接收区带毫秒时间戳、收发分色、CRC 错误红色标记
  * 定时发送、自动 CRC、快捷指令、总线质量统计
  * 报文解析 + CRC 工具 + ASCII/HEX 互转
  * 从机地址扫描（1~247）
  * 多从机寄存器轮询、Excel 点表导入、数据导出 Excel
"""

from __future__ import annotations

import json
import queue
import re
import struct
import sys
import threading
import time
from datetime import datetime
from pathlib import Path

import serial
from serial.tools import list_ports

from PyQt6.QtCore import Qt, QThread, QTimer, pyqtSignal
from PyQt6.QtGui import (
    QColor,
    QFont,
    QIcon,
    QKeySequence,
    QPalette,
    QShortcut,
    QTextCharFormat,
    QTextCursor,
)
from PyQt6.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QButtonGroup,
    QCheckBox,
    QComboBox,
    QDialog,
    QFileDialog,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QListWidget,
    QMainWindow,
    QMessageBox,
    QPlainTextEdit,
    QPushButton,
    QRadioButton,
    QScrollArea,
    QSizePolicy,
    QSpinBox,
    QSplitter,
    QStatusBar,
    QTableWidget,
    QTableWidgetItem,
    QTabWidget,
    QVBoxLayout,
    QWidget,
)

mc = sys.modules[__name__]

try:
    from openpyxl import Workbook, load_workbook
except Exception:  # 未安装 openpyxl 时，仅影响点表导入/导出
    Workbook = None
    load_workbook = None


DARK = {
    "bg": "#0A0D13",
    "card": "#121821",
    "card2": "#0E141C",
    "border": "#1E2836",
    "border2": "#2A3646",
    "text": "#E8EEF6",
    "muted": "#94A5B8",
    "faint": "#5D6E82",
    "input": "#0C121A",
    "view": "#090D13",
    "accent": "#2DD4BF",
    "accent_hover": "#45E0CC",
    "accent_text": "#04211D",
    "sel": "#134E4A",
    "sel_text": "#CCFBF1",
    "danger": "#F87171",
    "ok": "#34D399",
    "warn": "#FBBF24",
    "pill_bg": "#1A2431",
    "pill_on": "#10352C",
    "pill_busy": "#33290F",
    "rx": "#7DD3FC",
    "tx": "#FCD34D",
    "hover_bg": "#182230",
    "accent_soft": "rgba(45, 212, 191, 34)",
    "border_soft": "rgba(45, 212, 191, 60)",
}

LIGHT = {
    "bg": "#F1F4F9",
    "card": "#FFFFFF",
    "card2": "#F7F9FC",
    "border": "#E4E9F1",
    "border2": "#CBD5E1",
    "text": "#111A2B",
    "muted": "#5A6B80",
    "faint": "#93A1B4",
    "input": "#FFFFFF",
    "view": "#FFFFFF",
    "accent": "#0D9488",
    "accent_hover": "#0F766E",
    "accent_text": "#FFFFFF",
    "sel": "#CCFBF1",
    "sel_text": "#0F766E",
    "danger": "#DC2626",
    "ok": "#15803D",
    "warn": "#B45309",
    "pill_bg": "#EEF2F8",
    "pill_on": "#DCFCE7",
    "pill_busy": "#FEF3C7",
    "rx": "#0369A1",
    "tx": "#B45309",
    "hover_bg": "#EFF4FA",
    "accent_soft": "rgba(13, 148, 136, 24)",
    "border_soft": "rgba(13, 148, 136, 48)",
}

TEMPLATE = """
QWidget {
    color: @text;
    font-size: 13px;
}
QWidget#root, QMainWindow, QDialog {
    background-color: @bg;
}

/* ------------- 卡片 -------------- */
QFrame#card {
    background-color: @card;
    border: 1px solid @border;
    border-radius: 12px;
}

/* ------------- 文字层级 ------------- */
QLabel#appTitle {
    font-size: 18px;
    font-weight: 700;
    color: @text;
}
QLabel#appSubtitle {
    font-size: 11px;
    color: @faint;
}
QLabel#appMark {
    background-color: @accent;
    color: @accent_text;
    border-radius: 10px;
    font-size: 20px;
    font-weight: 800;
}
QLabel#sectionTitle {
    font-size: 13px;
    font-weight: 600;
    color: @text;
}
QLabel#fieldLabel {
    font-size: 12px;
    color: @muted;
}
QLabel#hint {
    font-size: 12px;
    color: @faint;
}
QLabel#statePill {
    font-size: 12px;
    padding: 3px 12px;
    border-radius: 11px;
    background-color: @pill_bg;
    color: @muted;
    border: 1px solid @border;
}
QLabel#statePill[state="on"] {
    background-color: @pill_on;
    color: @ok;
    border: 1px solid @ok;
}
QLabel#statePill[state="busy"] {
    background-color: @pill_busy;
    color: @warn;
    border: 1px solid @warn;
}

/* ------------- 标签页 -------------- */
QTabWidget::pane {
    border: none;
    background: transparent;
    margin-top: 8px;
}
QTabBar {
    background: transparent;
}
QTabBar::tab {
    background-color: transparent;
    border: 1px solid @border;
    border-radius: 9px;
    color: @muted;
    padding: 8px 22px;
    margin-right: 6px;
    font-size: 13px;
}
QTabBar::tab:hover:!selected {
    background-color: @hover_bg;
    border-color: @border2;
    color: @text;
}
QTabBar::tab:selected {
    background-color: @accent;
    border-color: @accent;
    color: @accent_text;
    font-weight: 600;
}

/* ------------- 输入控件 ------------- */
QComboBox, QSpinBox, QLineEdit {
    background-color: @input;
    border: 1px solid @border2;
    border-radius: 8px;
    padding: 5px 10px;
    min-height: 21px;
    color: @text;
}
QSpinBox {
    padding-right: 22px;
}
QComboBox {
    padding-right: 26px;
}
QSpinBox::up-button, QSpinBox::down-button {
    subcontrol-origin: border;
    subcontrol-position: top right;
    width: 18px;
    border: none;
    background: transparent;
}
QSpinBox::down-button {
    subcontrol-position: bottom right;
}
QSpinBox::up-arrow {
    image: url("@arrow_up");
    width: 9px;
    height: 9px;
}
QSpinBox::down-arrow {
    image: url("@arrow_down");
    width: 9px;
    height: 9px;
}
QComboBox:hover, QSpinBox:hover, QLineEdit:hover {
    border-color: @border_soft;
}
QComboBox:focus, QSpinBox:focus, QLineEdit:focus {
    border-color: @accent;
}
QComboBox:disabled, QSpinBox:disabled, QLineEdit:disabled {
    background-color: @card2;
    color: @faint;
    border-color: @border;
}
QComboBox::drop-down {
    border: none;
    width: 22px;
    background: transparent;
}
QComboBox::down-arrow {
    image: url("@arrow_down");
    width: 11px;
    height: 11px;
}
QComboBox QAbstractItemView {
    background-color: @card;
    border: 1px solid @border2;
    border-radius: 10px;
    padding: 4px;
    color: @text;
    outline: none;
    selection-background-color: @accent;
    selection-color: @accent_text;
}

/* ------------- 文本区 ------------- */
QPlainTextEdit {
    background-color: @view;
    border: 1px solid @border;
    border-radius: 10px;
    padding: 8px;
    color: @text;
    selection-background-color: @sel;
    selection-color: @sel_text;
}
QPlainTextEdit:focus {
    border-color: @border_soft;
}

/* ------------- 按钮 ------------- */
QPushButton {
    background-color: @card2;
    border: 1px solid @border2;
    border-radius: 8px;
    padding: 7px 16px;
    color: @text;
    min-height: 20px;
}
QPushButton:hover {
    background-color: @hover_bg;
    border-color: @border_soft;
    color: @text;
}
QPushButton:pressed {
    background-color: @input;
}
QPushButton:disabled {
    color: @faint;
    border-color: @border;
    background-color: @card2;
}
QPushButton#primary {
    background-color: @accent;
    border: 1px solid @accent;
    color: @accent_text;
    font-weight: 600;
    font-size: 13px;
}
QPushButton#primary:hover {
    background-color: @accent_hover;
    border-color: @accent_hover;
    color: @accent_text;
}
QPushButton#primary:pressed {
    background-color: @accent;
}
QPushButton#primary:disabled {
    background-color: @border2;
    border-color: @border2;
    color: @faint;
}
QPushButton#primary[danger="true"] {
    background-color: @danger;
    border-color: @danger;
    color: #ffffff;
}
QPushButton#ghost {
    background-color: transparent;
    border: 1px solid @border2;
    color: @muted;
}
QPushButton#ghost:hover {
    background-color: @hover_bg;
    border-color: @accent;
    color: @text;
}

/* ------------- 分段选择（ASCII / HEX） ------------- */
QRadioButton[seg="true"] {
    background-color: @card2;
    border: 1px solid @border2;
    border-radius: 8px;
    padding: 5px 16px;
    color: @muted;
    font-size: 12px;
}
QRadioButton[seg="true"]:hover {
    border-color: @border_soft;
    color: @text;
}
QRadioButton[seg="true"]:checked {
    background-color: @accent;
    border-color: @accent;
    color: @accent_text;
    font-weight: 600;
}
QRadioButton::indicator {
    width: 0px;
    height: 0px;
    margin: 0px;
    border: none;
    background: transparent;
}

/* ------------- 复选框 ------------- */
QCheckBox {
    color: @muted;
    spacing: 8px;
    font-size: 12px;
}
QCheckBox:hover {
    color: @text;
}
QCheckBox::indicator {
    width: 15px;
    height: 15px;
    border-radius: 4px;
    border: 1px solid @border2;
    background-color: @input;
}
QCheckBox::indicator:hover {
    border-color: @accent;
}
QCheckBox::indicator:checked {
    background-color: @accent;
    border-color: @accent;
    image: url("@check");
}
QCheckBox::indicator:checked:hover {
    background-color: @accent_hover;
    border-color: @accent_hover;
}
QCheckBox::indicator:disabled {
    background-color: @card2;
    border-color: @border;
}
QCheckBox::indicator:checked:disabled {
    background-color: @border2;
    border-color: @border2;
    image: url("@check_dim");
}

/* ------------- 表格 / 列表 ------------- */
QTableWidget, QListWidget {
    background-color: @view;
    alternate-background-color: @card2;
    color: @text;
    border: 1px solid @border;
    border-radius: 10px;
    gridline-color: @border;
    selection-background-color: @sel;
    selection-color: @sel_text;
    outline: none;
}
QTableWidget::item {
    padding: 2px 8px;
    border: none;
}
QTableWidget::item:selected, QListWidget::item:selected {
    background-color: @sel;
    color: @sel_text;
}
QTableWidget::item:hover, QListWidget::item:hover {
    background-color: @hover_bg;
}
QListWidget::item {
    padding: 5px 8px;
    border-radius: 6px;
}
QHeaderView::section {
    background-color: @card;
    color: @muted;
    border: none;
    border-bottom: 1px solid @border;
    padding: 8px 10px;
    font-size: 12px;
    font-weight: 600;
}
QTableCornerButton::section {
    background-color: @card;
    border: none;
}

/* ------------- 滚动条 ------------- */
QScrollBar:vertical {
    background: transparent;
    width: 10px;
    margin: 2px 2px 2px 0;
}
QScrollBar::handle:vertical {
    background: @border2;
    min-height: 30px;
    border-radius: 4px;
}
QScrollBar::handle:vertical:hover {
    background: @accent;
}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
    height: 0;
}
QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {
    background: none;
}
QScrollBar:horizontal {
    background: transparent;
    height: 10px;
    margin: 0 2px 2px 2px;
}
QScrollBar::handle:horizontal {
    background: @border2;
    min-width: 30px;
    border-radius: 4px;
}
QScrollBar::handle:horizontal:hover {
    background: @accent;
}
QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {
    width: 0;
}
QScrollBar::add-page:horizontal, QScrollBar::sub-page:horizontal {
    background: none;
}

/* ------------- 分隔条 / 滚动区域 ------------- */
QSplitter::handle {
    background-color: transparent;
}
QSplitter::handle:hover {
    background-color: @border_soft;
}
QScrollArea {
    background: transparent;
    border: none;
}

/* ------------- 提示 / 菜单 ------------- */
QToolTip {
    background-color: @card;
    color: @text;
    border: 1px solid @border2;
    border-radius: 6px;
    padding: 6px 8px;
}
QMenu {
    background-color: @card;
    border: 1px solid @border2;
    border-radius: 10px;
    padding: 6px;
}
QMenu::item {
    padding: 7px 18px;
    border-radius: 6px;
    color: @text;
}
QMenu::item:selected {
    background-color: @accent_soft;
    color: @text;
}
QMenu::separator {
    height: 1px;
    background: @border;
    margin: 4px 8px;
}

/* ------------- 状态栏 ------------- */
QStatusBar {
    background-color: @card;
    border-top: 1px solid @border;
    min-height: 28px;
}
QStatusBar::item {
    border: none;
}
QLabel#statusText {
    color: @muted;
    font-size: 12px;
    padding-left: 10px;
}
QLabel#counterText {
    color: @faint;
    font-size: 12px;
    padding-right: 12px;
}
QMessageBox, QMessageBox QLabel {
    background-color: @bg;
    color: @text;
}
"""



def colors(theme: str) -> dict[str, str]:
    return LIGHT if theme == "light" else DARK


_ARROW_CACHE: dict[str, str] = {}


def _arrow_png(color: str, direction: str) -> str:
    """按主题色即时生成下拉/上下箭头图标，返回 QSS url() 可用的路径"""
    key = f"{direction}-{color}"
    cached = _ARROW_CACHE.get(key)
    if cached and Path(cached).exists():
        return cached

    import tempfile

    from PyQt6.QtCore import QPointF
    from PyQt6.QtGui import QPainter, QPen, QPixmap, QPolygonF

    size = 20
    pixmap = QPixmap(size, size)
    pixmap.fill(Qt.GlobalColor.transparent)
    painter = QPainter(pixmap)
    painter.setRenderHint(QPainter.RenderHint.Antialiasing, True)
    pen = QPen(QColor(color))
    pen.setWidthF(3.2)
    pen.setCapStyle(Qt.PenCapStyle.RoundCap)
    pen.setJoinStyle(Qt.PenJoinStyle.RoundJoin)
    painter.setPen(pen)
    left, right = 4.5, size - 4.5
    mid = size / 2.0
    span = 3.6
    if direction == "down":
        points = [(left, mid - span), (mid, mid + span), (right, mid - span)]
    else:
        points = [(left, mid + span), (mid, mid - span), (right, mid + span)]
    painter.drawPolyline(QPolygonF([QPointF(x, y) for x, y in points]))
    painter.end()

    folder = Path(tempfile.gettempdir()) / "serial_debug_assets"
    folder.mkdir(parents=True, exist_ok=True)
    target = folder / f"arrow_{direction}_{color.lstrip('#')}.png"
    pixmap.save(str(target), "PNG")
    path = target.as_posix()
    _ARROW_CACHE[key] = path
    return path


def _check_png(color: str) -> str:
    """生成复选框对勾图标"""
    key = f"check-{color}"
    cached = _ARROW_CACHE.get(key)
    if cached and Path(cached).exists():
        return cached

    import tempfile

    from PyQt6.QtCore import QPointF
    from PyQt6.QtGui import QPainter, QPen, QPixmap, QPolygonF

    size = 20
    pixmap = QPixmap(size, size)
    pixmap.fill(Qt.GlobalColor.transparent)
    painter = QPainter(pixmap)
    painter.setRenderHint(QPainter.RenderHint.Antialiasing, True)
    pen = QPen(QColor(color))
    pen.setWidthF(3.4)
    pen.setCapStyle(Qt.PenCapStyle.RoundCap)
    pen.setJoinStyle(Qt.PenJoinStyle.RoundJoin)
    painter.setPen(pen)
    painter.drawPolyline(QPolygonF([
        QPointF(size * 0.20, size * 0.52),
        QPointF(size * 0.42, size * 0.73),
        QPointF(size * 0.80, size * 0.26),
    ]))
    painter.end()

    folder = Path(tempfile.gettempdir()) / "serial_debug_assets"
    folder.mkdir(parents=True, exist_ok=True)
    target = folder / f"check_{color.lstrip('#')}.png"
    pixmap.save(str(target), "PNG")
    path = target.as_posix()
    _ARROW_CACHE[key] = path
    return path


def build_qss(theme: str) -> str:
    data = colors(theme)
    qss = TEMPLATE
    for key in sorted(data, key=len, reverse=True):
        qss = qss.replace("@" + key, data[key])
    qss = qss.replace("@arrow_down", _arrow_png(data["muted"], "down"))
    qss = qss.replace("@arrow_up", _arrow_png(data["muted"], "up"))
    qss = qss.replace("@check_dim", _check_png(data["muted"]))
    qss = qss.replace("@check", _check_png(data["accent_text"]))
    return qss


def build_palette(theme: str) -> QPalette:
    c = colors(theme)
    palette = QPalette()
    palette.setColor(QPalette.ColorRole.Window, QColor(c["bg"]))
    palette.setColor(QPalette.ColorRole.WindowText, QColor(c["text"]))
    palette.setColor(QPalette.ColorRole.Base, QColor(c["view"]))
    palette.setColor(QPalette.ColorRole.AlternateBase, QColor(c["card2"]))
    palette.setColor(QPalette.ColorRole.Mid, QColor(c["border2"]))
    palette.setColor(QPalette.ColorRole.Midlight, QColor(c["border"]))
    palette.setColor(QPalette.ColorRole.Dark, QColor(c["border"]))
    palette.setColor(QPalette.ColorRole.Light, QColor(c["card"]))
    palette.setColor(QPalette.ColorRole.Shadow, QColor(c["faint"]))
    palette.setColor(QPalette.ColorRole.Text, QColor(c["text"]))
    palette.setColor(QPalette.ColorRole.Button, QColor(c["card2"]))
    palette.setColor(QPalette.ColorRole.ButtonText, QColor(c["text"]))
    palette.setColor(QPalette.ColorRole.Highlight, QColor(c["sel"]))
    palette.setColor(QPalette.ColorRole.HighlightedText, QColor(c["sel_text"]))
    palette.setColor(QPalette.ColorRole.ToolTipBase, QColor(c["card"]))
    palette.setColor(QPalette.ColorRole.ToolTipText, QColor(c["text"]))
    palette.setColor(QPalette.ColorRole.PlaceholderText, QColor(c["faint"]))
    palette.setColor(QPalette.ColorGroup.Disabled, QPalette.ColorRole.Text, QColor(c["faint"]))
    palette.setColor(QPalette.ColorGroup.Disabled, QPalette.ColorRole.ButtonText, QColor(c["faint"]))
    return palette


FUNCTION_NAMES = {
    1: "读线圈",
    2: "读离散输入",
    3: "读保持寄存器",
    4: "读输入寄存器",
    5: "写单个线圈",
    6: "写单个寄存器",
    15: "写多个线圈",
    16: "写多个寄存器",
    23: "读写多个寄存器",
}

EXCEPTION_TEXT = {
    1: "非法功能码",
    2: "非法数据地址",
    3: "非法数据值",
    4: "从机设备故障",
    5: "确认（处理中）",
    6: "从机设备忙",
    8: "存储奇偶校验错",
    10: "网关路径不可用",
    11: "网关目标设备无响应",
}

_SEP = re.compile(r"[\s,;:，；、|\-]+")
_TOK = re.compile(r"^(?:0[xX])?([0-9A-Fa-f]+)$")


def parse_hex_text(text: str) -> bytes:
    tokens = [t for t in _SEP.split(text.strip()) if t]
    if not tokens:
        raise ValueError("没有可用的十六进制数据")
    result = bytearray()
    for token in tokens:
        match = _TOK.match(token)
        if not match:
            raise ValueError(f"无法识别的片段 {token!r}")
        digits = match.group(1)
        if len(digits) % 2:
            raise ValueError(f"片段 {token!r} 的十六进制位数为奇数")
        result.extend(bytes.fromhex(digits))
    return bytes(result)


def bytes_to_hex(data: bytes) -> str:
    return data.hex(" ").upper()


def text_to_hex(text: str, encoding: str = "utf-8") -> str:
    return bytes_to_hex(text.encode(encoding))


def hex_to_text(text: str, encoding: str = "utf-8") -> str:
    return parse_hex_text(text).decode(encoding, errors="replace")


def crc16(data: bytes) -> int:
    crc = 0xFFFF
    for byte in data:
        crc ^= byte
        for _ in range(8):
            crc = (crc >> 1) ^ 0xA001 if crc & 1 else crc >> 1
    return crc & 0xFFFF


def with_crc(frame: bytes) -> bytes:
    value = crc16(frame)
    return frame + bytes((value & 0xFF, (value >> 8) & 0xFF))


def crc_check(frame: bytes) -> tuple[bool, int, int]:
    if len(frame) < 3:
        return False, 0, 0
    expected = crc16(frame[:-2])
    actual = frame[-2] | (frame[-1] << 8)
    return expected == actual, expected, actual


def mbap_build(pdu: bytes, unit: int, tid: int) -> bytes:
    return struct.pack(">HHHB", tid & 0xFFFF, 0, len(pdu) + 1, unit & 0xFF) + pdu


def mbap_parse(data: bytes):
    if len(data) < 7:
        return None
    return struct.unpack(">HHHB", data[:7])


def exception_name(code: int) -> str:
    return EXCEPTION_TEXT.get(code, f"未知异常码 {code}")


def build_read_request(protocol: str, unit: int, func: int, start: int, count: int, tid: int) -> bytes:
    pdu = bytes((
        func & 0xFF,
        (start >> 8) & 0xFF, start & 0xFF,
        (count >> 8) & 0xFF, count & 0xFF,
    ))
    if protocol == "tcp":
        return mbap_build(pdu, unit, tid)
    return with_crc(bytes((unit & 0xFF,)) + pdu)


def decode_read_response(protocol: str, frame: bytes, unit: int, func: int, count: int) -> dict:
    if protocol == "tcp":
        header = mbap_parse(frame)
        if header is None:
            return {"ok": False, "error": "MBAP 头不完整"}
        _tid, pid, length, uid = header
        if pid != 0:
            return {"ok": False, "error": f"协议标识错误（{pid}）"}
        if uid != unit:
            return {"ok": False, "error": f"单元标识不匹配（期望 {unit}，实际 {uid}）"}
        total = 6 + length
        if len(frame) < total:
            return {"ok": False, "error": "响应长度不足"}
        pdu = frame[7:total]
        if not pdu:
            return {"ok": False, "error": "PDU 为空"}
        fc = pdu[0]
        if fc & 0x80:
            code = pdu[1] if len(pdu) > 1 else 0
            return {"ok": False, "exception": code, "error": f"从机异常：{exception_name(code)}"}
        if fc != func:
            return {"ok": False, "error": f"功能码不匹配（期望 {func:02X}，实际 {fc:02X}）"}
        if len(pdu) < 2:
            return {"ok": False, "error": "PDU 数据不足"}
        byte_count = pdu[1]
        if byte_count != count * 2:
            return {"ok": False, "error": f"字节数不符（期望 {count * 2}，实际 {byte_count}）"}
        return {"ok": True, "registers": _registers(pdu[2:2 + byte_count])}

    if len(frame) < 5:
        return {"ok": False, "error": "报文长度不足"}
    if frame[0] != unit:
        return {"ok": False, "error": f"从站地址不匹配（期望 {unit}，实际 {frame[0]}）"}
    passed, expected, actual = crc_check(frame)
    if not passed:
        return {"ok": False, "error": f"CRC 校验失败（计算 {expected:04X}，报文 {actual:04X}）"}
    fc = frame[1]
    if fc & 0x80:
        return {"ok": False, "exception": frame[2], "error": f"从机异常：{exception_name(frame[2])}"}
    if fc != func:
        return {"ok": False, "error": f"功能码不匹配（期望 {func:02X}，实际 {fc:02X}）"}
    byte_count = frame[2]
    if byte_count != count * 2:
        return {"ok": False, "error": f"字节数不符（期望 {count * 2}，实际 {byte_count}）"}
    return {"ok": True, "registers": _registers(frame[3:3 + byte_count])}


def _registers(data: bytes) -> list[int]:
    return [(data[i] << 8) | data[i + 1] for i in range(0, len(data) - 1, 2)]


def extract_response(protocol: str, buffer: bytes, unit: int, func: int, tid: int | None = None):
    if protocol == "tcp":
        index = 0
        while index + 8 <= len(buffer):
            pid = (buffer[index + 2] << 8) | buffer[index + 3]
            if pid != 0:
                index += 1
                continue
            length = (buffer[index + 4] << 8) | buffer[index + 5]
            total = 6 + length
            if total < 8 or total > 520:
                index += 1
                continue
            if index + total > len(buffer):
                break
            frame = buffer[index:index + total]
            if frame[6] != unit:
                index += 1
                continue
            if tid is not None and ((frame[0] << 8) | frame[1]) != tid:
                index += 1
                continue
            return frame, index + total
        return None, 0

    index = 0
    while index + 5 <= len(buffer):
        if buffer[index] != unit:
            index += 1
            continue
        fc = buffer[index + 1]
        if fc == (func | 0x80):
            total = 5
        elif fc == func:
            total = 5 + buffer[index + 2]
        else:
            index += 1
            continue
        if total > 520:
            index += 1
            continue
        if index + total > len(buffer):
            break
        if crc_check(buffer[index:index + total])[0]:
            return buffer[index:index + total], index + total
        index += 1
    return None, 0


def looks_modbus(protocol: str, data: bytes) -> bool:
    if protocol == "tcp":
        if len(data) < 8:
            return False
        return ((data[2] << 8) | data[3]) == 0
    if len(data) < 4:
        return False
    func = data[1] & 0x7F
    return 1 <= data[0] <= 247 and func in (1, 2, 3, 4, 5, 6, 15, 16, 23)


def split_frames(protocol: str, buffer: bytes):
    frames: list[tuple[bytes, bool, str]] = []
    if protocol == "tcp":
        index = 0
        while index + 8 <= len(buffer):
            pid = (buffer[index + 2] << 8) | buffer[index + 3]
            if pid != 0:
                index += 1
                continue
            length = (buffer[index + 4] << 8) | buffer[index + 5]
            total = 6 + length
            if total < 8 or total > 520:
                index += 1
                continue
            if index + total > len(buffer):
                break
            frames.append((buffer[index:index + total], True, ""))
            index += total
        return frames, index

    index = 0
    size = len(buffer)
    while index + 5 <= size:
        func = buffer[index + 1]
        if func & 0x80:
            total = 5
        elif func in (1, 2, 3, 4):
            total = 5 + buffer[index + 2]
        elif func in (5, 6, 15, 16):
            total = 8
        else:
            index += 1
            continue
        if total > 520:
            index += 1
            continue
        if index + total > size:
            break
        frame = buffer[index:index + total]
        if crc_check(frame)[0]:
            frames.append((frame, True, ""))
        else:
            frames.append((frame, False, "CRC 校验失败"))
        index += total
    return frames, index


def format_registers(registers: list[int], data_type: str) -> list[str]:
    if data_type == "INT16":
        return [str(v - 65536 if v > 32767 else v) for v in registers]
    if data_type == "UINT16":
        return [str(v) for v in registers]
    if data_type in ("INT32", "UINT32", "FLOAT"):
        out = []
        for i in range(0, len(registers) - 1, 2):
            raw = (registers[i] << 16) | registers[i + 1]
            if data_type == "FLOAT":
                out.append(f"{struct.unpack('>f', struct.pack('>I', raw))[0]:.4f}")
            elif data_type == "INT32":
                out.append(str(raw - 2 ** 32 if raw >= 2 ** 31 else raw))
            else:
                out.append(str(raw))
        return out
    if data_type == "BIT":
        return ["".join(f"{v:016b}" for v in registers[:1])]
    return [str(v) for v in registers]


def format_scalar(value: int, fmt: str) -> str:
    if fmt == "有符号":
        return str(value - 65536 if value > 32767 else value)
    if fmt == "十六进制":
        return f"0x{value:04X}"
    return str(value)


def _describe_pdu(pdu: bytes, message_type: str, data_type: str) -> list[str]:
    if not pdu:
        return ["PDU 为空"]
    func = pdu[0]
    lines = [f"功能码：{func:02X}（{FUNCTION_NAMES.get(func & 0x7F, '未知功能码')}）"]
    if func & 0x80:
        code = pdu[1] if len(pdu) > 1 else 0
        lines.append(f"异常应答：{code:02X} —— {exception_name(code)}")
        return lines

    if func in (1, 2):
        if message_type == "主机指令" or len(pdu) == 5:
            start = (pdu[1] << 8) | pdu[2]
            count = (pdu[3] << 8) | pdu[4]
            lines.append(f"起始地址：{start}（0x{start:04X}）")
            lines.append(f"数量：{count}")
            return lines
        if len(pdu) < 2:
            lines.append("数据不足")
            return lines
        byte_count = pdu[1]
        data = pdu[2:2 + byte_count]
        bits = []
        for byte in data:
            bits.extend((byte >> bit) & 1 for bit in range(8))
        lines.append(f"数据长度：{byte_count} 字节")
        lines.append("位状态：" + " ".join(str(bit) for bit in bits))
        return lines

    if func in (3, 4):
        if message_type == "主机指令" or len(pdu) == 5:
            start = (pdu[1] << 8) | pdu[2]
            count = (pdu[3] << 8) | pdu[4]
            lines.append(f"起始地址：{start}（0x{start:04X}）")
            lines.append(f"寄存器数量：{count}")
            return lines
        if len(pdu) < 2:
            lines.append("数据不足")
            return lines
        byte_count = pdu[1]
        registers = _registers(pdu[2:2 + byte_count])
        values = format_registers(registers, data_type)
        lines.append(f"数据长度：{byte_count} 字节（{len(registers)} 个寄存器）")
        for index, text in enumerate(values):
            lines.append(f"  [{index}] {text}")
        return lines

    if func in (5, 6):
        if len(pdu) < 5:
            lines.append("数据不足")
            return lines
        address = (pdu[1] << 8) | pdu[2]
        value = (pdu[3] << 8) | pdu[4]
        lines.append(f"地址：{address}（0x{address:04X}）")
        if func == 5:
            lines.append(f"值：0x{value:04X}（{'ON' if value == 0xFF00 else 'OFF'}）")
        else:
            lines.append(f"值：{value}（0x{value:04X}）")
        return lines

    if func in (15, 16):
        if len(pdu) < 5:
            lines.append("数据不足")
            return lines
        start = (pdu[1] << 8) | pdu[2]
        quantity = (pdu[3] << 8) | pdu[4]
        lines.append(f"起始地址：{start}（0x{start:04X}）")
        lines.append(f"数量：{quantity}")
        if len(pdu) >= 6:
            byte_count = pdu[5]
            data = pdu[6:6 + byte_count]
            lines.append(f"数据长度：{byte_count} 字节")
            if func == 16:
                for index, text in enumerate(format_registers(_registers(data), data_type)):
                    lines.append(f"  [{index}] {text}")
            else:
                lines.append(f"线圈数据：{bytes_to_hex(data)}")
        return lines

    lines.append(f"数据：{bytes_to_hex(pdu[1:])}")
    return lines


def describe_frame(protocol: str, data: bytes, message_type: str = "从机应答",
                   data_type: str = "UINT16") -> str:
    lines = [f"报文长度：{len(data)} 字节", f"原始数据：{bytes_to_hex(data)}"]
    if not data:
        return "报文为空"

    if protocol == "tcp":
        header = mbap_parse(data)
        if header is None or header[1] != 0:
            lines.append("")
            lines.append("MBAP 校验：失败（未检测到协议标识 0，按 PDU 解析）")
            lines.append("")
            lines.extend(_describe_pdu(data, message_type, data_type))
            return "\n".join(lines)
        tid, pid, length, unit = header
        declared = 6 + length
        lines.append("")
        lines.append("—— MBAP 头 ——")
        lines.append(f"事务标识：{tid}（0x{tid:04X}）")
        lines.append(f"协议标识：{pid}")
        lines.append(f"长度字段：{length}")
        lines.append(f"单元标识：{unit}")
        if len(data) == declared:
            lines.append("MBAP 校验：通过")
        else:
            lines.append(f"MBAP 校验：失败（声明 {declared} 字节，实际 {len(data)} 字节）")
        lines.append("")
        lines.append("—— PDU ——")
        lines.extend(_describe_pdu(data[7:declared], message_type, data_type))
        return "\n".join(lines)

    if len(data) < 4:
        lines.append("RTU 报文过短（至少 4 字节）")
        return "\n".join(lines)
    passed, expected, actual = crc_check(data)
    lines.append("")
    lines.append(f"从机地址：{data[0]}")
    lines.append(f"CRC16 校验：{'通过' if passed else '失败'}")
    lines.append(f"CRC 计算值：0x{expected:04X}  报文内：0x{actual:04X}")
    lines.append("")
    lines.append("—— PDU ——")
    lines.extend(_describe_pdu(data[1:-2], message_type, data_type))
    return "\n".join(lines)


APP_TITLE = "串口调试助手 · Modbus v3.1"
if getattr(sys, "frozen", False):
    CONFIG_PATH = Path(sys.executable).resolve().with_name("config.json")
else:
    CONFIG_PATH = Path(__file__).resolve().with_name("config.json")

BAUDRATE_LIST = [1200, 2400, 4800, 9600, 19200, 38400, 57600, 115200, 230400]
DATABITS_ITEMS = [("7", 7), ("8", 8)]
PARITY_ITEMS = [("无校验 (N)", "N"), ("奇校验 (O)", "O"), ("偶校验 (E)", "E")]
STOPBITS_ITEMS = [("1", 1), ("2", 2)]
ENCODING_ITEMS = [("UTF-8", "utf-8"), ("GBK", "gbk")]
PROTOCOL_ITEMS = [("Modbus RTU", "rtu"), ("Modbus TCP", "tcp")]
FUNCTION_ITEMS = [("03 读保持寄存器", 3), ("04 读输入寄存器", 4)]
VALUE_FORMATS = ["无符号", "有符号", "十六进制"]
DATA_TYPES = ["BIT", "INT16", "UINT16", "INT32", "UINT32", "FLOAT"]

DEFAULTS = {
    "serial": {"port": "", "baudrate": 9600, "bytesize": 8, "parity": "N", "stopbits": 1, "station": 1},
    "protocol": "rtu",
    "theme": "dark",
    "quick": [{"name": "读保持 0-9", "hex": "01 03 00 00 00 0A"}],
    "poll": {"slaves": [1], "points": []},
}


def now_stamp() -> str:
    return datetime.now().strftime("%H:%M:%S.%f")[:-3]


# ---------------------------------------------------------------------------
# 本地配置
# ---------------------------------------------------------------------------
class Config:
    def __init__(self) -> None:
        self.path = CONFIG_PATH
        self.data = json.loads(json.dumps(DEFAULTS))
        self.load()

    def load(self) -> None:
        try:
            if self.path.exists():
                loaded = json.loads(self.path.read_text("utf-8"))
                self._merge(self.data, loaded)
        except Exception:
            pass

    @staticmethod
    def _merge(base: dict, extra: dict) -> None:
        for key, value in extra.items():
            if isinstance(value, dict) and isinstance(base.get(key), dict):
                Config._merge(base[key], value)
            else:
                base[key] = value

    def save(self) -> None:
        try:
            self.path.write_text(json.dumps(self.data, ensure_ascii=False, indent=2), "utf-8")
        except Exception:
            pass


# ---------------------------------------------------------------------------
# 串口线程
# ---------------------------------------------------------------------------
class SerialWorker(QThread):
    opened = pyqtSignal()
    open_failed = pyqtSignal(str)
    data_received = pyqtSignal(bytes)
    bytes_sent = pyqtSignal(int)
    error = pyqtSignal(str)

    def __init__(self, config: dict) -> None:
        super().__init__()
        self._config = config
        self._tx: queue.Queue[bytes] = queue.Queue()
        self._stop = threading.Event()
        self._serial: serial.Serial | None = None

    def send(self, payload: bytes) -> bool:
        if not payload or self._stop.is_set():
            return False
        self._tx.put(bytes(payload))
        return True

    def stop(self) -> None:
        self._stop.set()

    def run(self) -> None:
        cfg = self._config
        try:
            port = serial.Serial(
                port=cfg["port"],
                baudrate=cfg["baudrate"],
                bytesize=cfg["bytesize"],
                parity=cfg["parity"],
                stopbits=cfg["stopbits"],
                timeout=0.02,
                write_timeout=1.0,
            )
        except (serial.SerialException, ValueError, OSError) as exc:
            self.open_failed.emit(str(exc))
            return

        try:
            try:
                port.reset_input_buffer()
            except (serial.SerialException, OSError):
                pass
            self.opened.emit()
            while not self._stop.is_set():
                self._drain(port)
                try:
                    chunk = port.read(4096)
                except (serial.SerialException, OSError) as exc:
                    self.error.emit(f"读取串口失败：{exc}")
                    break
                if chunk:
                    self.data_received.emit(chunk)
        finally:
            try:
                port.close()
            except Exception:
                pass

    def _drain(self, port: serial.Serial) -> None:
        while not self._stop.is_set():
            try:
                payload = self._tx.get_nowait()
            except queue.Empty:
                return
            try:
                port.write(payload)
            except serial.SerialTimeoutException:
                self.error.emit("发送超时")
                return
            except (serial.SerialException, OSError) as exc:
                self.error.emit(f"发送失败：{exc}")
                return
            self.bytes_sent.emit(len(payload))


# ---------------------------------------------------------------------------
# 快捷指令编辑
# ---------------------------------------------------------------------------
class QuickCommandDialog(QDialog):
    def __init__(self, commands: list[dict], parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("快捷指令管理")
        self.resize(560, 420)
        self.commands = [dict(c) for c in commands]

        layout = QVBoxLayout(self)
        layout.setContentsMargins(14, 14, 14, 14)
        layout.setSpacing(8)

        self.list_widget = QListWidget()
        layout.addWidget(self.list_widget, 1)

        row = QHBoxLayout()
        self.name_edit = QLineEdit()
        self.name_edit.setPlaceholderText("名称，如：读保持寄存器")
        self.hex_edit = QLineEdit()
        self.hex_edit.setPlaceholderText("HEX 报文，如 01 03 00 00 00 0A")
        add_btn = QPushButton("添加")
        add_btn.clicked.connect(self._add)
        row.addWidget(self.name_edit, 2)
        row.addWidget(self.hex_edit, 3)
        row.addWidget(add_btn)
        layout.addLayout(row)

        bottom = QHBoxLayout()
        remove_btn = QPushButton("删除选中")
        remove_btn.clicked.connect(self._remove)
        ok_btn = QPushButton("确定")
        ok_btn.setObjectName("primary")
        ok_btn.clicked.connect(self.accept)
        bottom.addWidget(remove_btn)
        bottom.addStretch(1)
        bottom.addWidget(ok_btn)
        layout.addLayout(bottom)

        self._refresh()

    def _refresh(self) -> None:
        self.list_widget.clear()
        for cmd in self.commands:
            self.list_widget.addItem(f"{cmd['name']}    {cmd['hex']}")

    def _add(self) -> None:
        name = self.name_edit.text().strip()
        text = self.hex_edit.text().strip()
        if not name or not text:
            QMessageBox.warning(self, "提示", "名称和报文都不能为空")
            return
        try:
            mc.parse_hex_text(text)
        except ValueError as exc:
            QMessageBox.warning(self, "HEX 格式错误", str(exc))
            return
        self.commands.append({"name": name, "hex": text})
        self.name_edit.clear()
        self.hex_edit.clear()
        self._refresh()

    def _remove(self) -> None:
        row = self.list_widget.currentRow()
        if 0 <= row < len(self.commands):
            del self.commands[row]
            self._refresh()


# ---------------------------------------------------------------------------
# 主窗口
# ---------------------------------------------------------------------------
class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()

        self.config = Config()
        self._theme = "dark"
        self._protocol = "rtu"
        self._worker: SerialWorker | None = None
        self._state = "closed"
        self._status_is_error = False
        self._tid = 0

        self._rx_total = 0
        self._tx_total = 0
        self._frames = 0
        self._crc_errors = 0
        self._entries: list[tuple] = []
        self._rx_buffer = bytearray()
        self._rx_mode = "hex"
        self._send_mode = "hex"
        self._decoder = None

        self._quick_commands: list[dict] = []
        self._protocol_combos: list[QComboBox] = []

        # 从机地址扫描
        self._scan_active = False
        self._scan_cur = 1
        self._scan_begin = 1
        self._scan_last = 247
        self._scan_pending = None
        self._scan_buf = bytearray()
        self._scan_t0 = 0.0

        # 寄存器轮询
        self._poll_active = False
        self._poll_slaves: list[int] = [1]
        self._poll_idx = 0
        self._poll_pending = None
        self._poll_buf = bytearray()
        self._poll_base = 0
        self._poll_count = 10
        self._poll_func = 3
        self._poll_points: dict[int, dict] = {}

        # 定时器
        self._frame_timer = QTimer(self)
        self._frame_timer.setSingleShot(True)
        self._frame_timer.setInterval(60)
        self._frame_timer.timeout.connect(self._flush_frames)

        self._timed_timer = QTimer(self)
        self._timed_timer.timeout.connect(self._on_timed_tick)

        self._scan_wait = QTimer(self)
        self._scan_wait.setSingleShot(True)
        self._scan_wait.timeout.connect(self._scan_next)

        self._scan_timeout = QTimer(self)
        self._scan_timeout.setSingleShot(True)
        self._scan_timeout.timeout.connect(self._on_scan_timeout)

        self._poll_timer = QTimer(self)
        self._poll_timer.timeout.connect(self._poll_tick)

        self._poll_timeout = QTimer(self)
        self._poll_timeout.setSingleShot(True)
        self._poll_timeout.timeout.connect(self._on_poll_timeout)

        self._port_timer = QTimer(self)
        self._port_timer.setInterval(2000)
        self._port_timer.timeout.connect(self._check_ports)
        self._known_ports: list[str] | None = None

        self.setWindowTitle(APP_TITLE)
        self.resize(1300, 880)
        self.setMinimumSize(1080, 720)

        self._build_ui()
        self._bind_shortcuts()
        self._refresh_ports()
        self._load_config()
        self._apply_state()
        self._apply_theme(self._theme, persist=False)
        self._port_timer.start()
        self._set_status("就绪：选择串口后点击「打开串口」")
        self._update_stats()

    # ==================================================================
    # 界面
    # ==================================================================
    def _build_ui(self) -> None:
        root = QWidget()
        root.setObjectName("root")
        self.setCentralWidget(root)

        outer = QVBoxLayout(root)
        outer.setContentsMargins(16, 14, 16, 8)
        outer.setSpacing(10)

        header = QHBoxLayout()
        header.setSpacing(12)

        mark = QLabel("M")
        mark.setObjectName("appMark")
        mark.setFixedSize(38, 38)
        mark.setAlignment(Qt.AlignmentFlag.AlignCenter)

        title_box = QVBoxLayout()
        title_box.setContentsMargins(0, 0, 0, 0)
        title_box.setSpacing(1)
        title = QLabel("串口调试助手 · Modbus")
        title.setObjectName("appTitle")
        subtitle = QLabel("串口 / Modbus RTU · TCP 调试与报文分析")
        subtitle.setObjectName("appSubtitle")
        title_box.addWidget(title)
        title_box.addWidget(subtitle)

        self.protocol_badge = QLabel("Modbus RTU")
        self.protocol_badge.setObjectName("statePill")
        self.theme_btn = QPushButton("浅色模式")
        self.theme_btn.setObjectName("ghost")
        self.theme_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.theme_btn.clicked.connect(self._toggle_theme)

        header.addWidget(mark, 0, Qt.AlignmentFlag.AlignVCenter)
        header.addLayout(title_box)
        header.addSpacing(6)
        header.addWidget(self.protocol_badge, 0, Qt.AlignmentFlag.AlignVCenter)
        header.addStretch(1)
        header.addWidget(self.theme_btn, 0, Qt.AlignmentFlag.AlignVCenter)
        outer.addLayout(header)

        self.tabs = QTabWidget()
        self.tabs.addTab(self._build_monitor_tab(), "通讯监视")
        self.tabs.addTab(self._build_parser_tab(), "报文解析")
        self.tabs.addTab(self._build_scan_tab(), "地址扫描")
        self.tabs.addTab(self._build_poll_tab(), "数据监测")
        outer.addWidget(self.tabs, 1)

        mono = QFont()
        mono.setFamilies(["Cascadia Mono", "Consolas", "Courier New"])
        mono.setPointSize(10)
        for widget in (self.rx_view, self.send_edit, self.parser_input, self.parser_output):
            widget.setFont(mono)

        status = QStatusBar(self)
        status.setSizeGripEnabled(False)
        status.setContentsMargins(0, 0, 6, 0)
        self.setStatusBar(status)
        self.status_label = QLabel("就绪")
        self.status_label.setObjectName("statusText")
        self.stats_label = QLabel("接收帧 0 · CRC错误 0 · 丢包率 0.00%")
        self.stats_label.setObjectName("counterText")
        self.counter_label = QLabel("RX 0 B · TX 0 B")
        self.counter_label.setObjectName("counterText")
        status.addWidget(self.status_label, 1)
        status.addPermanentWidget(self.stats_label)
        status.addPermanentWidget(self.counter_label)

    @staticmethod
    def _card(title: str = "", margins: tuple[int, int, int, int] = (16, 12, 16, 14)):
        card = QFrame()
        card.setObjectName("card")
        box = QVBoxLayout(card)
        box.setContentsMargins(*margins)
        box.setSpacing(10)
        if title:
            label = QLabel(title)
            label.setObjectName("sectionTitle")
            box.addWidget(label)
        return card, box

    def _build_monitor_tab(self) -> QWidget:
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(2, 12, 2, 2)
        layout.setSpacing(12)
        layout.addWidget(self._build_serial_card())

        splitter = QSplitter(Qt.Orientation.Vertical)
        splitter.setChildrenCollapsible(False)
        splitter.setHandleWidth(8)
        splitter.addWidget(self._build_receive_card())
        splitter.addWidget(self._build_send_card())
        splitter.setStretchFactor(0, 3)
        splitter.setStretchFactor(1, 2)
        splitter.setSizes([430, 300])
        layout.addWidget(splitter, 1)
        return page

    def _build_serial_card(self) -> QFrame:
        card, box = self._card("串口设置")

        head = QHBoxLayout()
        self.state_pill = QLabel("未连接")
        self.state_pill.setObjectName("statePill")
        head.addStretch(1)
        head.addWidget(self.state_pill)
        box.addLayout(head)

        grid = QGridLayout()
        grid.setHorizontalSpacing(10)
        grid.setVerticalSpacing(4)
        labels = ("端口", "波特率", "数据位", "校验位", "停止位", "协议 / 站号")
        for col, text in enumerate(labels):
            label = QLabel(text)
            label.setObjectName("fieldLabel")
            grid.addWidget(label, 0, col)

        self.port_combo = QComboBox()
        self.port_combo.setMinimumWidth(190)
        self.refresh_btn = QPushButton("刷新")
        self.refresh_btn.clicked.connect(self._refresh_ports)
        port_row = QHBoxLayout()
        port_row.setContentsMargins(0, 0, 0, 0)
        port_row.setSpacing(6)
        port_row.addWidget(self.port_combo, 1)
        port_row.addWidget(self.refresh_btn)
        grid.addLayout(port_row, 1, 0)

        self.baud_combo = QComboBox()
        self.baud_combo.addItems([str(v) for v in BAUDRATE_LIST])
        self.baud_combo.setCurrentText("9600")
        grid.addWidget(self.baud_combo, 1, 1)

        self.databits_combo = QComboBox()
        for text, value in DATABITS_ITEMS:
            self.databits_combo.addItem(text, value)
        self.databits_combo.setCurrentIndex(1)
        grid.addWidget(self.databits_combo, 1, 2)

        self.parity_combo = QComboBox()
        for text, value in PARITY_ITEMS:
            self.parity_combo.addItem(text, value)
        grid.addWidget(self.parity_combo, 1, 3)

        self.stopbits_combo = QComboBox()
        for text, value in STOPBITS_ITEMS:
            self.stopbits_combo.addItem(text, value)
        grid.addWidget(self.stopbits_combo, 1, 4)

        self.protocol_combo = self._make_protocol_combo()
        self.station_spin = self._make_spin(0, 255, 1, 76)
        self.station_spin.setToolTip("Modbus TCP 的单元标识；RTU 由报文内从机地址决定")
        proto_row = QHBoxLayout()
        proto_row.setContentsMargins(0, 0, 0, 0)
        proto_row.setSpacing(6)
        proto_row.addWidget(self.protocol_combo, 1)
        proto_row.addWidget(self.station_spin)
        grid.addLayout(proto_row, 1, 5)

        self.open_btn = QPushButton("打开串口")
        self.open_btn.setObjectName("primary")
        self.open_btn.setMinimumHeight(44)
        self.open_btn.setMinimumWidth(124)
        self.open_btn.clicked.connect(self._toggle_port)
        grid.addWidget(self.open_btn, 0, 6, 2, 1)

        grid.setColumnStretch(0, 4)
        grid.setColumnStretch(1, 2)
        grid.setColumnStretch(5, 3)
        box.addLayout(grid)
        return card

    def _build_receive_card(self) -> QFrame:
        card, box = self._card("接收区")
        head = QHBoxLayout()
        head.setSpacing(8)

        self.rx_ascii_btn = self._seg_button("ASCII")
        self.rx_hex_btn = self._seg_button("HEX", checked=True)
        group = QButtonGroup(self)
        group.addButton(self.rx_ascii_btn, 0)
        group.addButton(self.rx_hex_btn, 1)
        group.idClicked.connect(self._on_rx_mode)
        seg = QHBoxLayout()
        seg.setContentsMargins(0, 0, 0, 0)
        seg.setSpacing(4)
        seg.addWidget(self.rx_ascii_btn)
        seg.addWidget(self.rx_hex_btn)
        head.addLayout(seg)
        head.addSpacing(6)

        self.encoding_combo = QComboBox()
        for text, value in ENCODING_ITEMS:
            self.encoding_combo.addItem(text, value)
        self.encoding_combo.currentIndexChanged.connect(self._on_encoding_changed)
        head.addWidget(self.encoding_combo)

        head.addStretch(1)
        self.auto_scroll = QCheckBox("自动滚动")
        self.auto_scroll.setChecked(True)
        self.clear_rx_btn = QPushButton("清空接收")
        self.clear_rx_btn.clicked.connect(self._clear_rx)
        head.addWidget(self.auto_scroll)
        head.addWidget(self.clear_rx_btn)
        box.addLayout(head)

        self.rx_view = QPlainTextEdit()
        self.rx_view.setReadOnly(True)
        self.rx_view.setUndoRedoEnabled(False)
        self.rx_view.setPlaceholderText("收发报文将带 [HH:MM:SS.xxx] 时间戳显示在这里…")
        box.addWidget(self.rx_view, 1)
        return card

    def _build_send_card(self) -> QFrame:
        card, box = self._card("发送区")
        head = QHBoxLayout()
        head.setSpacing(8)

        self.tx_ascii_btn = self._seg_button("ASCII")
        self.tx_hex_btn = self._seg_button("HEX", checked=True)
        group = QButtonGroup(self)
        group.addButton(self.tx_ascii_btn, 0)
        group.addButton(self.tx_hex_btn, 1)
        group.idClicked.connect(self._on_send_mode)
        seg = QHBoxLayout()
        seg.setContentsMargins(0, 0, 0, 0)
        seg.setSpacing(4)
        seg.addWidget(self.tx_ascii_btn)
        seg.addWidget(self.tx_hex_btn)
        head.addLayout(seg)
        head.addSpacing(6)

        self.crlf_check = QCheckBox("追加 \\r\\n")
        head.addWidget(self.crlf_check)
        head.addStretch(1)
        hint = QLabel("Ctrl+Enter 发送")
        hint.setObjectName("hint")
        head.addWidget(hint)
        box.addLayout(head)

        self.send_edit = QPlainTextEdit()
        self.send_edit.setPlaceholderText("RTU 输入完整帧（00 03 00 00 00 0A）或由自动 CRC 补全；TCP 只输入 PDU")
        self.send_edit.setMinimumHeight(70)
        self.send_edit.setTabChangesFocus(True)
        box.addWidget(self.send_edit, 1)

        quick_row = QHBoxLayout()
        quick_row.setSpacing(6)
        quick_label = QLabel("快捷指令")
        quick_label.setObjectName("fieldLabel")
        self.quick_area = QScrollArea()
        self.quick_area.setWidgetResizable(True)
        self.quick_area.setFixedHeight(40)
        self.quick_area.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.quick_area.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.quick_area.setStyleSheet("QScrollArea { background: transparent; border: none; }")
        self.quick_widget = QWidget()
        self.quick_layout = QHBoxLayout(self.quick_widget)
        self.quick_layout.setContentsMargins(0, 0, 0, 0)
        self.quick_layout.setSpacing(6)
        self.quick_area.setWidget(self.quick_widget)
        quick_edit = QPushButton("管理")
        quick_edit.setObjectName("ghost")
        quick_edit.clicked.connect(self._edit_quick)
        quick_row.addWidget(quick_label)
        quick_row.addWidget(self.quick_area, 1)
        quick_row.addWidget(quick_edit)
        box.addLayout(quick_row)

        footer = QHBoxLayout()
        footer.setSpacing(8)
        self.timed_check = QCheckBox("定时发送")
        self.timed_check.toggled.connect(self._on_timed_toggle)
        self.timed_interval = self._make_spin(100, 60000, 1000, 96, " ms")
        self.timed_interval.valueChanged.connect(self._on_timed_interval)
        self.crc_check = QCheckBox("自动追加CRC")
        self.crc_check.setChecked(True)
        self.crc_check.setToolTip("仅 RTU 生效；勾选后不要再手动输入 CRC")
        footer.addWidget(self.timed_check)
        footer.addWidget(self.timed_interval)
        footer.addWidget(self.crc_check)
        footer.addStretch(1)
        self.send_btn = QPushButton("发送")
        self.send_btn.setObjectName("primary")
        self.send_btn.setMinimumWidth(110)
        self.send_btn.clicked.connect(self._on_send)
        footer.addWidget(self.send_btn)
        box.addLayout(footer)
        return card

    def _build_parser_tab(self) -> QWidget:
        page = QWidget()
        layout = QHBoxLayout(page)
        layout.setContentsMargins(2, 12, 2, 2)
        layout.setSpacing(12)

        left, left_box = self._card("解析参数")
        left.setFixedWidth(260)
        self.parser_proto = self._make_protocol_combo()
        self.msg_type_combo = QComboBox()
        self.msg_type_combo.addItems(["从机应答", "主机指令"])
        self.data_type_combo = QComboBox()
        self.data_type_combo.addItems(DATA_TYPES)
        for text, widget in (
            ("协议类型", self.parser_proto),
            ("报文类型", self.msg_type_combo),
            ("数据类型", self.data_type_combo),
        ):
            label = QLabel(text)
            label.setObjectName("fieldLabel")
            left_box.addWidget(label)
            left_box.addWidget(widget)
        left_box.addStretch(1)
        self.parse_btn = QPushButton("解析报文")
        self.parse_btn.setObjectName("primary")
        self.parse_btn.clicked.connect(self._on_parse)
        self.parse_clear_btn = QPushButton("清空")
        self.parse_clear_btn.clicked.connect(self._clear_parser)
        left_box.addWidget(self.parse_btn)
        left_box.addWidget(self.parse_clear_btn)
        layout.addWidget(left)

        right = QVBoxLayout()
        right.setSpacing(10)

        input_card, input_box = self._card("报文输入（HEX）")
        self.parser_input = QPlainTextEdit()
        self.parser_input.setPlaceholderText("粘贴报文，空格可有可无：01 03 00 00 00 0A C5 CD")
        self.parser_input.setMinimumHeight(90)
        input_box.addWidget(self.parser_input)
        right.addWidget(input_card, 2)

        output_card, output_box = self._card("解析结果")
        self.parser_output = QPlainTextEdit()
        self.parser_output.setReadOnly(True)
        self.parser_output.setPlaceholderText("解析结果将显示在这里…")
        output_box.addWidget(self.parser_output)
        right.addWidget(output_card, 3)

        tools = QHBoxLayout()
        tools.setSpacing(10)

        crc_card, crc_box = self._card("CRC16 工具")
        self.crc_input = QLineEdit()
        self.crc_input.setPlaceholderText("输入十六进制数据")
        self.crc_result = QLabel("—")
        self.crc_result.setObjectName("hint")
        crc_btns = QHBoxLayout()
        calc_btn = QPushButton("计算 CRC")
        calc_btn.clicked.connect(self._calc_crc)
        copy_btn = QPushButton("复制带CRC报文")
        copy_btn.clicked.connect(self._copy_crc_frame)
        crc_btns.addWidget(calc_btn)
        crc_btns.addWidget(copy_btn)
        crc_btns.addStretch(1)
        crc_box.addWidget(self.crc_input)
        crc_box.addWidget(self.crc_result)
        crc_box.addLayout(crc_btns)
        tools.addWidget(crc_card, 1)

        conv_card, conv_box = self._card("ASCII / HEX 互转")
        self.conv_input = QLineEdit()
        self.conv_input.setPlaceholderText("文本或十六进制数据")
        self.conv_output = QLineEdit()
        self.conv_output.setReadOnly(True)
        self.conv_output.setPlaceholderText("转换结果")
        conv_btns = QHBoxLayout()
        to_hex = QPushButton("文本→HEX")
        to_hex.clicked.connect(lambda: self._convert("to_hex"))
        to_text = QPushButton("HEX→文本")
        to_text.clicked.connect(lambda: self._convert("to_text"))
        copy_conv = QPushButton("复制")
        copy_conv.clicked.connect(lambda: self._copy_text(self.conv_output.text()))
        conv_btns.addWidget(to_hex)
        conv_btns.addWidget(to_text)
        conv_btns.addWidget(copy_conv)
        conv_btns.addStretch(1)
        conv_box.addWidget(self.conv_input)
        conv_box.addWidget(self.conv_output)
        conv_box.addLayout(conv_btns)
        tools.addWidget(conv_card, 1)

        right.addLayout(tools)
        layout.addLayout(right, 1)
        return page

    def _build_scan_tab(self) -> QWidget:
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(2, 12, 2, 2)
        layout.setSpacing(12)

        card, box = self._card("从机地址扫描（1~247）")
        self.scan_proto = self._make_protocol_combo()
        self.scan_begin_spin = self._make_spin(1, 247, 1, 90)
        self.scan_end_spin = self._make_spin(1, 247, 20, 90)
        self.scan_reg_spin = self._make_spin(0, 65535, 0, 100)
        self.scan_interval_spin = self._make_spin(50, 60000, 300, 96, " ms")
        self.scan_timeout_spin = self._make_spin(100, 30000, 500, 96, " ms")

        fields = [
            ("协议", self.scan_proto),
            ("起始从机", self.scan_begin_spin),
            ("结束从机", self.scan_end_spin),
            ("测试寄存器", self.scan_reg_spin),
            ("扫描间隔", self.scan_interval_spin),
            ("响应超时", self.scan_timeout_spin),
        ]
        grid = QGridLayout()
        grid.setHorizontalSpacing(10)
        grid.setVerticalSpacing(4)
        for index, (text, widget) in enumerate(fields):
            label = QLabel(text)
            label.setObjectName("fieldLabel")
            grid.addWidget(label, 0, index)
            grid.addWidget(widget, 1, index)
        box.addLayout(grid)

        buttons = QHBoxLayout()
        self.scan_start_btn = QPushButton("开始扫描")
        self.scan_start_btn.setObjectName("primary")
        self.scan_start_btn.clicked.connect(self._scan_start)
        self.scan_stop_btn = QPushButton("停止扫描")
        self.scan_stop_btn.clicked.connect(self._scan_stop)
        self.scan_clear_btn = QPushButton("清空结果")
        self.scan_clear_btn.clicked.connect(self._clear_scan)
        self.scan_state = QLabel("空闲")
        self.scan_state.setObjectName("hint")
        buttons.addWidget(self.scan_start_btn)
        buttons.addWidget(self.scan_stop_btn)
        buttons.addWidget(self.scan_clear_btn)
        buttons.addStretch(1)
        buttons.addWidget(self.scan_state)
        box.addLayout(buttons)

        self.scan_table = self._make_table(("从机地址", "状态", "响应时间(ms)", "备注"))
        box.addWidget(self.scan_table, 1)
        layout.addWidget(card, 1)
        return page

    def _build_poll_tab(self) -> QWidget:
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(2, 12, 2, 2)
        layout.setSpacing(12)

        card, box = self._card("寄存器数据监测")
        self.poll_proto = self._make_protocol_combo()
        self.poll_func_combo = QComboBox()
        for text, value in FUNCTION_ITEMS:
            self.poll_func_combo.addItem(text, value)
        self.poll_start_spin = self._make_spin(0, 65535, 0, 100)
        self.poll_count_spin = self._make_spin(1, 125, 10, 90)
        self.poll_interval_spin = self._make_spin(100, 60000, 1000, 96, " ms")
        self.poll_format_combo = QComboBox()
        self.poll_format_combo.addItems(VALUE_FORMATS)

        fields = [
            ("协议", self.poll_proto),
            ("功能码", self.poll_func_combo),
            ("起始地址", self.poll_start_spin),
            ("寄存器数量", self.poll_count_spin),
            ("轮询间隔", self.poll_interval_spin),
            ("数值格式", self.poll_format_combo),
        ]
        grid = QGridLayout()
        grid.setHorizontalSpacing(10)
        grid.setVerticalSpacing(4)
        for index, (text, widget) in enumerate(fields):
            label = QLabel(text)
            label.setObjectName("fieldLabel")
            grid.addWidget(label, 0, index)
            grid.addWidget(widget, 1, index)
        box.addLayout(grid)

        slaves_row = QHBoxLayout()
        slaves_label = QLabel("从机列表")
        slaves_label.setObjectName("fieldLabel")
        self.slave_spin = self._make_spin(1, 247, 1, 80)
        add_slave = QPushButton("添加")
        add_slave.clicked.connect(self._add_slave)
        del_slave = QPushButton("删除")
        del_slave.clicked.connect(self._del_slave)
        self.slave_list = QListWidget()
        self.slave_list.setFixedHeight(62)
        self.slave_list.setMaximumWidth(160)
        slaves_row.addWidget(slaves_label)
        slaves_row.addWidget(self.slave_spin)
        slaves_row.addWidget(add_slave)
        slaves_row.addWidget(del_slave)
        slaves_row.addWidget(self.slave_list)
        slaves_row.addStretch(1)

        import_btn = QPushButton("导入点表")
        import_btn.clicked.connect(self._import_points)
        export_btn = QPushButton("导出Excel")
        export_btn.clicked.connect(self._export_excel)
        slaves_row.addWidget(import_btn)
        slaves_row.addWidget(export_btn)
        box.addLayout(slaves_row)

        buttons = QHBoxLayout()
        self.poll_start_btn = QPushButton("开始监测")
        self.poll_start_btn.setObjectName("primary")
        self.poll_start_btn.clicked.connect(self._poll_start)
        self.poll_stop_btn = QPushButton("停止监测")
        self.poll_stop_btn.clicked.connect(self._poll_stop)
        self.poll_clear_btn = QPushButton("清空")
        self.poll_clear_btn.clicked.connect(self._clear_poll)
        self.poll_state = QLabel("空闲")
        self.poll_state.setObjectName("hint")
        buttons.addWidget(self.poll_start_btn)
        buttons.addWidget(self.poll_stop_btn)
        buttons.addWidget(self.poll_clear_btn)
        buttons.addStretch(1)
        buttons.addWidget(self.poll_state)
        box.addLayout(buttons)

        self.poll_table = self._make_table(
            ("序号", "设备地址", "寄存器地址(HEX)", "点位名称", "当前数值", "单位")
        )
        box.addWidget(self.poll_table, 1)
        layout.addWidget(card, 1)
        return page

    # ------------------------------------------------------------------
    # 控件工厂
    # ------------------------------------------------------------------
    def _make_protocol_combo(self) -> QComboBox:
        combo = QComboBox()
        for text, value in PROTOCOL_ITEMS:
            combo.addItem(text, value)
        combo.setMinimumWidth(118)
        combo.currentIndexChanged.connect(self._on_protocol_changed)
        self._protocol_combos.append(combo)
        return combo

    @staticmethod
    def _make_spin(minimum: int, maximum: int, value: int, width: int, suffix: str = "") -> QSpinBox:
        spin = QSpinBox()
        spin.setRange(minimum, maximum)
        spin.setValue(value)
        spin.setMinimumWidth(width)
        if suffix:
            spin.setSuffix(suffix)
        return spin

    @staticmethod
    def _make_table(headers: tuple[str, ...]) -> QTableWidget:
        table = QTableWidget(0, len(headers))
        table.setHorizontalHeaderLabels(list(headers))
        table.verticalHeader().setVisible(False)
        table.verticalHeader().setDefaultSectionSize(30)
        table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        table.setAlternatingRowColors(True)
        table.setShowGrid(False)
        table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        mono = QFont()
        mono.setFamilies(["Cascadia Mono", "Consolas"])
        mono.setPointSize(9)
        table.setFont(mono)
        return table

    @staticmethod
    def _seg_button(text: str, checked: bool = False) -> QRadioButton:
        button = QRadioButton(text)
        button.setProperty("seg", True)
        button.setChecked(checked)
        button.setCursor(Qt.CursorShape.PointingHandCursor)
        return button

    def _bind_shortcuts(self) -> None:
        for sequence in ("Ctrl+Return", "Ctrl+Enter"):
            QShortcut(QKeySequence(sequence), self).activated.connect(self._on_send)
        QShortcut(QKeySequence("Ctrl+L"), self).activated.connect(self._clear_rx)

    # ==================================================================
    # 配置持久化
    # ==================================================================
    def _load_config(self) -> None:
        data = self.config.data
        serial_cfg = data.get("serial", {})

        index = self.port_combo.findData(serial_cfg.get("port", ""))
        if index >= 0:
            self.port_combo.setCurrentIndex(index)
        index = self.baud_combo.findText(str(serial_cfg.get("baudrate", 9600)))
        if index >= 0:
            self.baud_combo.setCurrentIndex(index)
        index = self.databits_combo.findData(int(serial_cfg.get("bytesize", 8)))
        if index >= 0:
            self.databits_combo.setCurrentIndex(index)
        index = self.parity_combo.findData(str(serial_cfg.get("parity", "N")))
        if index >= 0:
            self.parity_combo.setCurrentIndex(index)
        index = self.stopbits_combo.findData(int(serial_cfg.get("stopbits", 1)))
        if index >= 0:
            self.stopbits_combo.setCurrentIndex(index)
        try:
            station = int(serial_cfg.get("station", 1) or 1)
        except Exception:
            station = 1
        self.station_spin.setValue(station)

        self._quick_commands = [dict(c) for c in data.get("quick", []) if isinstance(c, dict)]
        self._rebuild_quick_bar()

        poll_cfg = data.get("poll", {})
        slaves = []
        for value in poll_cfg.get("slaves", [1]):
            try:
                slaves.append(int(value))
            except Exception:
                continue
        self._poll_slaves = sorted(set(slaves)) or [1]
        self._poll_points = {}
        for point in poll_cfg.get("points", []):
            try:
                self._poll_points[int(point["address"])] = {
                    "name": str(point.get("name", "")),
                    "factor": float(point.get("factor", 1) or 1),
                    "unit": str(point.get("unit", "")),
                }
            except Exception:
                continue
        self._rebuild_slave_list()

        self._apply_protocol(data.get("protocol", "rtu"), persist=False)
        self._theme = "light" if data.get("theme") == "light" else "dark"
        self._theme_label()

    def _save_serial_config(self) -> None:
        cfg = self.config.data["serial"]
        cfg.update({
            "port": str(self.port_combo.currentData() or ""),
            "baudrate": int(self.baud_combo.currentText() or 9600),
            "bytesize": int(self.databits_combo.currentData() or 8),
            "parity": str(self.parity_combo.currentData() or "N"),
            "stopbits": int(self.stopbits_combo.currentData() or 1),
            "station": int(self.station_spin.value()),
        })
        self.config.save()

    def _save_poll_config(self) -> None:
        self.config.data["poll"] = {
            "slaves": list(self._poll_slaves),
            "points": [
                {"address": address, **point}
                for address, point in sorted(self._poll_points.items())
            ],
        }
        self.config.save()

    # ==================================================================
    # 主题
    # ==================================================================
    def _toggle_theme(self) -> None:
        self._apply_theme("light" if self._theme == "dark" else "dark")

    def _theme_label(self) -> None:
        self.theme_btn.setText("浅色模式" if self._theme == "dark" else "深色模式")

    def _apply_theme(self, theme: str, persist: bool = True) -> None:
        self._theme = "light" if theme == "light" else "dark"
        app = QApplication.instance()
        if app is not None:
            app.setPalette(build_palette(self._theme))
            app.setStyleSheet(build_qss(self._theme))
        self._theme_label()
        self._rebuild_display()
        if persist:
            self.config.data["theme"] = self._theme
            self.config.save()

    # ==================================================================
    # 协议
    # ==================================================================
    def _on_protocol_changed(self, _index: int) -> None:
        combo = self.sender()
        if isinstance(combo, QComboBox):
            self._apply_protocol(combo.currentData())

    def _apply_protocol(self, protocol: str, persist: bool = True) -> None:
        protocol = "tcp" if protocol == "tcp" else "rtu"
        changed = protocol != self._protocol
        self._protocol = protocol
        for combo in self._protocol_combos:
            combo.blockSignals(True)
            combo.setCurrentIndex(1 if protocol == "tcp" else 0)
            combo.blockSignals(False)
        self.crc_check.setEnabled(protocol == "rtu")
        self.protocol_badge.setText("Modbus TCP" if protocol == "tcp" else "Modbus RTU")
        if changed:
            self._scan_abort()
            self._poll_abort()
            self._apply_state()
        if persist:
            self.config.data["protocol"] = protocol
            self.config.save()

    # ==================================================================
    # 串口状态
    # ==================================================================
    def _refresh_ports(self) -> None:
        previous = self.port_combo.currentData()
        try:
            ports = sorted(list_ports.comports(), key=lambda p: p.device)
        except Exception:
            ports = []

        self.port_combo.blockSignals(True)
        self.port_combo.clear()
        if ports:
            for info in ports:
                text = f"{info.device} · {info.description}" if info.description else info.device
                self.port_combo.addItem(text, info.device)
                self.port_combo.setItemData(self.port_combo.count() - 1, info.hwid, Qt.ItemDataRole.ToolTipRole)
            if previous:
                index = self.port_combo.findData(previous)
                if index >= 0:
                    self.port_combo.setCurrentIndex(index)
        else:
            self.port_combo.addItem("未检测到可用串口", None)
        self.port_combo.blockSignals(False)
        self._known_ports = [info.device for info in ports]
        self._apply_state()

    def _check_ports(self) -> None:
        try:
            current = sorted(p.device for p in list_ports.comports())
        except Exception:
            return
        if self._known_ports is None:
            self._known_ports = current
            return
        if set(current) == set(self._known_ports):
            return

        added = sorted(set(current) - set(self._known_ports))
        removed = sorted(set(self._known_ports) - set(current))
        self._known_ports = current
        self._refresh_ports()

        parts = []
        if added:
            parts.append("新增：" + "、".join(added))
        if removed:
            parts.append("移除：" + "、".join(removed))
        QMessageBox.information(self, "串口变化", "\n".join(parts) or "串口列表已更新")

    def _apply_state(self) -> None:
        locked = self._state in ("opening", "open", "closing")
        self.port_combo.setEnabled(not locked and self.port_combo.currentData() is not None)
        self.refresh_btn.setEnabled(not locked)
        for widget in (self.baud_combo, self.databits_combo, self.parity_combo, self.stopbits_combo):
            widget.setEnabled(not locked)

        if self._state == "open":
            self.open_btn.setText("关闭串口")
        elif self._state == "opening":
            self.open_btn.setText("正在打开…")
        elif self._state == "closing":
            self.open_btn.setText("正在关闭…")
        else:
            self.open_btn.setText("打开串口")
        self.open_btn.setEnabled(self._state in ("closed", "open"))
        danger = self._state == "open"
        if bool(self.open_btn.property("danger")) != danger:
            self.open_btn.setProperty("danger", danger)
            self.open_btn.style().unpolish(self.open_btn)
            self.open_btn.style().polish(self.open_btn)

        is_open = self._state == "open"
        self.send_btn.setEnabled(is_open)
        self.crlf_check.setEnabled(self._send_mode == "ascii")
        self.timed_check.setEnabled(is_open)
        self.timed_interval.setEnabled(is_open)
        self.crc_check.setEnabled(is_open and self._protocol == "rtu")
        if not is_open:
            self._timed_timer.stop()
            self._set_checked(self.timed_check, False)
        self.scan_start_btn.setEnabled(is_open and not self._scan_active)
        self.scan_stop_btn.setEnabled(self._scan_active)
        self.poll_start_btn.setEnabled(is_open and not self._poll_active)
        self.poll_stop_btn.setEnabled(self._poll_active)

        text, state = {
            "closed": ("未连接", "off"),
            "opening": ("连接中…", "busy"),
            "open": ("已连接", "on"),
            "closing": ("断开中…", "busy"),
        }[self._state]
        self.state_pill.setText(text)
        if self.state_pill.property("state") != state:
            self.state_pill.setProperty("state", state)
            self.state_pill.style().unpolish(self.state_pill)
            self.state_pill.style().polish(self.state_pill)

    @staticmethod
    def _set_checked(widget: QCheckBox, checked: bool) -> None:
        widget.blockSignals(True)
        widget.setChecked(checked)
        widget.blockSignals(False)

    def _toggle_port(self) -> None:
        if self._state == "open":
            self._close_port()
        elif self._state == "closed":
            self._open_port()

    def _open_port(self) -> None:
        port = self.port_combo.currentData()
        if not port:
            self._set_status("请先选择串口", error=True)
            return
        try:
            baudrate = int(self.baud_combo.currentText())
        except ValueError:
            self._set_status("波特率无效", error=True)
            return

        cfg = {
            "port": str(port),
            "baudrate": baudrate,
            "bytesize": int(self.databits_combo.currentData()),
            "parity": str(self.parity_combo.currentData()),
            "stopbits": int(self.stopbits_combo.currentData()),
        }
        worker = SerialWorker(cfg)
        worker.opened.connect(self._on_opened)
        worker.open_failed.connect(lambda msg: self._set_status(f"打开串口失败：{msg}", error=True))
        worker.data_received.connect(self._on_data)
        worker.bytes_sent.connect(self._on_bytes_sent)
        worker.error.connect(lambda msg: self._set_status(f"串口异常：{msg}", error=True))
        worker.finished.connect(self._on_worker_finished)

        self._worker = worker
        self._state = "opening"
        self._apply_state()
        self._set_status(f"正在打开 {port}…")
        worker.start()

    def _close_port(self) -> None:
        if self._worker is None or self._state != "open":
            return
        self._save_serial_config()
        self._state = "closing"
        self._apply_state()
        self._set_status("正在关闭串口…")
        self._worker.stop()

    def _on_opened(self) -> None:
        self._state = "open"
        self._apply_state()
        self._save_serial_config()
        self._set_status(f"已连接 {self.port_combo.currentData()} · {self.baud_combo.currentText()}")

    def _on_worker_finished(self) -> None:
        self._timed_timer.stop()
        self._set_checked(self.timed_check, False)
        self._scan_abort()
        self._poll_abort()
        worker = self._worker
        self._worker = None
        if worker is not None:
            worker.deleteLater()
        self._state = "closed"
        self._apply_state()
        if not self._status_is_error:
            self._set_status("串口已关闭")

    def _on_bytes_sent(self, count: int) -> None:
        self._tx_total += count
        self._update_counters()

    # ==================================================================
    # 接收与显示
    # ==================================================================
    def _on_data(self, data: bytes) -> None:
        if not data:
            return
        self._rx_total += len(data)
        self._update_counters()
        self._rx_buffer.extend(data)
        if len(self._rx_buffer) >= 4096:
            self._frame_timer.stop()
            self._flush_frames()
        else:
            self._frame_timer.start()
        if self._scan_active:
            self._scan_feed(data)
        if self._poll_active:
            self._poll_feed(data)

    @staticmethod
    def _exception_code(protocol: str, frame: bytes):
        if protocol == "rtu":
            if len(frame) >= 3 and frame[1] & 0x80:
                return frame[2]
            return None
        pdu = frame[7:]
        if len(pdu) >= 2 and pdu[0] & 0x80:
            return pdu[1]
        return None

    def _flush_frames(self) -> None:
        data = bytes(self._rx_buffer)
        self._rx_buffer.clear()
        if not data:
            return
        if not mc.looks_modbus(self._protocol, data):
            self._add_entry("RX", data, True, "")
            return
        frames, consumed = mc.split_frames(self._protocol, data)
        if not frames:
            self._add_entry("RX", data, True, "")
            return
        if consumed < len(data):
            self._rx_buffer.extend(data[consumed:])
            self._frame_timer.start()
        for frame, ok, note in frames:
            text = note
            if ok:
                code = self._exception_code(self._protocol, frame)
                if code is not None:
                    text = f"异常应答 {code:02X}：{mc.exception_name(code)}"
            self._frames += 1
            if not ok:
                self._crc_errors += 1
            self._add_entry("RX", frame, ok, text)
        self._update_stats()

    def _add_entry(self, direction: str, payload: bytes, ok: bool, note: str) -> None:
        self._entries.append((now_stamp(), direction, payload, ok, note))
        if len(self._entries) > 3000:
            del self._entries[:500]
            self._rebuild_display()
        else:
            self._append_entry(self._entries[-1])

    def _entry_text(self, payload: bytes) -> str:
        if self._rx_mode == "hex":
            return payload.hex(" ").upper()
        text = payload.decode(self._encoding(), errors="replace")
        return text.replace("\r\n", " ").replace("\r", " ").replace("\n", " ")

    def _append_entry(self, entry: tuple) -> None:
        stamp, direction, payload, ok, note = entry
        c = colors(self._theme)
        body = self._entry_text(payload)
        cursor = self.rx_view.textCursor()
        cursor.movePosition(QTextCursor.MoveOperation.End)
        fmt_meta = QTextCharFormat()
        fmt_meta.setForeground(QColor(c["faint"]))
        cursor.insertText(f"[{stamp}] {direction} ", fmt_meta)
        color = c["tx"] if direction == "TX" else (c["danger"] if not ok else c["rx"])
        fmt_body = QTextCharFormat()
        fmt_body.setForeground(QColor(color))
        cursor.insertText(body, fmt_body)
        if note:
            fmt_note = QTextCharFormat()
            fmt_note.setForeground(QColor(c["danger"] if not ok else c["muted"]))
            cursor.insertText(f"   [{note}]", fmt_note)
        cursor.insertText("\n")
        if self.auto_scroll.isChecked():
            bar = self.rx_view.verticalScrollBar()
            bar.setValue(bar.maximum())

    def _rebuild_display(self) -> None:
        self.rx_view.clear()
        for entry in self._entries:
            self._append_entry(entry)

    def _encoding(self) -> str:
        return self.encoding_combo.currentData() or "utf-8"

    def _on_rx_mode(self, index: int) -> None:
        self._rx_mode = "hex" if index == 1 else "ascii"
        self._rebuild_display()

    def _on_encoding_changed(self, _index: int) -> None:
        self._rebuild_display()

    def _clear_rx(self) -> None:
        self._entries.clear()
        self._rx_buffer.clear()
        self.rx_view.clear()
        self._rx_total = 0
        self._frames = 0
        self._crc_errors = 0
        self._update_counters()
        self._update_stats()

    def _update_stats(self) -> None:
        loss = (self._crc_errors / self._frames * 100) if self._frames else 0.0
        self.stats_label.setText(
            f"接收帧 {self._frames} · CRC错误 {self._crc_errors} · 丢包率 {loss:.2f}%"
        )

    def _update_counters(self) -> None:
        self.counter_label.setText(f"RX {self._rx_total:,} B · TX {self._tx_total:,} B")

    # ==================================================================
    # 发送
    # ==================================================================
    def _on_send_mode(self, index: int) -> None:
        self._send_mode = "hex" if index == 1 else "ascii"
        self.crlf_check.setEnabled(self._send_mode == "ascii")

    def _next_tid(self) -> int:
        self._tid = (self._tid + 1) & 0xFFFF
        return self._tid

    def _build_payload(self, raw: str, mode: str) -> tuple[bytes, str]:
        if not raw.strip():
            raise ValueError("发送内容为空")
        if mode == "hex":
            payload = mc.parse_hex_text(raw)
        else:
            payload = raw.encode(self._encoding())
            if self.crlf_check.isChecked():
                payload += b"\r\n"
        if not payload:
            raise ValueError("发送内容为空")

        if self._protocol == "tcp":
            tid = self._next_tid()
            return mc.mbap_build(payload, self.station_spin.value(), tid), f"MBAP tid={tid}"
        if self.crc_check.isChecked():
            value = mc.crc16(payload)
            return mc.with_crc(payload), f"CRC 0x{value:04X}"
        return payload, ""

    def _transmit(self, payload: bytes, note: str) -> bool:
        if self._state != "open" or self._worker is None:
            self._set_status("串口未打开，无法发送", error=True)
            return False
        if not self._worker.send(payload):
            self._set_status("发送失败：串口已关闭", error=True)
            return False
        self._add_entry("TX", payload, True, note)
        return True

    def _on_send(self) -> None:
        try:
            payload, note = self._build_payload(self.send_edit.toPlainText(), self._send_mode)
        except ValueError as exc:
            self._set_status(f"发送失败：{exc}", error=True)
            return
        if self._transmit(payload, note):
            preview = payload[:24].hex(" ").upper()
            suffix = " …" if len(payload) > 24 else ""
            self._set_status(f"已发送 {len(payload)} 字节：{preview}{suffix}")

    def _send_quick(self, hex_text: str) -> None:
        try:
            payload, note = self._build_payload(hex_text, "hex")
        except ValueError as exc:
            self._set_status(f"快捷指令无效：{exc}", error=True)
            return
        if self._transmit(payload, note):
            self._set_status(f"快捷指令已发送：{payload[:24].hex(' ').upper()}")

    def _on_timed_toggle(self, checked: bool) -> None:
        if checked:
            if self._state != "open":
                self._set_checked(self.timed_check, False)
                self._set_status("请先打开串口", error=True)
                return
            self._on_timed_tick()
            if not self.timed_check.isChecked():
                return
            self._timed_timer.start(max(100, self.timed_interval.value()))
            self._set_status(f"定时发送已启动，间隔 {self.timed_interval.value()} ms")
        else:
            self._timed_timer.stop()
            if self._state == "open":
                self._set_status("定时发送已停止")

    def _on_timed_interval(self, value: int) -> None:
        if self._timed_timer.isActive():
            self._timed_timer.setInterval(max(100, value))

    def _on_timed_tick(self) -> None:
        if self._state != "open":
            self._timed_timer.stop()
            self._set_checked(self.timed_check, False)
            return
        try:
            payload, note = self._build_payload(self.send_edit.toPlainText(), self._send_mode)
        except ValueError as exc:
            self._timed_timer.stop()
            self._set_checked(self.timed_check, False)
            self._set_status(f"定时发送已停止：{exc}", error=True)
            return
        self._transmit(payload, note)

    # ------------------------------------------------------------------
    # 快捷指令
    # ------------------------------------------------------------------
    def _rebuild_quick_bar(self) -> None:
        while self.quick_layout.count():
            item = self.quick_layout.takeAt(0)
            widget = item.widget()
            if widget is not None:
                widget.deleteLater()
        for command in self._quick_commands:
            button = QPushButton(command["name"])
            button.setObjectName("ghost")
            button.setToolTip(command["hex"])
            button.clicked.connect(lambda _=False, h=command["hex"]: self._send_quick(h))
            self.quick_layout.addWidget(button)
        self.quick_layout.addStretch(1)

    def _edit_quick(self) -> None:
        dialog = QuickCommandDialog(self._quick_commands, self)
        if dialog.exec():
            self._quick_commands = dialog.commands
            self.config.data["quick"] = self._quick_commands
            self.config.save()
            self._rebuild_quick_bar()

    # ==================================================================
    # 报文解析页
    # ==================================================================
    def _on_parse(self) -> None:
        text = self.parser_input.toPlainText().strip()
        if not text:
            self._set_status("请先粘贴报文", error=True)
            return
        try:
            data = mc.parse_hex_text(text)
        except ValueError as exc:
            self.parser_output.setPlainText(f"HEX 解析失败：{exc}")
            return
        report = mc.describe_frame(
            self._protocol,
            data,
            self.msg_type_combo.currentText(),
            self.data_type_combo.currentText(),
        )
        self.parser_output.setPlainText(report)
        self._set_status("报文解析完成")

    def _clear_parser(self) -> None:
        self.parser_input.clear()
        self.parser_output.clear()

    def _calc_crc(self) -> None:
        try:
            data = mc.parse_hex_text(self.crc_input.text())
        except ValueError as exc:
            self.crc_result.setText(f"输入无效：{exc}")
            return
        value = mc.crc16(data)
        self.crc_frame = mc.with_crc(data)
        self.crc_result.setText(
            f"CRC16 = 0x{value:04X}（低字节 0x{value & 0xFF:02X}，高字节 0x{value >> 8:02X}）"
        )

    def _copy_crc_frame(self) -> None:
        frame = getattr(self, "crc_frame", b"")
        if not frame:
            self._set_status("请先计算 CRC", error=True)
            return
        self._copy_text(mc.bytes_to_hex(frame))

    def _convert(self, direction: str) -> None:
        try:
            if direction == "to_hex":
                result = mc.text_to_hex(self.conv_input.text(), self._encoding())
            else:
                result = mc.hex_to_text(self.conv_input.text(), self._encoding())
        except ValueError as exc:
            result = f"转换失败：{exc}"
        self.conv_output.setText(result)

    def _copy_text(self, text: str) -> None:
        if not text:
            return
        QApplication.clipboard().setText(text)
        self._set_status("已复制到剪贴板")

    # ==================================================================
    # 从机地址扫描
    # ==================================================================
    def _scan_start(self) -> None:
        if self._state != "open":
            self._set_status("串口未打开，无法扫描", error=True)
            return
        begin = self.scan_begin_spin.value()
        last = self.scan_end_spin.value()
        if begin > last:
            self._set_status("起始从机不能大于结束从机", error=True)
            return
        self._scan_active = True
        self._scan_begin = begin
        self._scan_last = last
        self._scan_cur = begin
        self._scan_pending = None
        self._scan_buf.clear()
        self.scan_table.setRowCount(0)
        self.scan_state.setText("扫描中…")
        self._apply_state()
        self._set_status(f"开始扫描从机 {begin}~{last}")
        self._scan_next()

    def _scan_next(self) -> None:
        if not self._scan_active:
            return
        if self._state != "open" or self._worker is None:
            self._scan_stop()
            return
        if self._scan_cur > self._scan_last:
            self._scan_finish()
            return
        tid = self._next_tid()
        frame = mc.build_read_request(
            self._protocol, self._scan_cur, 3, self.scan_reg_spin.value(), 1, tid
        )
        self._scan_pending = (self._scan_cur, tid)
        self._scan_buf.clear()
        self._scan_t0 = time.time()
        if not self._worker.send(frame):
            self._scan_stop()
            return
        self._scan_timeout.start(max(100, self.scan_timeout_spin.value()))

    def _scan_feed(self, data: bytes) -> None:
        if not self._scan_active or self._scan_pending is None:
            return
        self._scan_buf.extend(data)
        slave, tid = self._scan_pending
        frame, consumed = mc.extract_response(self._protocol, self._scan_buf, slave, 3, tid)
        if frame is None:
            return
        del self._scan_buf[:consumed]
        self._scan_timeout.stop()
        elapsed = (time.time() - self._scan_t0) * 1000
        code = self._exception_code(self._protocol, frame)
        if code is not None:
            self._append_scan_row(slave, "在线(异常)", elapsed, f"异常应答 {code:02X}")
        else:
            self._append_scan_row(slave, "在线", elapsed, "")
        self._scan_pending = None
        self._scan_cur = slave + 1
        self._scan_wait.start(max(50, self.scan_interval_spin.value()))

    def _on_scan_timeout(self) -> None:
        if not self._scan_active or self._scan_pending is None:
            return
        slave, _tid = self._scan_pending
        self._scan_pending = None
        self._scan_cur = slave + 1
        self._scan_wait.start(max(50, self.scan_interval_spin.value()))

    def _append_scan_row(self, slave: int, state: str, elapsed: float, note: str) -> None:
        row = self.scan_table.rowCount()
        self.scan_table.insertRow(row)
        for col, text in enumerate((str(slave), state, f"{elapsed:.0f}", note)):
            item = QTableWidgetItem(text)
            item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.scan_table.setItem(row, col, item)

    def _scan_finish(self) -> None:
        self._scan_active = False
        self._scan_timeout.stop()
        self._scan_wait.stop()
        self._scan_pending = None
        self.scan_state.setText("已完成")
        self._apply_state()
        self._set_status(f"扫描完成，在线从机 {self.scan_table.rowCount()} 个")

    def _scan_stop(self) -> None:
        self._scan_abort()
        self.scan_state.setText("空闲")
        self._apply_state()
        self._set_status("扫描已停止")

    def _scan_abort(self) -> None:
        self._scan_active = False
        self._scan_timeout.stop()
        self._scan_wait.stop()
        self._scan_pending = None
        self._scan_buf.clear()
        if hasattr(self, "scan_state"):
            self.scan_state.setText("空闲")

    def _clear_scan(self) -> None:
        if self._scan_active:
            self._scan_stop()
        self.scan_table.setRowCount(0)
        self._set_status("扫描结果已清空")

    # ==================================================================
    # 数据监测
    # ==================================================================
    def _rebuild_slave_list(self) -> None:
        self.slave_list.clear()
        for slave in self._poll_slaves:
            self.slave_list.addItem(str(slave))

    def _add_slave(self) -> None:
        slave = self.slave_spin.value()
        if slave in self._poll_slaves:
            self._set_status(f"从机 {slave} 已在列表中", error=True)
            return
        if self._poll_active:
            self._poll_stop()
        self._poll_slaves.append(slave)
        self._poll_slaves.sort()
        self._rebuild_slave_list()
        self._save_poll_config()

    def _del_slave(self) -> None:
        row = self.slave_list.currentRow()
        if 0 <= row < len(self._poll_slaves):
            if self._poll_active:
                self._poll_stop()
            del self._poll_slaves[row]
            self._rebuild_slave_list()
            self._save_poll_config()

    def _point_for(self, address: int) -> dict:
        return self._poll_points.get(address, {})

    def _poll_start(self) -> None:
        if self._state != "open":
            self._set_status("串口未打开，无法监测", error=True)
            return
        if not self._poll_slaves:
            self._set_status("请先添加从机", error=True)
            return
        self._poll_active = True
        self._poll_base = self.poll_start_spin.value()
        self._poll_count = self.poll_count_spin.value()
        self._poll_func = int(self.poll_func_combo.currentData())
        self._poll_idx = 0
        self._poll_pending = None
        self._poll_buf.clear()
        self._build_poll_table()
        self.poll_state.setText("监测中…")
        self._apply_state()
        self._set_status(f"开始监测 {len(self._poll_slaves)} 个从机")
        self._poll_tick()
        if self._poll_active:
            self._poll_timer.start(max(100, self.poll_interval_spin.value()))

    def _build_poll_table(self) -> None:
        self.poll_table.setRowCount(0)
        for slave in self._poll_slaves:
            for offset in range(self._poll_count):
                address = self._poll_base + offset
                point = self._point_for(address)
                row = self.poll_table.rowCount()
                self.poll_table.insertRow(row)
                texts = (
                    str(row + 1),
                    str(slave),
                    f"0x{address:04X}",
                    point.get("name", ""),
                    "—",
                    point.get("unit", ""),
                )
                for col, text in enumerate(texts):
                    item = QTableWidgetItem(text)
                    item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                    self.poll_table.setItem(row, col, item)

    def _value_text(self, address: int, raw: int) -> str:
        point = self._point_for(address)
        try:
            factor = float(point.get("factor", 1) or 1)
        except Exception:
            factor = 1.0
        fmt = self.poll_format_combo.currentText()
        if factor != 1:
            scaled = raw * factor
            return f"{scaled:.3f}".rstrip("0").rstrip(".")
        return mc.format_scalar(raw, fmt)

    def _poll_tick(self) -> None:
        if not self._poll_active or self._poll_pending is not None:
            return
        if self._state != "open" or self._worker is None or not self._poll_slaves:
            self._poll_abort()
            self._apply_state()
            return
        slave = self._poll_slaves[self._poll_idx]
        tid = self._next_tid()
        frame = mc.build_read_request(
            self._protocol, slave, self._poll_func, self._poll_base, self._poll_count, tid
        )
        self._poll_pending = (self._poll_idx, slave, tid)
        self._poll_buf.clear()
        if not self._worker.send(frame):
            self._poll_abort()
            self._apply_state()
            return
        self._poll_timeout.start(1500)

    def _poll_feed(self, data: bytes) -> None:
        if not self._poll_active or self._poll_pending is None:
            return
        self._poll_buf.extend(data)
        index, slave, tid = self._poll_pending
        frame, consumed = mc.extract_response(
            self._protocol, self._poll_buf, slave, self._poll_func, tid
        )
        if frame is None:
            return
        del self._poll_buf[:consumed]
        self._poll_timeout.stop()
        self._poll_pending = None
        result = mc.decode_read_response(
            self._protocol, frame, slave, self._poll_func, self._poll_count
        )
        self._apply_poll_result(index, result)
        self._poll_idx = (index + 1) % len(self._poll_slaves)

    def _apply_poll_result(self, index: int, result: dict) -> None:
        base_row = index * self._poll_count
        stamp = time.strftime("%H:%M:%S")
        if result.get("ok"):
            registers = result.get("registers", [])
            for offset in range(self._poll_count):
                row = base_row + offset
                if row >= self.poll_table.rowCount() or offset >= len(registers):
                    continue
                address = self._poll_base + offset
                item = self.poll_table.item(row, 4)
                item.setText(self._value_text(address, registers[offset]))
                item.setToolTip(f"{stamp} 更新")
        else:
            message = result.get("error", "读取失败")
            for offset in range(self._poll_count):
                row = base_row + offset
                if row >= self.poll_table.rowCount():
                    continue
                item = self.poll_table.item(row, 4)
                item.setText(message)
                item.setToolTip(message)

    def _on_poll_timeout(self) -> None:
        if not self._poll_active or self._poll_pending is None:
            return
        index, _slave, _tid = self._poll_pending
        self._poll_pending = None
        base_row = index * self._poll_count
        for offset in range(self._poll_count):
            row = base_row + offset
            if row < self.poll_table.rowCount():
                self.poll_table.item(row, 4).setText("超时")
        self._poll_idx = (index + 1) % len(self._poll_slaves)

    def _poll_stop(self) -> None:
        self._poll_abort()
        self.poll_state.setText("空闲")
        self._apply_state()
        self._set_status("监测已停止")

    def _poll_abort(self) -> None:
        self._poll_active = False
        self._poll_timer.stop()
        self._poll_timeout.stop()
        self._poll_pending = None
        self._poll_buf.clear()
        if hasattr(self, "poll_state"):
            self.poll_state.setText("空闲")

    def _clear_poll(self) -> None:
        if self._poll_active:
            self._poll_stop()
        self.poll_table.setRowCount(0)

    # ------------------------------------------------------------------
    # Excel 点表
    # ------------------------------------------------------------------
    @staticmethod
    def _parse_address(value) -> int:
        if isinstance(value, str):
            text = value.strip().lower()
            if text.startswith("0x"):
                return int(text, 16)
            return int(float(text))
        return int(value)

    @staticmethod
    def _find_column(header: list[str], *names: str) -> int:
        for index, text in enumerate(header):
            if any(name in text for name in names):
                return index
        return -1

    def _import_points(self) -> None:
        if load_workbook is None:
            QMessageBox.warning(self, "缺少依赖", "导入点表需要 openpyxl：\n\npip install openpyxl")
            return
        path, _ = QFileDialog.getOpenFileName(self, "导入点表", "", "Excel 文件 (*.xlsx)")
        if not path:
            return
        try:
            workbook = load_workbook(path, read_only=True, data_only=True)
            sheet = workbook.active
            rows = list(sheet.iter_rows(values_only=True))
        except Exception as exc:
            QMessageBox.warning(self, "导入失败", str(exc))
            return
        if not rows:
            return

        header = [str(cell or "") for cell in rows[0]]
        i_addr = self._find_column(header, "寄存器地址", "地址")
        i_name = self._find_column(header, "点位名称", "名称")
        i_factor = self._find_column(header, "换算系数", "系数")
        i_unit = self._find_column(header, "单位")
        if i_addr < 0:
            QMessageBox.warning(self, "导入失败", "未找到「寄存器地址」列")
            return

        points: dict[int, dict] = {}
        for row in rows[1:]:
            if not row:
                continue
            try:
                address = self._parse_address(row[i_addr])
            except Exception:
                continue
            name = str(row[i_name] or "") if 0 <= i_name < len(row) else ""
            factor = 1.0
            if 0 <= i_factor < len(row):
                try:
                    factor = float(row[i_factor] or 1)
                except Exception:
                    factor = 1.0
            unit = str(row[i_unit] or "") if 0 <= i_unit < len(row) else ""
            points[address] = {"name": name, "factor": factor, "unit": unit}

        self._poll_points = points
        self._save_poll_config()
        self._set_status(f"点表导入完成，共 {len(points)} 个点位")
        if self._poll_active:
            self._build_poll_table()

    def _export_excel(self) -> None:
        if Workbook is None:
            QMessageBox.warning(self, "缺少依赖", "导出 Excel 需要 openpyxl：\n\npip install openpyxl")
            return
        if self.poll_table.rowCount() == 0:
            self._set_status("没有可导出的数据", error=True)
            return
        path, _ = QFileDialog.getSaveFileName(self, "导出监测数据", "modbus_data.xlsx", "Excel 文件 (*.xlsx)")
        if not path:
            return
        try:
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "监测数据"
            headers = [
                self.poll_table.horizontalHeaderItem(col).text()
                for col in range(self.poll_table.columnCount())
            ]
            sheet.append(headers)
            for row in range(self.poll_table.rowCount()):
                sheet.append([
                    self.poll_table.item(row, col).text() if self.poll_table.item(row, col) else ""
                    for col in range(self.poll_table.columnCount())
                ])
            workbook.save(path)
        except Exception as exc:
            QMessageBox.warning(self, "导出失败", str(exc))
            return
        self._set_status(f"已导出：{path}")

    # ==================================================================
    # 其他
    # ==================================================================
    def _set_status(self, text: str, error: bool = False) -> None:
        self._status_is_error = error
        self.status_label.setText(text)
        self.status_label.setStyleSheet("color: #f87171;" if error else "")

    def closeEvent(self, event) -> None:
        self._save_serial_config()
        self._save_poll_config()
        self._port_timer.stop()
        for timer in (
            self._frame_timer,
            self._timed_timer,
            self._scan_wait,
            self._scan_timeout,
            self._poll_timer,
            self._poll_timeout,
        ):
            timer.stop()
        worker = self._worker
        self._worker = None
        if worker is not None:
            worker.blockSignals(True)
            worker.stop()
            if not worker.wait(2500):
                worker.terminate()
                worker.wait(500)
        event.accept()


def main() -> int:
    app = QApplication(sys.argv)
    app.setApplicationName("串口调试助手")
    app.setStyle("Fusion")
    font = QFont()
    font.setFamilies(["Microsoft YaHei UI", "Microsoft YaHei", "Segoe UI"])
    font.setPointSize(9)
    app.setFont(font)

    icon_path = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent)) / "icon.ico"
    if icon_path.exists():
        app.setWindowIcon(QIcon(str(icon_path)))

    window = MainWindow()
    window.show()
    return app.exec()


if __name__ == "__main__":
    sys.exit(main())