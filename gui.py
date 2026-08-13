#!/usr/bin/env python3
"""扫描版 PDF -> 可编辑 Word / Excel  图形界面 (PySide6, 跨平台)

运行:  python3 gui.py
打包:  pyinstaller pdfrec-gui.spec --noconfirm
"""
import os, sys, time, collections, subprocess, traceback

from PySide6.QtCore import Qt, QThread, Signal, QObject, QSize
from PySide6.QtGui import QPixmap, QAction, QIcon, QFont
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QSplitter,
    QListWidget, QListWidgetItem, QPushButton, QLabel, QProgressBar, QPlainTextEdit,
    QFileDialog, QGroupBox, QFormLayout, QComboBox, QCheckBox, QSlider,
    QDoubleSpinBox, QLineEdit, QMessageBox, QToolBar, QStatusBar, QTabWidget,
    QAbstractItemView, QSizePolicy, QDialog, QDialogButtonBox, QMenu,
)

import convert as core

APP_NAME = '扫描件转 Word'
STATUS = {'wait': '待转换', 'run': '转换中', 'ok': '完成', 'err': '失败', 'skip': '已取消'}
FMT_LABEL = {'docx': 'Word', 'xlsx': 'Excel(只导表格)', 'both': 'Word + Excel'}


# ============================ 错误详情弹窗 ============================
class ErrorDialog(QDialog):
    """出错时不闪退, 把完整信息摊开给用户: 可复制、可存盘

    没有这个窗口, 打包后的 GUI 崩溃只会静悄悄消失, 用户什么线索都拿不到。
    """

    def __init__(self, parent, summary, detail):
        super().__init__(parent)
        self.setWindowTitle(f'{APP_NAME} — 出错了')
        self.resize(820, 480)
        self.detail = detail

        v = QVBoxLayout(self)
        head = QLabel(summary)
        head.setWordWrap(True)
        head.setStyleSheet('font-weight:bold; color:#b02020;')
        v.addWidget(head)
        v.addWidget(QLabel('程序没有退出, 其余文件仍可继续转换。下面是详细信息:'))

        self.txt = QPlainTextEdit(detail)
        self.txt.setReadOnly(True)
        self.txt.setFont(QFont('Menlo' if sys.platform == 'darwin' else 'Consolas', 10))
        self.txt.setLineWrapMode(QPlainTextEdit.NoWrap)
        v.addWidget(self.txt, 1)

        bb = QDialogButtonBox()
        b_copy = bb.addButton('复制全部', QDialogButtonBox.ActionRole)
        b_save = bb.addButton('保存为文件...', QDialogButtonBox.ActionRole)
        bb.addButton('关闭', QDialogButtonBox.AcceptRole)
        b_copy.clicked.connect(self.copy)
        b_save.clicked.connect(self.save)
        bb.accepted.connect(self.accept)
        v.addWidget(bb)

    def copy(self):
        QApplication.clipboard().setText(self.detail)

    def save(self):
        d = time.strftime('%Y%m%d_%H%M%S')
        p, _ = QFileDialog.getSaveFileName(self, '保存错误详情',
                                           os.path.expanduser(f'~/pdfrec_error_{d}.txt'),
                                           '文本文件 (*.txt)')
        if p:
            with open(p, 'w', encoding='utf-8') as f:
                f.write(self.detail)


class CrashBridge(QObject):
    """sys.excepthook 可能在任意线程触发; 用信号把它排队回主线程再弹窗"""
    raised = Signal(str, str)


_BRIDGE = None


def install_excepthook(window):
    """兜底: 任何没被接住的异常都变成弹窗, 而不是让解释器直接终止进程"""
    global _BRIDGE
    _BRIDGE = CrashBridge()
    _BRIDGE.raised.connect(lambda s, d: window.show_error(s, d))

    def hook(etype, value, tb):
        if issubclass(etype, KeyboardInterrupt):
            sys.__excepthook__(etype, value, tb); return
        text = ''.join(traceback.format_exception(etype, value, tb))
        sys.__excepthook__(etype, value, tb)        # 同时留一份在控制台
        _BRIDGE.raised.emit(f'{etype.__name__}: {value}', text)

    sys.excepthook = hook


# ============================ 后台工作线程 ============================
class Worker(QThread):
    """在后台线程跑转换; 所有 UI 更新通过 signal 回主线程"""
    file_started = Signal(int, str)                 # index, name
    progressed   = Signal(int, str, int, int, str)  # index, stage, cur, total, msg
    page_ready   = Signal(str)                      # 当前页 PNG 路径(用于预览)
    logged       = Signal(str)                      # 引擎里的每一条执行日志
    file_done    = Signal(int, str)                 # index, 输出路径
    file_failed  = Signal(int, str, str)            # index, 摘要, traceback
    crashed      = Signal(str, str)                 # 摘要, traceback
    all_done     = Signal(int, int, list)           # 成功数, 总数, 页级错误列表

    def __init__(self, files, fmts, cfg, out_dir):
        super().__init__()
        self.files, self.fmts = files, fmts        # fmts 与 files 一一对应
        self.cfg, self.out_dir = cfg, out_dir
        self._stop = False
        self.ok = 0
        self.problems = []

    def stop(self):
        self._stop = True

    def run(self):
        try:
            self._convert_all()
        except Exception:
            # 兜底: 线程里任何漏网异常都只报告, 不让进程死掉
            self.crashed.emit('转换线程异常终止', traceback.format_exc())
        self.all_done.emit(self.ok, len(self.files), self.problems)

    def _convert_all(self):
        for i, path in enumerate(self.files):
            if self._stop:
                break
            name = os.path.basename(path)
            self.file_started.emit(i, name)
            errs = []
            try:
                def prog(stage, cur, total, msg, _i=i):
                    self.progressed.emit(_i, stage, cur, total, msg)
                    if stage == 'render' and cur % 3 == 0:
                        png = os.path.join(core.CACHE, 'pages',
                                           f'{core._last_tag}_{cur:03d}.png')
                        if os.path.exists(png):
                            self.page_ready.emit(png)

                outs = core.convert(path, progress=prog, stop=lambda: self._stop,
                                    out_dir=self.out_dir or None, cfg=self.cfg,
                                    errors=errs, log=self.logged.emit,
                                    fmt=self.fmts[i])
                self.file_done.emit(i, ' , '.join(outs))
                self.ok += 1
            except core.Cancelled:
                break
            except Exception as e:
                self.file_failed.emit(i, f'{type(e).__name__}: {e}',
                                      traceback.format_exc())
            finally:
                self.problems += [dict(e, file=name) for e in errs]


# ============================ 参数面板 ============================
class SettingsPanel(QWidget):
    """把 core.DEFAULTS 暴露成可调控件"""

    def __init__(self):
        super().__init__()
        tabs = QTabWidget()
        tabs.addTab(self._tab_basic(), '常用')
        tabs.addTab(self._tab_layout(), '版面')
        tabs.addTab(self._tab_font(), '字体')
        lay = QVBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.addWidget(tabs)

    # ---- 常用 ----
    def _tab_basic(self):
        w = QWidget(); f = QFormLayout(w)
        self.cb_quality = QComboBox()
        self.cb_quality.addItem('快速 (1600 px, 约快 1 倍)', 1600)
        self.cb_quality.addItem('标准 (2560 px, 推荐)', 2560)
        self.cb_quality.addItem('高精度 (3200 px, 小字更准)', 3200)
        self.cb_quality.setCurrentIndex(1)
        f.addRow('识别精度', self.cb_quality)

        self.ck_marker = QCheckBox('插入「—— 原第 N 页 ——」分隔标记')
        self.ck_marker.setChecked(True)
        self.ck_marker.setToolTip('方便对照原件核对; 定稿后可在 Word 中批量删除')
        f.addRow('', self.ck_marker)

        self.ck_header = QCheckBox('剔除页眉'); self.ck_header.setChecked(True)
        self.ck_footer = QCheckBox('剔除页脚 / 页码'); self.ck_footer.setChecked(True)
        self.ck_stamp  = QCheckBox('剔除印章 / 手写签名噪声'); self.ck_stamp.setChecked(True)
        self.ck_stamp.setToolTip('骑缝章常被识别成乱码混进正文, 建议保持勾选')
        for c in (self.ck_header, self.ck_footer, self.ck_stamp):
            f.addRow('', c)
        return w

    # ---- 版面 ----
    def _tab_layout(self):
        w = QWidget(); f = QFormLayout(w)
        self.ck_table = QCheckBox('把多列版面还原成无边框表格')
        self.ck_table.setChecked(True)
        self.ck_table.setToolTip('列数由页面自动判定(不限两列); 技术规格、供货清单类文档\n'
                                 '强烈建议开启, 否则改字时排版会散')
        f.addRow('', self.ck_table)

        self.ck_grid = QCheckBox('照框线还原表格(含合并单元格)')
        self.ck_grid.setChecked(True)
        self.ck_grid.setToolTip('原件画了框线的表格按框线还原行列, 原本合并的格子保持合并;\n'
                                '关掉则一律按文字间距猜列, 格内多行文字会被拆成多行')
        f.addRow('', self.ck_grid)

        self.sp_col = self._spin(0.015, 0.120, 0.035, '列间空白',
                                 '行内横向空白超过此比例(占页宽)才算跨列。\n'
                                 '该成表格的没成 -> 调低; 正文被误判成表格 -> 调高')
        self.sp_full = self._spin(0.60, 0.95, 0.78, '续行判定(排满比例)',
                                  '上一行右端超过此比例才可能有续行;\n段落被错误粘连时调高, 该合的没合时调低')
        self.sp_bullet = self._spin(0.010, 0.080, 0.030, '项目符号缩进',
                                    '缩进超出正文基准多少视为列表项')
        self.sp_stamp = self._spin(0.50, 0.99, 0.88, '印章置信度阈值',
                                   '低于此置信度的短字符串按噪声剔除')
        for sp, label in ((self.sp_col, '列间空白'), (self.sp_full, '续行判定'),
                          (self.sp_bullet, '项目符号缩进'), (self.sp_stamp, '印章阈值')):
            f.addRow(label, sp)

        self.ck_debug = QCheckBox('生成版面调试图')
        self.ck_debug.setToolTip('在 PDF 同目录 .pdfrec_cache/debug/ 下画出每页的识别框与块划分,\n'
                                 '一眼看出哪行被当噪声剔除了, 调上面这些阈值时用')
        f.addRow('', self.ck_debug)

        btn = QPushButton('恢复默认值'); btn.clicked.connect(self.reset)
        f.addRow('', btn)
        return w

    def _spin(self, lo, hi, val, name, tip):
        sp = QDoubleSpinBox()
        sp.setRange(lo, hi); sp.setSingleStep(0.005); sp.setDecimals(3); sp.setValue(val)
        sp.setToolTip(tip)
        return sp

    # ---- 字体 ----
    def _tab_font(self):
        w = QWidget(); f = QFormLayout(w)
        self.cb_zh = QComboBox(); self.cb_zh.setEditable(True)
        self.cb_zh.addItems(['宋体', '仿宋', '楷体', '黑体', '微软雅黑', 'PingFang SC'])
        self.cb_en = QComboBox(); self.cb_en.setEditable(True)
        self.cb_en.addItems(['Times New Roman', 'Arial', 'Calibri', 'Georgia'])
        self.sp_size = QDoubleSpinBox()
        self.sp_size.setRange(7.0, 16.0); self.sp_size.setSingleStep(0.5); self.sp_size.setValue(10.5)
        f.addRow('中文字体', self.cb_zh)
        f.addRow('西文字体', self.cb_en)
        f.addRow('正文字号 (pt)', self.sp_size)
        return w

    def reset(self):
        self.sp_col.setValue(core.DEFAULTS['gutter'])
        self.sp_full.setValue(core.DEFAULTS['full_line'])
        self.sp_bullet.setValue(core.DEFAULTS['bullet_ind'])
        self.sp_stamp.setValue(core.DEFAULTS['stamp_conf'])
        self.ck_debug.setChecked(core.DEFAULTS['debug'])
        self.ck_grid.setChecked(core.DEFAULTS['grid_tables'])

    def cfg(self):
        return {
            'long_edge':   self.cb_quality.currentData(),
            'gutter':      self.sp_col.value(),
            'full_line':   self.sp_full.value(),
            'bullet_ind':  self.sp_bullet.value(),
            'stamp_conf':  self.sp_stamp.value(),
            'drop_header': self.ck_header.isChecked(),
            'drop_footer': self.ck_footer.isChecked(),
            'drop_stamp':  self.ck_stamp.isChecked(),
            'page_marker': self.ck_marker.isChecked(),
            'tables':      self.ck_table.isChecked(),
            'grid_tables': self.ck_grid.isChecked(),
            'debug':       self.ck_debug.isChecked(),
            'zh_font':     self.cb_zh.currentText().strip() or '宋体',
            'en_font':     self.cb_en.currentText().strip() or 'Times New Roman',
            'font_size':   self.sp_size.value(),
        }


# ============================ 文件队列 ============================
class FileList(QListWidget):
    """待转队列; 每个文件各自带一个目标格式(默认 Word)"""
    files_added = Signal(int)

    def __init__(self):
        super().__init__()
        self.setAcceptDrops(True)
        self.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.paths, self.fmts, self.states = [], [], []
        self.default_fmt = 'docx'
        self.setContextMenuPolicy(Qt.CustomContextMenu)
        self.customContextMenuRequested.connect(self._menu)
        self.itemDoubleClicked.connect(
            lambda it: self.cycle_fmt([self.row(it)]))

    def dragEnterEvent(self, e):
        if e.mimeData().hasUrls():
            e.acceptProposedAction()

    dragMoveEvent = dragEnterEvent

    def dropEvent(self, e):
        added = 0
        for url in e.mimeData().urls():
            p = url.toLocalFile()
            if os.path.isdir(p):
                for f in sorted(os.listdir(p)):
                    if f.lower().endswith('.pdf'):
                        added += self.add(os.path.join(p, f))
            elif p.lower().endswith('.pdf'):
                added += self.add(p)
        if added:
            self.files_added.emit(added)
        e.acceptProposedAction()

    def add(self, path, fmt=None):
        path = os.path.abspath(path)
        if path in self.paths:
            return 0
        self.paths.append(path)
        self.fmts.append(fmt or self.default_fmt)
        self.states.append(('wait', ''))
        it = QListWidgetItem()
        it.setToolTip(f'{path}\n\n双击可切换目标格式, 右键可批量指定')
        self.addItem(it)
        self._render(len(self.paths) - 1)
        return 1

    def _render(self, i):
        key, extra = self.states[i]
        self.item(i).setText(f'{os.path.basename(self.paths[i])}   →  '
                             f'{FMT_LABEL[self.fmts[i]]}\n    {STATUS[key]}{extra}')

    def set_status(self, i, key, extra=''):
        if 0 <= i < self.count():
            self.states[i] = (key, extra)
            self._render(i)

    def rows_or_all(self):
        """没选中任何一项时, 动作作用于全部 —— 批量改格式的常见意图"""
        return sorted(self.row(it) for it in self.selectedItems()) or \
            list(range(self.count()))

    def set_fmt(self, fmt, rows=None):
        for r in (self.rows_or_all() if rows is None else rows):
            self.fmts[r] = fmt
            self._render(r)

    def cycle_fmt(self, rows):
        order = list(FMT_LABEL)
        for r in rows:
            self.fmts[r] = order[(order.index(self.fmts[r]) + 1) % len(order)]
            self._render(r)

    def _menu(self, pos):
        it = self.itemAt(pos)
        if it is None and not self.count():
            return
        rows = sorted(self.row(x) for x in self.selectedItems())
        if it is not None and self.row(it) not in rows:
            rows = [self.row(it)]
        m = QMenu(self)
        for fmt, label in FMT_LABEL.items():
            m.addAction(f'转成 {label}').triggered.connect(
                lambda _=False, f=fmt, rs=rows or None: self.set_fmt(f, rs))
        m.exec(self.viewport().mapToGlobal(pos))

    def remove_selected(self):
        for it in sorted(self.selectedItems(), key=self.row, reverse=True):
            r = self.row(it)
            self.takeItem(r)
            del self.paths[r], self.fmts[r], self.states[r]

    def clear_all(self):
        self.clear()
        self.paths.clear(); self.fmts.clear(); self.states.clear()


# ============================ 日志区 ============================
class LogPanel(QGroupBox):
    """执行日志: 引擎每走一步都往这里写一行, 出问题时先看它"""

    def __init__(self):
        super().__init__('执行日志')
        self.txt = QPlainTextEdit()
        self.txt.setReadOnly(True)
        self.txt.setMaximumBlockCount(20000)     # 上限防止长批量把内存吃光
        self.txt.setMinimumHeight(130)
        self.txt.setLineWrapMode(QPlainTextEdit.NoWrap)
        self.txt.setFont(QFont('Menlo' if sys.platform == 'darwin' else 'Consolas', 10))

        self.ck_auto = QCheckBox('自动滚动'); self.ck_auto.setChecked(True)
        b_copy  = QPushButton('复制');  b_copy.clicked.connect(self.copy)
        b_save  = QPushButton('保存');  b_save.clicked.connect(self.save)
        b_clear = QPushButton('清空');  b_clear.clicked.connect(self.txt.clear)
        row = QHBoxLayout()
        row.addWidget(self.ck_auto)
        row.addStretch(1)
        for b in (b_copy, b_save, b_clear):
            b.setFixedWidth(56); row.addWidget(b)

        v = QVBoxLayout(self)
        v.setContentsMargins(6, 4, 6, 6); v.setSpacing(4)
        v.addLayout(row)
        v.addWidget(self.txt, 1)

    def append(self, msg):
        for line in str(msg).rstrip().splitlines() or ['']:
            self.txt.appendPlainText(f'{time.strftime("%H:%M:%S")}  {line}')
        if self.ck_auto.isChecked():
            self.txt.verticalScrollBar().setValue(self.txt.verticalScrollBar().maximum())

    def text(self):
        return self.txt.toPlainText()

    def copy(self):
        QApplication.clipboard().setText(self.text())

    def save(self):
        d = time.strftime('%Y%m%d_%H%M%S')
        p, _ = QFileDialog.getSaveFileName(self, '保存日志',
                                           os.path.expanduser(f'~/pdfrec_log_{d}.txt'),
                                           '文本文件 (*.txt)')
        if p:
            with open(p, 'w', encoding='utf-8') as f:
                f.write(self.text())


# ============================ 主窗口 ============================
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(APP_NAME)
        self.resize(1080, 800)
        self.worker = None
        self.details = []          # 累积的错误详情, 供「错误详情」按钮回看
        self._build()

    def _build(self):
        # --- 工具栏 ---
        tb = QToolBar(); tb.setIconSize(QSize(18, 18)); tb.setMovable(False)
        self.addToolBar(tb)
        self.act_add   = QAction('添加 PDF', self); self.act_add.triggered.connect(self.pick_files)
        self.act_del   = QAction('移除所选', self)
        self.act_clear = QAction('清空', self)
        self.act_run   = QAction('开始转换', self); self.act_run.triggered.connect(self.start)
        self.act_stop  = QAction('停止', self);     self.act_stop.triggered.connect(self.stop)
        self.act_open  = QAction('打开输出目录', self); self.act_open.triggered.connect(self.open_out)
        self.act_err   = QAction('错误详情', self); self.act_err.triggered.connect(self.show_details)
        self.act_stop.setEnabled(False)
        self.act_err.setEnabled(False)
        for a in (self.act_add, self.act_del, self.act_clear, None,
                  self.act_run, self.act_stop, None, self.act_open, self.act_err):
            tb.addSeparator() if a is None else tb.addAction(a)

        # --- 左: 文件队列 ---
        self.list = FileList()
        self.act_del.triggered.connect(self.list.remove_selected)
        self.act_clear.triggered.connect(self.list.clear_all)
        self.list.files_added.connect(lambda n: self.log(f'已添加 {n} 个文件'))
        left = QWidget(); lv = QVBoxLayout(left)
        lv.setContentsMargins(6, 6, 3, 6)
        tip = QLabel('把 PDF 或整个文件夹拖到下面')
        tip.setStyleSheet('color:#888;')
        # 目标格式逐文件可选: 这里选的是新加入文件的默认值, 「应用到所选」把
        # 它套到已在队列里的文件上(没选中就是全部); 双击/右键也能改单个文件
        self.cb_fmt = QComboBox()
        for fmt, label in FMT_LABEL.items():
            self.cb_fmt.addItem(label, fmt)
        self.cb_fmt.currentIndexChanged.connect(
            lambda: setattr(self.list, 'default_fmt', self.cb_fmt.currentData()))
        b_fmt = QPushButton('应用到所选')
        b_fmt.setToolTip('把上面选的格式套到列表中选中的文件; 未选中则应用到全部')
        b_fmt.clicked.connect(lambda: self.apply_fmt())
        fr = QHBoxLayout()
        fr.addWidget(QLabel('目标格式')); fr.addWidget(self.cb_fmt, 1); fr.addWidget(b_fmt)
        lv.addWidget(tip)
        lv.addLayout(fr)
        lv.addWidget(self.list)

        # --- 右上: 预览 ---
        self.preview = QLabel('转换开始后这里显示当前页')
        self.preview.setAlignment(Qt.AlignCenter)
        self.preview.setMinimumHeight(200)
        self.preview.setStyleSheet('background:#f4f4f4; color:#999; border:1px solid #ddd;')
        self.preview.setSizePolicy(QSizePolicy.Ignored, QSizePolicy.Expanding)

        # --- 右下: 参数 + 输出 ---
        self.settings = SettingsPanel()
        out_box = QGroupBox('输出')
        og = QHBoxLayout(out_box)
        self.ed_out = QLineEdit(); self.ed_out.setPlaceholderText('留空 = 与原 PDF 同目录')
        b_out = QPushButton('选择...'); b_out.clicked.connect(self.pick_out)
        og.addWidget(QLabel('目录')); og.addWidget(self.ed_out, 1); og.addWidget(b_out)

        right = QWidget(); rv = QVBoxLayout(right)
        rv.setContentsMargins(3, 6, 6, 6)
        rv.addWidget(self.preview, 3)
        rv.addWidget(self.settings, 2)
        rv.addWidget(out_box)

        split = QSplitter(Qt.Horizontal)
        split.addWidget(left); split.addWidget(right)
        split.setSizes([330, 750])

        # --- 底部: 进度 + 日志 (放进竖向 splitter, 日志区可随手拉大) ---
        self.bar_file = QProgressBar(); self.bar_file.setFormat('%p%  %v/%m')
        self.bar_all  = QProgressBar(); self.bar_all.setFormat('总进度 %v/%m')
        self.logs = LogPanel()

        bottom = QWidget(); bv = QVBoxLayout(bottom)
        bv.setContentsMargins(6, 0, 6, 6); bv.setSpacing(4)
        bv.addWidget(self.bar_file); bv.addWidget(self.bar_all); bv.addWidget(self.logs, 1)

        vsplit = QSplitter(Qt.Vertical)
        vsplit.addWidget(split); vsplit.addWidget(bottom)
        vsplit.setStretchFactor(0, 3); vsplit.setStretchFactor(1, 2)
        vsplit.setSizes([430, 320])          # 日志区嫌小就直接往上拖分隔条

        self.setCentralWidget(vsplit)
        self.setStatusBar(QStatusBar())
        self.statusBar().showMessage('就绪 —— 全部处理在本机完成, 不联网')
        self.setAcceptDrops(True)
        self.log('就绪。把 PDF 拖进左侧列表, 点「开始转换」。')

    # ---- 拖拽到窗口任意处 ----
    def dragEnterEvent(self, e):
        if e.mimeData().hasUrls():
            e.acceptProposedAction()

    dragMoveEvent = dragEnterEvent

    def dropEvent(self, e):
        self.list.dropEvent(e)

    # ---- 动作 ----
    def log(self, msg):
        self.logs.append(msg)

    def show_error(self, summary, detail):
        """统一的出错出口: 记日志 + 存详情 + 弹窗"""
        self.log(f'✗ {summary}')
        self.details.append(f'===== {time.strftime("%Y-%m-%d %H:%M:%S")}  {summary} =====\n{detail}')
        self.act_err.setEnabled(True)
        ErrorDialog(self, summary, detail).exec()

    def show_details(self):
        if not self.details:
            QMessageBox.information(self, APP_NAME, '目前没有错误记录。'); return
        ErrorDialog(self, f'共 {len(self.details)} 条错误记录',
                    '\n\n'.join(self.details)).exec()

    def pick_files(self):
        fs, _ = QFileDialog.getOpenFileNames(self, '选择扫描版 PDF', '', 'PDF 文件 (*.pdf)')
        n = sum(self.list.add(f) for f in fs)
        if n:
            self.log(f'已添加 {n} 个文件')

    def apply_fmt(self):
        fmt = self.cb_fmt.currentData()
        rows = self.list.rows_or_all()
        self.list.set_fmt(fmt, rows)
        if rows:
            self.log(f'{len(rows)} 个文件的目标格式设为 {FMT_LABEL[fmt]}')

    def pick_out(self):
        d = QFileDialog.getExistingDirectory(self, '选择输出目录')
        if d:
            self.ed_out.setText(d)

    def open_out(self):
        d = self.ed_out.text().strip()
        if not d and self.list.paths:
            d = os.path.dirname(self.list.paths[0])
        if not d or not os.path.isdir(d):
            QMessageBox.information(self, APP_NAME, '还没有输出目录。'); return
        if sys.platform == 'darwin':
            subprocess.Popen(['open', d])
        elif os.name == 'nt':
            os.startfile(d)                                   # noqa: S606
        else:
            subprocess.Popen(['xdg-open', d])

    def start(self):
        if not self.list.paths:
            QMessageBox.information(self, APP_NAME, '请先添加 PDF 文件。'); return
        self.act_run.setEnabled(False); self.act_stop.setEnabled(True)
        self.bar_all.setMaximum(len(self.list.paths)); self.bar_all.setValue(0)
        cfg = self.settings.cfg()
        tally = collections.Counter(self.list.fmts)
        self.log(f'开始转换 {len(self.list.paths)} 个文件 (精度 {cfg["long_edge"]} px, '
                 + ', '.join(f'{FMT_LABEL[f]} {n}' for f, n in tally.items()) + ')...')

        self.worker = Worker(list(self.list.paths), list(self.list.fmts),
                             cfg, self.ed_out.text().strip())
        self.worker.file_started.connect(self.on_start)
        self.worker.progressed.connect(self.on_prog)
        self.worker.page_ready.connect(self.on_page)
        self.worker.logged.connect(self.log)
        self.worker.file_done.connect(self.on_done)
        self.worker.file_failed.connect(self.on_fail)
        self.worker.crashed.connect(self.on_crash)
        self.worker.all_done.connect(self.on_all)
        self.worker.start()

    def stop(self):
        if self.worker:
            self.worker.stop()
            self.log('正在停止...')
            self.act_stop.setEnabled(False)

    # ---- 信号槽 ----
    def on_start(self, i, name):
        self.list.set_status(i, 'run')
        self.statusBar().showMessage(f'正在处理: {name}')

    def on_prog(self, i, stage, cur, total, msg):
        self.bar_file.setMaximum(max(total, 1)); self.bar_file.setValue(cur)
        self.statusBar().showMessage(msg)

    def on_page(self, png):
        pm = QPixmap(png)
        if not pm.isNull():
            self.preview.setPixmap(pm.scaled(self.preview.size(), Qt.KeepAspectRatio,
                                             Qt.SmoothTransformation))

    def on_done(self, i, out):
        self.list.set_status(i, 'ok')
        self.bar_all.setValue(self.bar_all.value() + 1)

    def on_fail(self, i, err, tb):
        self.list.set_status(i, 'err', f' — {err[:60]}')
        self.bar_all.setValue(self.bar_all.value() + 1)
        self.log(f'  ✗ {err}')
        # 批量转换中途不打断: 先攒着, 全部跑完再一次性弹出
        name = os.path.basename(self.list.paths[i]) if i < len(self.list.paths) else f'#{i+1}'
        self.details.append(f'===== {time.strftime("%Y-%m-%d %H:%M:%S")}  {name} =====\n{tb}')
        self.act_err.setEnabled(True)

    def on_crash(self, summary, tb):
        self.show_error(summary, tb)

    def on_all(self, ok, total, problems):
        self.act_run.setEnabled(True); self.act_stop.setEnabled(False)
        self.statusBar().showMessage(f'完成: {ok}/{total} 个文件')
        self.log(f'全部结束: 成功 {ok} / 共 {total}'
                 + (f', {len(problems)} 页局部失败' if problems else ''))
        if problems:
            lines = [f'{p["file"]}  第 {p["page"]} 页  {p["stage"]}: {p["error"]}'
                     for p in problems]
            self.details.append('===== 页级失败(已跳过, 其余内容正常输出) =====\n'
                                + '\n\n'.join(f'{l}\n{p["trace"]}'
                                              for l, p in zip(lines, problems)))
            self.act_err.setEnabled(True)
            for l in lines:
                self.log(f'  ! {l}')
        if self.details and (ok < total or problems):
            self.show_details()
        elif ok:
            QMessageBox.information(self, APP_NAME, f'转换完成: {ok}/{total} 个文件。')

    def closeEvent(self, e):
        if self.worker and self.worker.isRunning():
            self.worker.stop(); self.worker.wait(3000)
        e.accept()


def main():
    app = QApplication(sys.argv)
    app.setApplicationName(APP_NAME)
    w = MainWindow()
    install_excepthook(w)
    for arg in sys.argv[1:]:
        if arg.lower().endswith('.pdf'):
            w.list.add(arg)
    w.show()
    sys.exit(app.exec())


if __name__ == '__main__':
    main()
