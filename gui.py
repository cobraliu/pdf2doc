#!/usr/bin/env python3
"""扫描版 PDF -> 可编辑 Word  图形界面 (PySide6, 跨平台)

运行:  python3 gui.py
打包:  pyinstaller pdfrec-gui.spec --noconfirm
"""
import os, sys, subprocess, traceback

from PySide6.QtCore import Qt, QThread, Signal, QSize
from PySide6.QtGui import QPixmap, QAction, QIcon, QFont
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QSplitter,
    QListWidget, QListWidgetItem, QPushButton, QLabel, QProgressBar, QPlainTextEdit,
    QFileDialog, QGroupBox, QFormLayout, QComboBox, QCheckBox, QSlider,
    QDoubleSpinBox, QLineEdit, QMessageBox, QToolBar, QStatusBar, QTabWidget,
    QAbstractItemView, QSizePolicy,
)

import convert as core

APP_NAME = '扫描件转 Word'
STATUS = {'wait': '待转换', 'run': '转换中', 'ok': '完成', 'err': '失败', 'skip': '已取消'}


# ============================ 后台工作线程 ============================
class Worker(QThread):
    """在后台线程跑转换; 所有 UI 更新通过 signal 回主线程"""
    file_started = Signal(int, str)              # index, name
    progressed   = Signal(int, str, int, int, str)  # index, stage, cur, total, msg
    page_ready   = Signal(str)                   # 当前页 PNG 路径(用于预览)
    file_done    = Signal(int, str)              # index, 输出路径
    file_failed  = Signal(int, str)              # index, 错误信息
    all_done     = Signal(int, int)              # 成功数, 总数

    def __init__(self, files, cfg, out_dir):
        super().__init__()
        self.files, self.cfg, self.out_dir = files, cfg, out_dir
        self._stop = False

    def stop(self):
        self._stop = True

    def run(self):
        ok = 0
        for i, path in enumerate(self.files):
            if self._stop:
                break
            self.file_started.emit(i, os.path.basename(path))
            try:
                def prog(stage, cur, total, msg, _i=i):
                    self.progressed.emit(_i, stage, cur, total, msg)
                    if stage == 'render' and cur % 3 == 0:
                        png = os.path.join(core.CACHE, 'pages',
                                           f'{core._last_tag}_{cur:03d}.png')
                        if os.path.exists(png):
                            self.page_ready.emit(png)

                out = core.convert(path, progress=prog, stop=lambda: self._stop,
                                   out_dir=self.out_dir or None, cfg=self.cfg)
                self.file_done.emit(i, out)
                ok += 1
            except core.Cancelled:
                break
            except Exception as e:
                self.file_failed.emit(i, f'{type(e).__name__}: {e}')
                traceback.print_exc()
        self.all_done.emit(ok, len(self.files))


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
        self.ck_twocol = QCheckBox('把「标签—值」两列还原成无边框表格')
        self.ck_twocol.setChecked(True)
        self.ck_twocol.setToolTip('技术规格类文档强烈建议开启, 否则改字时排版会散')
        f.addRow('', self.ck_twocol)

        self.sp_col = self._spin(0.30, 0.75, 0.50, '两列分界(占页宽)',
                                 '右侧内容超过此比例视为“值”列')
        self.sp_full = self._spin(0.60, 0.95, 0.78, '续行判定(排满比例)',
                                  '上一行右端超过此比例才可能有续行;\n段落被错误粘连时调高, 该合的没合时调低')
        self.sp_bullet = self._spin(0.010, 0.080, 0.030, '项目符号缩进',
                                    '缩进超出正文基准多少视为列表项')
        self.sp_stamp = self._spin(0.50, 0.99, 0.88, '印章置信度阈值',
                                   '低于此置信度的短字符串按噪声剔除')
        for sp, label in ((self.sp_col, '两列分界'), (self.sp_full, '续行判定'),
                          (self.sp_bullet, '项目符号缩进'), (self.sp_stamp, '印章阈值')):
            f.addRow(label, sp)
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
        self.sp_col.setValue(core.DEFAULTS['col_split'])
        self.sp_full.setValue(core.DEFAULTS['full_line'])
        self.sp_bullet.setValue(core.DEFAULTS['bullet_ind'])
        self.sp_stamp.setValue(core.DEFAULTS['stamp_conf'])

    def cfg(self):
        return {
            'long_edge':   self.cb_quality.currentData(),
            'col_split':   self.sp_col.value(),
            'full_line':   self.sp_full.value(),
            'bullet_ind':  self.sp_bullet.value(),
            'stamp_conf':  self.sp_stamp.value(),
            'drop_header': self.ck_header.isChecked(),
            'drop_footer': self.ck_footer.isChecked(),
            'drop_stamp':  self.ck_stamp.isChecked(),
            'page_marker': self.ck_marker.isChecked(),
            'two_col':     self.ck_twocol.isChecked(),
            'zh_font':     self.cb_zh.currentText().strip() or '宋体',
            'en_font':     self.cb_en.currentText().strip() or 'Times New Roman',
            'font_size':   self.sp_size.value(),
        }


# ============================ 文件队列 ============================
class FileList(QListWidget):
    files_added = Signal(int)

    def __init__(self):
        super().__init__()
        self.setAcceptDrops(True)
        self.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.paths = []

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

    def add(self, path):
        path = os.path.abspath(path)
        if path in self.paths:
            return 0
        self.paths.append(path)
        it = QListWidgetItem(f'{os.path.basename(path)}\n    {STATUS["wait"]}')
        it.setToolTip(path)
        self.addItem(it)
        return 1

    def set_status(self, i, key, extra=''):
        if 0 <= i < self.count():
            name = os.path.basename(self.paths[i])
            self.item(i).setText(f'{name}\n    {STATUS[key]}{extra}')

    def remove_selected(self):
        for it in sorted(self.selectedItems(), key=self.row, reverse=True):
            r = self.row(it)
            self.takeItem(r)
            del self.paths[r]

    def clear_all(self):
        self.clear(); self.paths.clear()


# ============================ 主窗口 ============================
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(APP_NAME)
        self.resize(1080, 720)
        self.worker = None
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
        self.act_stop.setEnabled(False)
        for a in (self.act_add, self.act_del, self.act_clear, None,
                  self.act_run, self.act_stop, None, self.act_open):
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
        lv.addWidget(tip)
        lv.addWidget(self.list)

        # --- 右上: 预览 ---
        self.preview = QLabel('转换开始后这里显示当前页')
        self.preview.setAlignment(Qt.AlignCenter)
        self.preview.setMinimumHeight(240)
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

        # --- 底部: 进度 + 日志 ---
        self.bar_file = QProgressBar(); self.bar_file.setFormat('%p%  %v/%m')
        self.bar_all  = QProgressBar(); self.bar_all.setFormat('总进度 %v/%m')
        self.txt_log  = QPlainTextEdit(); self.txt_log.setReadOnly(True)
        self.txt_log.setMaximumHeight(120)
        self.txt_log.setFont(QFont('Menlo' if sys.platform == 'darwin' else 'Consolas', 10))

        bottom = QWidget(); bv = QVBoxLayout(bottom)
        bv.setContentsMargins(6, 0, 6, 6)
        bv.addWidget(self.bar_file); bv.addWidget(self.bar_all); bv.addWidget(self.txt_log)

        central = QWidget(); cv = QVBoxLayout(central)
        cv.setContentsMargins(0, 0, 0, 0)
        cv.addWidget(split, 1); cv.addWidget(bottom)
        self.setCentralWidget(central)
        self.setStatusBar(QStatusBar())
        self.statusBar().showMessage('就绪 —— 全部处理在本机完成, 不联网')
        self.setAcceptDrops(True)

    # ---- 拖拽到窗口任意处 ----
    def dragEnterEvent(self, e):
        if e.mimeData().hasUrls():
            e.acceptProposedAction()

    dragMoveEvent = dragEnterEvent

    def dropEvent(self, e):
        self.list.dropEvent(e)

    # ---- 动作 ----
    def log(self, msg):
        self.txt_log.appendPlainText(msg)

    def pick_files(self):
        fs, _ = QFileDialog.getOpenFileNames(self, '选择扫描版 PDF', '', 'PDF 文件 (*.pdf)')
        n = sum(self.list.add(f) for f in fs)
        if n:
            self.log(f'已添加 {n} 个文件')

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
        self.log(f'开始转换 {len(self.list.paths)} 个文件...')

        self.worker = Worker(list(self.list.paths), self.settings.cfg(),
                             self.ed_out.text().strip())
        self.worker.file_started.connect(self.on_start)
        self.worker.progressed.connect(self.on_prog)
        self.worker.page_ready.connect(self.on_page)
        self.worker.file_done.connect(self.on_done)
        self.worker.file_failed.connect(self.on_fail)
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
        if stage == 'ocr' and (cur == 1 or cur % 10 == 0):
            self.log(f'  {msg}')

    def on_page(self, png):
        pm = QPixmap(png)
        if not pm.isNull():
            self.preview.setPixmap(pm.scaled(self.preview.size(), Qt.KeepAspectRatio,
                                             Qt.SmoothTransformation))

    def on_done(self, i, out):
        self.list.set_status(i, 'ok')
        self.bar_all.setValue(self.bar_all.value() + 1)
        self.log(f'  ✓ {out}')

    def on_fail(self, i, err):
        self.list.set_status(i, 'err', f' — {err[:60]}')
        self.bar_all.setValue(self.bar_all.value() + 1)
        self.log(f'  ✗ {err}')

    def on_all(self, ok, total):
        self.act_run.setEnabled(True); self.act_stop.setEnabled(False)
        self.statusBar().showMessage(f'完成: {ok}/{total} 个文件')
        self.log(f'全部结束: 成功 {ok} / 共 {total}\n')
        if ok:
            QMessageBox.information(self, APP_NAME, f'转换完成: {ok}/{total} 个文件。')

    def closeEvent(self, e):
        if self.worker and self.worker.isRunning():
            self.worker.stop(); self.worker.wait(3000)
        e.accept()


def main():
    app = QApplication(sys.argv)
    app.setApplicationName(APP_NAME)
    w = MainWindow()
    for arg in sys.argv[1:]:
        if arg.lower().endswith('.pdf'):
            w.list.add(arg)
    w.show()
    sys.exit(app.exec())


if __name__ == '__main__':
    main()
