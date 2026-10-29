# -*- coding: utf-8 -*-
"""
宿委会查寝总分计算器 v2.5
------------------------------------------------------------
v2.5 新增：
  1. 第 1 格填 w 按回车 = 无人（分数列写「无人」）；填 s = 锁（分数列写「锁」）；
     这两种情况下等级列留空、不参与算分。
  2. 选了「使用 F 列 tag」但表里 tag 列为空时，会弹窗提醒，并回到「选择匹配方式」
     窗口让你换一种方式，不用重启程序、也不用重新选文件。
  3. 修掉一个老 bug：文件准备失败后，程序还会去操作已经关掉的窗口，于是多弹一个
     吓人的「程序出错」、日志里也看不出真正原因。现在只弹那条有用的提示，
     并且会把**真正的失败原因**写进 出错日志.txt。

v2.4 的功能（沿用）：界面上那句红色提醒改成了"用户可以自己编辑"的提示语句
  · 默认显示「（单击我编辑或修改提示语句）」，用鼠标左键单击它就会弹出小编辑窗；
  · 输入自己的提醒（例如"记得看看有没有写上厕所 -3"）后点"确定"，
    会保存到程序同目录的「自定义提示.txt」里，下次打开程序还是这句话；
  · 想改回默认：把编辑窗里的内容清空再点"确定"即可（配置文件会被删掉）；
  · 长度上限 30 个汉字（英文、数字、一般符号按半个汉字算，合计不超过 60），
    超出上限的字直接打不进去。

v2.3 的功能（沿用）：寝室号来源可选
  · 使用 F 列 tag（推荐）—— 沿用原来的做法，靠 F 列的 tag 找行
  · 由程序读取寝室号（测试中）—— 不依赖 F 列，直接从 A 列寝室号里解析出
    楼栋号+房号（例如 "桃花坪学生宿舍19栋201" → 19栋201），并兼容只输入
    房号（如 201）的写法

其余功能与 v2.2 相同：
  1. 选择"总表原件"Excel，程序自动在原件同目录的「副本」子文件夹里生成一份
     带时间戳的副本（该文件夹不存在会自动创建），之后所有读写都只针对副本，
     绝不修改原件，也不会再出现"副本 - 副本"这种套娃名字。
  2. 输入寝室号，匹配到行号后把得分写入 D 列。
  3. 每寝室最多 11 个"扣分绝对值"，总分 = 100 - 各项之和；
     输入 0 提前结束，填满 11 项自动结束。
  4. 等级（优秀/良好/合格/不合格）由程序生成，不再依赖 Excel 公式。
  5. 已有分数时弹窗确认是否覆盖，覆盖时分数和等级一起更新。
  6. 结算后等待按回车进入下一个寝室；输入 -1 结束并保存副本。
  7. 扣分按回车记录后，还可以用退格键（Backspace）把上一格撤销回来重填。
"""

import os
import re
import io
import sys
import stat
import time
import shutil
import datetime
import unicodedata
import tkinter as tk
from tkinter import filedialog, messagebox

import openpyxl

# ---------------------------------------------------------------- 常量配置
VERSION = "v2.5"
MAX_ITEMS = 11          # 纸质表共 11 个扣分词段
FULL_SCORE = 100        # 基准分
WAIT_ENTER = True       # 结算后等待手动按回车（True）；改为 False 则用 PAUSE_MS 自动继续
PAUSE_MS = 2000         # 自动继续模式下每次计算后的防连按暂停（毫秒）
D_FIRST_ROW = 2         # 数据起始行
ROOM_COL = 1            # A 列为寝室号
SCORE_COL = 4           # D 列为分数列
GRADE_COL = 5           # E 列为等级列
TAG_COL = 6             # F 列为 tag 字段
END_CODE = "-1"         # 结束程序
COPY_DIRNAME = "副本"    # 副本统一放在原件同目录的这个子文件夹里

# ---- 可编辑的提示语句（v2.4 新增）----
DEFAULT_HINT = "（单击我编辑或修改提示语句）"   # 用户还没自定义过时显示的占位提示
HINT_FILE = "自定义提示.txt"                    # 自定义提示保存在程序同目录的这个文件里
HINT_MAX_WIDTH = 60      # 提示语句最大"显示宽度"：一个汉字算 2，英文/数字/半角符号算 1
HINT_MAX_HANZI = 30      # 也就是最多 30 个汉字（用于提示文案）
HINT_COLOR_PLACEHOLDER = "#888"   # 还没自定义时：灰色（表示这只是个占位提示）
HINT_COLOR_CUSTOM = "#c00"        # 用户自定义之后：红色（保持原来"红字提醒"的效果）

# ---- 项目地址（界面最底下一行，旁边有「复制网址」按钮）----
GITHUB_URL = "https://github.com/5baiyifeng/DormitoryScoreCalculator"

# ---- 特殊情况标记（v2.5）：第 1 格填这些字母，就把对应文字写进分数列，等级留空 ----
MARK_NONE_KEY = "w"      # 输入 w = 无人
MARK_NONE_TEXT = "无人"
MARK_LOCK_KEY = "s"      # 输入 s = 锁
MARK_LOCK_TEXT = "锁"

# 寝室号来源
SOURCE_TAG = "tag"      # 使用 F 列 tag（推荐）
SOURCE_ROOM = "room"    # 由程序从 A 列寝室号解析（测试中）
SOURCE_TITLES = {SOURCE_TAG: "F列tag", SOURCE_ROOM: "A列寝室号"}

# 等级分档：分数下限（含）→ 等级名称，从高到低排列
GRADE_LEVELS = [
    (90, "优秀"),
    (80, "良好"),
    (60, "合格"),
    (0, "不合格"),
]

# ---------------- 免责声明（每次启动程序时都要先看并确认） ----------------
DISCLAIMER_TITLE = "免责声明"
AUTHOR_QQ = "3521228485"
DISCLAIMER_TEXT = (
    "【免责声明】\n\n"
    "本程序是作者第一次动手写的小工具，可能还存在没有发现的 Bug，\n"
    "计算结果仅供参考。\n\n"
    "请务必自行核实程序生成的表格内容（分数、等级、寝室号是否对得上），\n"
    "确认无误后再使用或上缴。\n\n"
    "如果发现程序出错、数字不对，或者有改进建议，\n"
    "欢迎加 QQ 反馈：%s\n\n"
    "点击“我已知晓”或直接按回车，开始使用程序。" % AUTHOR_QQ
)

FONT_UI = ("Microsoft YaHei UI", 11)
FONT_BIG = ("Microsoft YaHei UI", 16, "bold")
FONT_MONO = ("Consolas", 14)


# ---------------------------------------------------------------- 工具函数
def norm_tag(value):
    """把值统一成可比较的字符串：数字 201 与文本 '201' 都能对上。"""
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def cell_has_value(value):
    """判断单元格是否已有内容。"""
    if value is None:
        return False
    if isinstance(value, str):
        return value.strip() != ""
    return True


def grade_of(score):
    """按分数算出等级；分数为空则返回 None（不写等级）。"""
    if score is None:
        return None
    try:
        value = float(score)
    except (TypeError, ValueError):
        return None
    for lower, name in GRADE_LEVELS:
        if value >= lower:
            return name
    return None


def room_keys(value):
    """
    从 A 列寝室号文本里解析出可用来匹配的键。
    例："桃花坪学生宿舍19栋201" → {"19栋201", "19-201", "201"}
        "3栋102"                → {"3栋102", "3-102", "102"}
    返回按"精确优先"排序的键列表；解析不出来时返回空列表。
    """
    if value is None:
        return []
    text = str(value).strip()
    if not text:
        return []
    keys = []

    # 楼栋号：阿拉伯数字 + 栋/号楼；房号：结尾的 2~4 位数字
    m = re.search(r"(\d{1,3})\s*(?:栋|号楼|幢|舍)", text)
    tail = re.search(r"(\d{2,4})\s*$", text)
    if m and tail:
        building, room = m.group(1), tail.group(1)
        for key in ("%s栋%s" % (building, room), "%s-%s" % (building, room)):
            if key not in keys:
                keys.append(key)
    if tail:
        room = tail.group(1)
        if room not in keys:
            keys.append(room)
    if not keys:
        keys.append(text)
    return keys


class NoRoomFound(ValueError):
    """
    表里一个寝室号都找不到（例如选了「使用 F 列 tag」，但 F 列是空的）。
    继承 ValueError，原来按 ValueError 捕获的地方不受影响；
    单独分出来是为了让程序能判断"这种情况可以换个匹配方式重来"。
    """


def special_mark_of(text):
    """
    把用户在第 1 格输入的 w / s 换成要写进分数列的文字。
    不是这两种就返回 None（当普通数字处理）。
    """
    key = (text or "").strip().lower()
    if key == MARK_NONE_KEY:
        return MARK_NONE_TEXT
    if key == MARK_LOCK_KEY:
        return MARK_LOCK_TEXT
    return None


def build_room_map(src_path):
    """
    读取原件，按 A 列寝室号建立 {匹配键: 行号}。
    同一个键出现多次时以第一次出现的行为准。
    返回 (工作表名, keys_map, {行号: (显示名, 键列表)}, 可匹配键总数)
    """
    wb = openpyxl.load_workbook(src_path, data_only=True, read_only=True)
    result = None
    for ws in wb.worksheets:
        if (ws.max_row or 0) < D_FIRST_ROW or (ws.max_column or 0) < ROOM_COL:
            continue
        keys_map = {}
        row_names = {}
        for row in range(D_FIRST_ROW, (ws.max_row or 0) + 1):
            raw = ws.cell(row, ROOM_COL).value
            keys = room_keys(raw)
            if not keys:
                continue
            name = str(raw).strip()
            row_names[row] = (name, keys)
            for key in keys:
                if key not in keys_map:
                    keys_map[key] = row
        if keys_map:
            result = (ws.title, keys_map, row_names, len(keys_map))
            break
    wb.close()
    if result is None:
        raise NoRoomFound(
            "没有在所选表格的 A 列（寝室号）中解析出任何寝室号。\n"
            "请确认所选文件的第一列是寝室号，例如“桃花坪学生宿舍19栋201”。"
        )
    return result


def load_tag_map(src_path):
    """
    读取原件（只读，用于取缓存值），返回 (工作表名, {归一化tag: 行号}, 来源说明)。
    tag 为公式时取 Excel 上次保存的缓存值。
    """
    wb = openpyxl.load_workbook(src_path, data_only=True, read_only=True)
    result = None
    for ws in wb.worksheets:
        if ws.max_column is None or ws.max_column < TAG_COL:
            continue
        tag_map = {}
        row = D_FIRST_ROW
        empty_streak = 0
        while row <= (ws.max_row or 0):
            key = norm_tag(ws.cell(row, TAG_COL).value)
            if key:
                empty_streak = 0
                if key not in tag_map:
                    tag_map[key] = row
            else:
                empty_streak += 1
                if empty_streak >= 20:      # 连续空白视为表末
                    break
            row += 1
        if tag_map:
            result = (ws.title, tag_map)
            break
    wb.close()
    if result is None:
        raise NoRoomFound(
            "没有在所选表格的 F 列（tag）中找到任何寝室号。\n"
            "请确认所选文件是含 tag 列的总表原件，或改用“由程序读取寝室号”。"
        )
    return result


def make_copy_path(src_path):
    """
    生成副本路径：原件同一文件夹下的「副本」子文件夹里，带时间戳，不覆盖已有文件。
    例：D:\\...\\第三学期\\我的\\副本\\原件 - 副本_20261015_2137.xlsx
    """
    folder = os.path.dirname(os.path.abspath(src_path))
    copy_dir = os.path.join(folder, COPY_DIRNAME)
    os.makedirs(copy_dir, exist_ok=True)      # 没有就自动建
    stem, ext = os.path.splitext(os.path.basename(src_path))
    if not ext:
        ext = ".xlsx"
    stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M")
    candidate = os.path.join(copy_dir, "%s - 副本_%s%s" % (stem, stamp, ext))
    seq = 2
    while os.path.exists(candidate):
        candidate = os.path.join(copy_dir, "%s - 副本_%s(%d)%s" % (stem, stamp, seq, ext))
        seq += 1
    return candidate


def ensure_writable(path):
    """
    去掉文件身上从别处带来的「只读」属性，返回是否设置成功。

    为什么需要这一步：
      1. 程序用 shutil.copy2 复制副本，它会**连文件属性一起复制**，
         所以「原件是只读」时，生成的副本也会是只读；
      2. 之后 openpyxl 往只读的副本里写数据，会被 Windows 直接拒绝，
         报 "[Errno 13] Permission denied"（Windows 上是"拒绝访问"）。
    常见的只读来源：从 U 盘 / 移动硬盘 / 共享文件夹 / 虚拟磁盘复制过来的文件，
    或从别的电脑拷来的表。清掉只读后就能正常写入。
    """
    try:
        os.chmod(path, stat.S_IWRITE)   # Windows 上只会清掉「只读」这一个属性
        return True
    except Exception:
        return False


# ---------------------------------------------------------------- 提示语句
def app_dir():
    """返回程序自己所在的目录（打包成 exe 后是 exe 所在目录，直接跑源码时是 .py 所在目录）。"""
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


def hint_width(text):
    """
    算一段文字的"显示宽度"：全角字符（汉字、中文标点）算 2，半角字符
    （英文字母、数字、一般符号、空格）算 1。
    例："宿舍201 -3" = 2+2+1+1+1+1+1+1 = 10
    """
    total = 0
    for ch in text:
        total += 2 if unicodedata.east_asian_width(ch) in ("W", "F") else 1
    return total


def trim_to_width(text, max_width):
    """从前往后保留字符，直到再加一个就会超过 max_width 为止。"""
    kept = []
    total = 0
    for ch in text:
        w = 2 if unicodedata.east_asian_width(ch) in ("W", "F") else 1
        if total + w > max_width:
            break
        kept.append(ch)
        total += w
    return "".join(kept)


def hint_file_path():
    """自定义提示保存的位置：程序同目录下的「自定义提示.txt」。"""
    return os.path.join(app_dir(), HINT_FILE)


def load_custom_hint():
    """
    读取用户自定义的提示语句。
    返回空字符串表示"用户还没自定义过"，此时界面显示默认占位提示。
    文件不存在、读不出来、内容超长，都按"没有自定义"或截断处理，绝不让程序报错。
    """
    try:
        path = hint_file_path()
        if not os.path.exists(path):
            return ""
        with io.open(path, "r", encoding="utf-8") as f:
            text = f.read()
        text = text.replace("\r", "").replace("\n", "").strip()
        return trim_to_width(text, HINT_MAX_WIDTH)
    except Exception:
        return ""


def save_custom_hint(text):
    """
    把自定义提示写进「自定义提示.txt」。
    text 为空字符串 = 恢复默认（直接把这个文件删掉）。
    返回 (是否成功, 出错信息)。
    """
    path = hint_file_path()
    try:
        if not text:
            if os.path.exists(path):
                os.remove(path)
            return True, ""
        with io.open(path, "w", encoding="utf-8") as f:
            f.write(text)
        return True, ""
    except Exception as exc:
        return False, str(exc)


# ---------------------------------------------------------------- 主窗口
class ScoreApp:
    def __init__(self, root, src_path, source=SOURCE_TAG):
        self.root = root
        self.src_path = os.path.abspath(src_path)
        self.source = source

        self.entries = []
        self.items = []          # 已录入的扣分
        self.locked = False      # 结算/等待确认期间锁定扣分格
        self.awaiting_enter = False   # 已结算，等待用户按回车进入下一个寝室
        self.pending_after = None
        self.last_committed_row = None   # 最近一次写入的行（用于撤销时清回空）
        self.undo_bound = False
        self.current_tag = None
        self.row_by_tag = {}
        self.room_rows = {}      # 由 A 列解析时的 {行号: (显示名, 键)}
        self.sheet_title = ""
        self.wb = None
        self.ws = None
        self.formula_rows_cleaned = 0   # 被程序改写过的等级公式行数
        self._copy_reset_job = None     # 「复制网址」按钮的“已复制”提示复位用
        self._committing = False        # 防止“失去焦点自动记录”和回车记录互相触发

        self.ready = False          # 文件准备成功、界面可用之后才置 True
        self.failed_stage = ""      # 失败发生在哪一步（供 main() 判断要不要重来）
        self.room_count = 0

        # 先准备文件：这一步失败就直接返回，不建界面、也不销毁窗口。
        # 以前是"先建界面、失败再 destroy"，结果 main() 还会去操作已经销毁的窗口，
        # 触发 TclError 并多弹一个「程序出错」——v2.5 修的就是这个问题。
        if not self._prepare_files():
            return

        self._build_ui()
        self._show_status(self.room_count)
        self.ready = True
        self.room_entry.focus_set()

    # ------------------------------------------------------------ 界面
    def _build_ui(self):
        self.root.title("宿委会查寝总分计算器 %s" % VERSION)
        self.root.resizable(False, False)
        self.root.protocol("WM_DELETE_WINDOW", self.on_quit)

        top = tk.Frame(self.root, padx=12, pady=10)
        top.pack(fill="x")

        tk.Label(top, text="寝室号（输入 -1 结束程序）：", font=FONT_UI).grid(
            row=0, column=0, sticky="w")
        self.room_var = tk.StringVar()
        self.room_entry = tk.Entry(top, textvariable=self.room_var, font=FONT_MONO,
                                   width=10, justify="center")
        self.room_entry.grid(row=0, column=1, sticky="w", padx=6)
        self.room_entry.bind("<Return>", self.on_room_enter)

        self.room_hint = tk.Label(top, text="", font=FONT_UI, fg="#0a6")
        self.room_hint.grid(row=0, column=2, sticky="w", padx=6)

        tk.Label(top, text="匹配方式：%s" % SOURCE_TITLES.get(self.source, self.source),
                 font=("Microsoft YaHei UI", 9), fg="#777").grid(row=1, column=0,
                                                                 columnspan=3, sticky="w")

        # 可编辑的提示语句（v2.4）：默认灰色占位提示，用户自定义过就显示成红色；
        # 鼠标左键单击这一行会弹出小编辑窗。
        self.hint_text = load_custom_hint()          # "" 表示用户还没自定义过
        self.hint_var = tk.StringVar(value=self.hint_text or DEFAULT_HINT)
        self.hint_label = tk.Label(
            self.root, textvariable=self.hint_var, font=FONT_UI,
            fg=HINT_COLOR_CUSTOM if self.hint_text else HINT_COLOR_PLACEHOLDER,
            cursor="hand2")
        self.hint_label.pack(anchor="w", padx=12)
        self.hint_label.bind("<Button-1>", self.on_hint_click)

        grid = tk.LabelFrame(self.root, text=" 扣分录入（输入 0 提前结束，填满 11 项自动结束） ",
                             font=FONT_UI, padx=10, pady=8)
        grid.pack(fill="x", padx=12, pady=8)

        for i in range(MAX_ITEMS):
            r, c = divmod(i, 4)
            tk.Label(grid, text="%d." % (i + 1), font=FONT_UI).grid(
                row=r, column=c * 2, sticky="e", padx=(6, 2), pady=4)
            var = tk.StringVar()
            vcmd = (self.root.register(self._validate_item), "%P", "%s", "%S",
                    str(i), id(var))
            ent = tk.Entry(grid, textvariable=var, font=FONT_MONO, width=5,
                           justify="center", state="disabled",
                           validate="key", validatecommand=vcmd)
            ent.grid(row=r, column=c * 2 + 1, sticky="w", padx=(0, 10), pady=4)
            ent.bind("<Return>", lambda e, idx=i: self.on_item_enter(idx))
            # 撤销走 Ctrl + 退格；普通退格留给输入框自己删字（v2.5 改）
            ent.bind("<Control-BackSpace>", lambda e: self.on_undo_key())
            # 用鼠标点走时也把这一格记下来（v2.5）：不习惯按回车的人可以一路点着录
            ent.bind("<FocusOut>", lambda e, idx=i: self.on_item_focus_out(idx))
            self.entries.append((var, ent))

        bottom = tk.Frame(self.root, padx=12, pady=6)
        bottom.pack(fill="x")
        self.score_var = tk.StringVar(value="本次得分：--")
        tk.Label(bottom, textvariable=self.score_var, font=FONT_BIG,
                 fg="#06c").pack(side="left")
        self.grade_var = tk.StringVar(value="")
        tk.Label(bottom, textvariable=self.grade_var, font=FONT_BIG,
                 fg="#0a6").pack(side="left", padx=12)
        tk.Button(bottom, text="结束并保存 (-1)", font=FONT_UI,
                  command=self.on_quit).pack(side="right")

        self.status_var = tk.StringVar(value="正在准备文件……")
        tk.Label(self.root, textvariable=self.status_var, font=("Microsoft YaHei UI", 9),
                 fg="#555", anchor="w", justify="left", wraplength=620).pack(
            fill="x", padx=12, pady=(0, 10))

        # 最底下这一行：项目地址 + 一键复制（v2.4）
        foot = tk.Frame(self.root, padx=12)
        foot.pack(fill="x", pady=(0, 10))
        tk.Label(foot, text="获取最新版请前往 %s" % GITHUB_URL,
                 font=("Microsoft YaHei UI", 9), fg="#777").pack(side="left")
        self.copy_btn = tk.Button(foot, text="复制网址", width=8,
                                  font=("Microsoft YaHei UI", 9),
                                  command=self.copy_url)
        self.copy_btn.pack(side="left", padx=8)

    # ------------------------------------------------------------ 复制网址
    def copy_url(self):
        """把项目网址写进系统剪贴板，并给个“已复制”的反馈。"""
        try:
            self.root.clipboard_clear()
            self.root.clipboard_append(GITHUB_URL)
            # update() 不能省：让 Windows 真正接管剪贴板内容，
            # 这样即使马上关掉程序，粘贴出来的还是这个网址
            self.root.update()
        except Exception as exc:
            messagebox.showwarning("复制失败", "没能写入剪贴板：\n%s" % exc)
            return

        self.copy_btn.configure(text="已复制")
        if self._copy_reset_job is not None:
            try:
                self.root.after_cancel(self._copy_reset_job)
            except Exception:
                pass
        self._copy_reset_job = self.root.after(1500, self._reset_copy_btn)

    def _reset_copy_btn(self):
        self._copy_reset_job = None
        try:
            self.copy_btn.configure(text="复制网址")
        except Exception:
            pass

    # ------------------------------------------------------------ 提示语句
    def on_hint_click(self, event=None):
        """鼠标左键单击提示语句 → 打开编辑窗。"""
        self.edit_hint()

    def _apply_hint(self, text):
        """把提示语句显示到界面上；text 为空表示恢复默认占位提示。"""
        self.hint_text = text
        self.hint_var.set(text or DEFAULT_HINT)
        self.hint_label.configure(
            fg=HINT_COLOR_CUSTOM if text else HINT_COLOR_PLACEHOLDER)

    def _validate_hint(self, proposed):
        """编辑时的实时校验：会超过宽度上限的那个字直接不让打进去。"""
        return hint_width(proposed) <= HINT_MAX_WIDTH

    def edit_hint(self):
        """弹出小编辑窗，让用户自定义主界面上那句提示。"""
        dlg = tk.Toplevel(self.root)
        dlg.title("编辑提示语句")
        dlg.resizable(False, False)
        dlg.transient(self.root)

        body = tk.Frame(dlg, padx=14, pady=12)
        body.pack(fill="both", expand=True)

        tk.Label(body, text="这句话会一直显示在主界面上，用来提醒自己别忘事。",
                 font=FONT_UI, anchor="w").pack(anchor="w")
        tk.Label(body,
                 text="上限 %d 个汉字（英文、数字、一般符号按半个汉字算，合计不超过 %d）。\n"
                      "把内容清空再点“确定”＝恢复默认提示。"
                      % (HINT_MAX_HANZI, HINT_MAX_WIDTH),
                 font=("Microsoft YaHei UI", 9), fg="#666", justify="left",
                 anchor="w").pack(anchor="w", pady=(4, 8))

        var = tk.StringVar(value=self.hint_text)
        vcmd = (dlg.register(self._validate_hint), "%P")
        entry = tk.Entry(body, textvariable=var, font=FONT_UI, width=36,
                         validate="key", validatecommand=vcmd)
        entry.pack(anchor="w")
        entry.select_range(0, "end")

        counter = tk.StringVar()
        tk.Label(body, textvariable=counter, font=("Microsoft YaHei UI", 9),
                 fg="#888", anchor="w").pack(anchor="w", pady=(4, 10))

        def refresh_counter(*_args):
            used = hint_width(var.get())
            counter.set("已用 %d / %d（还可以再输入 %d）"
                        % (used, HINT_MAX_WIDTH, HINT_MAX_WIDTH - used))

        var.trace_add("write", refresh_counter)
        refresh_counter()

        def on_ok(event=None):
            text = var.get().replace("\r", "").replace("\n", "").strip()
            text = trim_to_width(text, HINT_MAX_WIDTH)
            saved, err = save_custom_hint(text)
            self._apply_hint(text)
            dlg.destroy()
            if not saved:
                messagebox.showwarning(
                    "提示语句没能保存",
                    "提示语句在本次运行中已经生效，但没能写入文件：\n%s\n\n"
                    "下次打开程序会恢复成默认提示。\n"
                    "常见原因：程序所在的文件夹没有写入权限。\n\n"
                    "文件位置：%s" % (err, hint_file_path()))

        def on_cancel(event=None):
            dlg.destroy()

        btns = tk.Frame(body)
        btns.pack(anchor="e")
        tk.Button(btns, text="确定", font=FONT_UI, width=8,
                  command=on_ok).pack(side="left", padx=4)
        tk.Button(btns, text="取消", font=FONT_UI, width=8,
                  command=on_cancel).pack(side="left")

        dlg.bind("<Return>", on_ok)
        dlg.bind("<Escape>", on_cancel)

        # 摆到主窗口正中间偏上的位置，并抓住焦点
        dlg.update_idletasks()
        x = self.root.winfo_rootx() + max(
            0, (self.root.winfo_width() - dlg.winfo_width()) // 2)
        y = self.root.winfo_rooty() + 90
        dlg.geometry("+%d+%d" % (x, y))
        dlg.grab_set()
        entry.focus_force()

    # ------------------------------------------------------------ 文件准备
    def _prepare_files(self):
        """准备副本与工作簿。失败时弹窗 + 把真实原因写进日志，并设好 failed_stage。"""
        try:
            if self.source == SOURCE_ROOM:
                self.sheet_title, self.row_by_tag, self.room_rows, count = \
                    build_room_map(self.src_path)
            else:
                self.sheet_title, self.row_by_tag = load_tag_map(self.src_path)
                count = len(self.row_by_tag)
        except NoRoomFound as exc:
            # 这种情况可以换一种匹配方式重来：提醒一句，交给 main() 回到选择窗口
            log_prepare_failure("找不到寝室号", self.src_path, self.source, exc)
            if self.source == SOURCE_TAG:
                title = "tag 列为空"
                head = "所选表格的 F 列（tag）里一个寝室号都没有。"
            else:
                title = "A 列没有寝室号"
                head = "所选表格的 A 列里解析不出寝室号。"
            messagebox.showwarning(
                title,
                "%s\n\n点“确定”后可以重新选一种找寝室号的方式：\n"
                "· 使用 F 列 tag —— 表里必须真的有 tag 列\n"
                "· 由程序读取寝室号 —— 直接从 A 列解析，不需要 tag" % head)
            self.failed_stage = "no_room"
            return False
        except Exception as exc:
            log_prepare_failure("读取原件", self.src_path, self.source, exc)
            messagebox.showerror("读取失败", "读取原件失败：\n%s" % exc)
            self.failed_stage = "read"
            return False

        try:
            self.copy_path = make_copy_path(self.src_path)
        except Exception as exc:
            log_prepare_failure("建副本文件夹", self.src_path, self.source, exc)
            messagebox.showerror("复制失败", "无法创建副本文件夹：\n%s" % exc)
            self.failed_stage = "copy"
            return False
        try:
            shutil.copy2(self.src_path, self.copy_path)
        except Exception as exc:
            log_prepare_failure("复制副本", self.src_path, self.source, exc)
            messagebox.showerror("复制失败", "无法创建副本：\n%s" % exc)
            self.failed_stage = "copy"
            return False
        # 清掉从原件继承过来的「只读」属性，否则后面保存副本会被拒绝
        ensure_writable(self.copy_path)

        try:
            self.wb = openpyxl.load_workbook(self.copy_path)
            self.ws = self.wb[self.sheet_title]
        except Exception as exc:
            log_prepare_failure("打开副本", self.src_path, self.source, exc)
            messagebox.showerror("打开副本失败", str(exc))
            self.failed_stage = "open"
            return False

        # 等级改由程序生成：把副本里原有的等级公式就地替换成程序算出的等级
        self._take_over_grades()

        # 界面还没建，这里只记数量，_show_status 由 __init__ 在 _build_ui 之后调用
        self.room_count = count
        return True

    def _take_over_grades(self):
        """把副本 E 列的等级公式改成程序算出的等级文字（原公式不再保留）。"""
        ws = self.ws
        changed = 0
        last_row = max(self.row_by_tag.values()) if self.row_by_tag else D_FIRST_ROW
        for row in range(D_FIRST_ROW, last_row + 1):
            cell = ws.cell(row, GRADE_COL)
            value = cell.value
            if isinstance(value, str) and value.startswith("="):
                score = ws.cell(row, SCORE_COL).value
                cell.value = grade_of(score)      # 没有分数就留空
                changed += 1
        self.formula_rows_cleaned = changed
        if changed:
            self._save_copy()

    def _save_copy(self):
        try:
            self.wb.save(self.copy_path)
            return True
        except Exception as exc:
            first_error = exc

        # 第一次写失败：可能是「只读」属性（旧版本生成的副本）、或杀毒软件
        # 正在扫描这个新文件把它短暂锁住了。清一次只读、稍等一下再试一次。
        ensure_writable(self.copy_path)
        time.sleep(0.3)
        try:
            self.wb.save(self.copy_path)
            return True
        except Exception as exc:
            messagebox.showerror(
                "保存失败",
                "写入副本失败：\n%s\n\n（第一次的报错：%s）\n"
                "重试一次仍然失败，常见原因：\n"
                "  1. 这个副本正被 Excel / WPS 打开着，关掉它再试；\n"
                "  2. 所在文件夹没有写入权限，或被杀毒软件拦住了；\n"
                "  3. 文件被设成了「只读」——右键 → 属性 → 取消只读。\n\n"
                "副本位置：%s" % (exc, first_error, self.copy_path))
            return False

    def _show_status(self, count):
        parts = ["宿委会查寝总分计算器 %s（寝室号匹配方式：%s）"
                 % (VERSION, SOURCE_TITLES.get(self.source, self.source)),
                 "原件：%s" % self.src_path,
                 "副本：%s" % self.copy_path]
        if self.source == SOURCE_ROOM:
            parts.append("已从 A 列解析出 %d 个可匹配的寝室号（测试中，请核对）。" % count)
        else:
            parts.append("已载入 %d 个寝室号。" % count)
        if self.formula_rows_cleaned:
            parts.append("等级列原有的 %d 行公式已由程序接管。" % self.formula_rows_cleaned)
        else:
            parts.append("等级由程序按分数生成（<60 不合格，60~79 合格，80~89 良好，≥90 优秀）。")
        self.status_var.set("\n".join(parts))

    # ------------------------------------------------------------ 交互逻辑
    def _focus_score_entry(self):
        """把焦点交给第一个扣分输入框。"""
        self._unlock_entries()
        self.entries[0][1].focus_set()

    def _lock_entries(self):
        self.locked = True
        for _, ent in self.entries:
            ent.configure(state="disabled")

    def _unlock_entries(self):
        self.locked = False
        for _, ent in self.entries:
            ent.configure(state="normal")

    def _lookup(self, text):
        """按当前匹配方式找行号，返回 (行号, 显示名)；找不到返回 (None, None)。"""
        row = self.row_by_tag.get(text)
        if row is not None:
            if self.source == SOURCE_ROOM and row in self.room_rows:
                return row, self.room_rows[row][0]
            return row, text
        if self.source == SOURCE_ROOM:
            # 允许“19栋201 / 19-201 / 201”几种写法互相匹配
            for key in room_keys(text):
                row = self.row_by_tag.get(key)
                if row is not None:
                    return row, self.room_rows.get(row, (key, []))[0]
        return None, None

    def on_room_enter(self, _event=None):
        """寝室号框按回车：结算后的确认、或者开始录入一个新寝室。"""
        text = self.room_var.get().strip()

        # 结算完成后等待确认：-1 退出，空回车继续，直接输入新寝室号也可以
        if self.awaiting_enter:
            if text == END_CODE:
                self.on_quit()
                return
            if text == "":
                self.next_room()
                return
        elif self.locked:
            return

        if text == END_CODE:
            self.on_quit()
            return
        if not text:
            messagebox.showwarning("输入有误", "请先输入寝室号。")
            return

        row, shown = self._lookup(text)
        if row is None:
            messagebox.showwarning(
                "未找到寝室",
                "按当前方式（%s）没有找到 %s。\n请核对输入，或检查所选原件是否正确。"
                % (SOURCE_TITLES.get(self.source, self.source), text))
            self.room_entry.select_range(0, "end")
            return

        self.awaiting_enter = False
        self.current_tag = text
        self.current_row = row
        hint = "当前寝室：%s" % shown if self.source == SOURCE_ROOM else "当前寝室：%s" % text
        self.room_hint.configure(text=hint)
        self.score_var.set("本次得分：--")
        self.grade_var.set("")
        for var, _ in self.entries:
            var.set("")
        self.items = []
        self.last_committed_row = None
        self._bind_undo_keys()
        self._focus_score_entry()

    def _validate_item(self, proposed, previous, inserted, idx, var_key):
        """扣分框只允许敲入 0~999 的整数；第 1 格还允许 w / s（无人 / 锁）。"""
        # 注意：这个函数是 Tk 通过 validatecommand 调用的，所有参数都是**字符串**，
        # 第 1 格传进来的 idx 是 "0" 而不是 0。必须先转成整数再比较，
        # 否则 w / s 会被当成“非数字”拦掉（v2.5 修过一次）。
        try:
            cell = int(idx)
        except (TypeError, ValueError):
            cell = -1
        if self.locked:
            # 结算/等待确认期间：只允许清空（换寝室时清空），不接受新输入
            return proposed == ""
        if proposed == "":
            return True
        if cell == 0 and proposed.lower() in (MARK_NONE_KEY, MARK_LOCK_KEY):
            return True
        return re.fullmatch(r"\d{0,3}", proposed) is not None

    def _bind_undo_keys(self):
        """把「Ctrl + 退格」绑定到主窗口，用于撤销上一格（每次录入只绑一次）。"""
        if not self.undo_bound:
            self.root.bind("<Control-BackSpace>", self.on_undo_key)
            self.undo_bound = True

    def on_undo_key(self, event=None):
        """
        Ctrl + 退格：把最近一次按回车记录的扣分撤回来重填。

        为什么从普通退格改成 Ctrl + 退格（v2.5）：
          以前普通退格同时绑在「每个扣分格」和「主窗口」两处，而输入框自身的删字符
          是在窗口级绑定之前执行的——所以在当前格里按一下退格，字符删掉之后
          窗口级那层又触发了一次撤销，连上一格也被撤了，改当前格特别难受。
          现在普通退格只做「删当前格的字」，撤销统一走 Ctrl + 退格。
        """
        if self.locked:
            # 等确认/结算状态下不撤销（此时已写入副本，避免与已保存的数据不一致）
            return "break"
        self.undo_last_item()
        return "break"

    def undo_last_item(self):
        """撤销最近按回车记录的那一格：把值放回该格、全选，方便直接重敲。"""
        if not self.items:
            self.status_var.set("还没有记录任何扣分，没有可撤销的格子。")
            return
        idx = len(self.items) - 1            # 撤销最近记录的那一格
        value = self.items.pop()
        var = self.entries[idx][0]
        ent = self.entries[idx][1]
        if value == 0:
            var.set("")                      # 那个 0 是“结束标志”，撤掉后不留内容
        else:
            var.set(str(value))              # 把原来的数字放回去，方便确认后修改
        ent.configure(state="normal")
        ent.focus_set()
        ent.select_range(0, "end")           # 全选，直接重敲即可覆盖
        self.status_var.set("已撤销第 %d 格（原值 %s），可以重新填写，再按回车继续。"
                            % (idx + 1, value))

    def on_item_enter(self, idx):
        """在第 idx 个扣分框按回车：记录这一格的数字，然后跳到下一格。

        填 0 表示本次录入到此为止；第 11 格回车则自动结算。
        """
        if self.locked:
            return
        if idx != len(self.items):
            # 前面的格子还没填完就被跳过了
            messagebox.showwarning(
                "顺序不对",
                "请从第 %d 格开始按顺序填。\n没有扣分的格子可以填 0 表示到此为止。"
                % (len(self.items) + 1))
            if len(self.items) < MAX_ITEMS:
                self.entries[len(self.items)][1].focus_set()
            return
        raw = self.entries[idx][0].get().strip()
        if raw == "":
            messagebox.showwarning("还没填完", "第 %d 格还没有填数字。\n"
                                                "这一项没扣分就填 0（表示到此为止）。" % (idx + 1))
            return
        mark = special_mark_of(raw)
        if mark is not None:
            if idx != 0:
                messagebox.showwarning("填错位置",
                                       "w（无人）/ s（锁）只能填在第 1 格。")
                return
            self.finish_special(mark)
            return
        self._commit_value(int(raw))

    def on_item_focus_out(self, idx):
        """
        鼠标点到别处时，把这格填好的数字也记下来，等价于按了一次回车（v2.5）。

        这样不习惯按回车的人可以「填一格 → 点下一格」地录。
        空格子、或者不是“该填的那一格”离开时，什么都不做，也不弹警告。
        """
        if self._committing or self.locked:
            return
        if idx != len(self.items):
            return
        if self.entries[idx][0].get().strip() == "":
            return
        # after_idle：等这次点击的落点先定下来，免得和用户的点击抢焦点
        self.root.after_idle(lambda: self._commit_on_leave(idx))

    def _commit_on_leave(self, idx):
        """失去焦点后真正记录这一格（走和按回车完全相同的一套逻辑）。"""
        if self._committing or self.locked:
            return
        if idx != len(self.items):
            return
        if self.entries[idx][0].get().strip() == "":
            return
        self._committing = True
        try:
            self.on_item_enter(idx)
        finally:
            self._committing = False
        # 万一点到了“还没轮到”的格子，把焦点送回该填的那一格
        self._redirect_focus(idx + 1)

    def _redirect_focus(self, expect_idx):
        """鼠标点歪了就把焦点送回该填的格子；只改状态栏，不弹窗。"""
        if self.locked or expect_idx >= MAX_ITEMS:
            return
        try:
            current = self.root.focus_get()
        except Exception:
            return
        if current is None:
            return
        for i, (_, ent) in enumerate(self.entries):
            if ent is current:
                if i != expect_idx:
                    self.entries[expect_idx][1].focus_set()
                    self.status_var.set("请从第 %d 格开始按顺序填"
                                        "（刚才那格已经记下了）。" % (expect_idx + 1))
                return

    def _commit_value(self, value):
        """记录一项扣分并累计总分；值为 0 表示本次到此为止，满 11 项自动结算。"""
        self.items.append(value)
        next_idx = len(self.items)
        if value == 0:
            self.finish_round("结束标志 0")
        elif next_idx >= MAX_ITEMS:
            self.finish_round("已填满 %d 项" % MAX_ITEMS)
        elif next_idx < MAX_ITEMS:
            self.entries[next_idx][1].focus_set()

    # ------------------------------------------------------------ 结算
    def _settle_lock(self):
        """结算后的公共动作：锁住扣分格、断开退格撤销、进入“等待按回车”。"""
        self._lock_entries()
        self.last_committed_row = None      # 本轮已写入，不再允许撤销
        if self.undo_bound:
            try:
                self.root.unbind("<Control-BackSpace>")
            except Exception:
                pass
            self.undo_bound = False
        self.awaiting_enter = True

    def _wait_for_enter(self, tip):
        """收尾：提示按回车进入下一个寝室（或自动继续）。"""
        if WAIT_ENTER:
            self.status_var.set(self.status_var.get() + "\n" + tip)
            self.room_var.set("")
            self.room_hint.configure(text="按回车继续")
            self.room_entry.focus_set()
        else:
            self.pending_after = self.root.after(PAUSE_MS, self.next_room)

    def finish_round(self, reason):
        """结算：写分数与等级，进入“等待确认”状态，由用户按回车进入下一个寝室。"""
        self._settle_lock()
        score = FULL_SCORE - sum(self.items)
        filled = len([v for v in self.items if v != 0])
        grade = grade_of(score)
        self.score_var.set("本次得分：%d" % score)
        self.grade_var.set("等级：%s" % grade)
        self.status_var.set("录入 %d 个扣分（%s），得分 %d，等级 %s。"
                            % (filled, reason, score, grade))

        self.apply_score(score, grade)
        self._wait_for_enter("请核对上面显示的分数与等级，按回车进入下一个寝室。")

    def finish_special(self, mark):
        """第 1 格填了 w / s：把「无人」/「锁」写进分数列，等级留空，不参与算分。"""
        self._settle_lock()
        self.score_var.set("本次得分：—")
        self.grade_var.set("等级：留空")
        self.status_var.set("本寝室记为「%s」，分数列写它、等级列留空。" % mark)
        self.apply_mark(mark)
        self._wait_for_enter("按回车进入下一个寝室。")

    def apply_score(self, score, grade):
        """把分数写进 D 列、等级写进 E 列；已有分数时先确认是否覆盖。"""
        tag = self.current_tag
        row = getattr(self, "current_row", None)
        if row is None:
            row = self.row_by_tag.get(tag)
        if row is None:
            messagebox.showerror("内部错误", "找不到 %s 对应的行，本次结果未写入。" % tag)
            return
        old = self.ws.cell(row, SCORE_COL).value
        if cell_has_value(old):
            old_grade = self.ws.cell(row, GRADE_COL).value
            ok = messagebox.askyesno(
                "覆盖确认",
                "该寝室已有分数，是否覆盖？\n\n是（Y）或否（N）\n\n"
                "当前：%s 分（等级 %s）\n本次：%d 分（等级 %s）"
                % (old, old_grade if cell_has_value(old_grade) else "—", score, grade),
                default=messagebox.NO)
            if not ok:
                self.status_var.set("%s 保留原有分数 %s，本次结果已丢弃。" % (tag, old))
                return
        self.ws.cell(row, SCORE_COL).value = score
        self.ws.cell(row, GRADE_COL).value = grade      # 等级由程序生成
        self.last_committed_row = row
        if self._save_copy():
            self.status_var.set("已写入：%s（第 %d 行）分数 %d、等级 %s，副本已保存。"
                                "（按错可退回上一格重填）" % (tag, row, score, grade))

    def apply_mark(self, mark):
        """把「无人」/「锁」写进 D 列、E 列清空；该行已有分数时先确认是否覆盖。"""
        tag = self.current_tag
        row = getattr(self, "current_row", None)
        if row is None:
            row = self.row_by_tag.get(tag)
        if row is None:
            messagebox.showerror("内部错误", "找不到 %s 对应的行，本次结果未写入。" % tag)
            return
        old = self.ws.cell(row, SCORE_COL).value
        if cell_has_value(old):
            old_grade = self.ws.cell(row, GRADE_COL).value
            ok = messagebox.askyesno(
                "覆盖确认",
                "该寝室已有分数，是否覆盖？\n\n是（Y）或否（N）\n\n"
                "当前：%s（等级 %s）\n本次：%s（等级留空）"
                % (old, old_grade if cell_has_value(old_grade) else "—", mark),
                default=messagebox.NO)
            if not ok:
                self.status_var.set("%s 保留原有内容 %s，本次结果已丢弃。" % (tag, old))
                return
        self.ws.cell(row, SCORE_COL).value = mark
        self.ws.cell(row, GRADE_COL).value = None       # 等级留空
        self.last_committed_row = row
        if self._save_copy():
            self.status_var.set("已写入：%s（第 %d 行）分数列填「%s」，等级留空。"
                                % (tag, row, mark))

    def undo_score(self):
        """撤销刚写入副本的分数：把这一行的分数和等级清回空白。"""
        row = self.last_committed_row
        if row is None:
            return
        self.ws.cell(row, SCORE_COL).value = None
        self.ws.cell(row, GRADE_COL).value = None
        self.last_committed_row = None
        if self._save_copy():
            self.status_var.set("已撤销刚写入的分数（第 %d 行的分数与等级已清回空白），"
                                "请重新填写扣分。注意：覆盖前该行原有的分数不会自动恢复。"
                                % row)

    def next_room(self):
        """回到寝室号输入，准备量下一个寝室。"""
        self.pending_after = None
        self.awaiting_enter = False
        self.room_var.set("")
        self.room_hint.configure(text="")
        self.score_var.set("本次得分：--")
        self.grade_var.set("")
        for var, _ in self.entries:
            var.set("")
        self.items = []
        self.current_tag = None
        self.current_row = None
        self._unlock_entries()
        for _, ent in self.entries:
            ent.configure(state="disabled")
        count = len(self.row_by_tag)
        self._show_status(count)
        self.room_entry.focus_set()

    # ------------------------------------------------------------ 退出
    def on_quit(self):
        if self.pending_after is not None:
            self.root.after_cancel(self.pending_after)
            self.pending_after = None
        try:
            if self.wb is not None:
                self.wb.save(self.copy_path)
        except Exception as exc:
            messagebox.showerror("保存失败", "退出前保存副本失败：\n%s" % exc)
            return
        messagebox.showinfo("已结束",
                            "程序结束。\n\n最终副本已生成：\n%s" % getattr(self, "copy_path", ""))
        self.root.destroy()


# ---------------------------------------------------------------- 入口
def show_disclaimer(root):
    """每次启动都要先显示免责声明：放在屏幕正中，必须按回车或点按钮才能继续。"""
    dialog = tk.Toplevel(root)
    dialog.title(DISCLAIMER_TITLE)
    dialog.resizable(False, False)

    frame = tk.Frame(dialog, padx=18, pady=14)
    frame.pack(fill="both", expand=True)

    tk.Label(frame, text=DISCLAIMER_TEXT, font=("Microsoft YaHei UI", 11),
             fg="#cc0000", justify="left", anchor="w").pack(fill="both", expand=True)

    button = tk.Button(frame, text="我已知晓（回车）", font=("Microsoft YaHei UI", 11),
                       width=18, height=1, command=dialog.destroy)
    button.pack(pady=(14, 0))

    def close(_event=None):
        try:
            dialog.destroy()
        except Exception:
            pass

    dialog.bind("<Return>", close)
    dialog.bind("<space>", close)
    dialog.bind("<Escape>", close)
    dialog.protocol("WM_DELETE_WINDOW", close)

    dialog.update_idletasks()
    w, h = dialog.winfo_reqwidth(), dialog.winfo_reqheight()
    x = max((dialog.winfo_screenwidth() - w) // 2, 0)
    y = max((dialog.winfo_screenheight() - h) // 2, 0)
    dialog.geometry("%dx%d+%d+%d" % (w, h, x, y))
    dialog.lift()
    dialog.focus_force()
    try:
        dialog.attributes("-topmost", True)
    except Exception:
        pass
    button.focus_set()

    dialog.grab_set()
    dialog.wait_window()
    try:
        dialog.attributes("-topmost", False)
    except Exception:
        pass


def choose_source(root):
    """
    让用户选择寝室号的匹配方式，返回 SOURCE_TAG / SOURCE_ROOM / None(取消)。
    两个按钮都写成完整文字，默认（回车）使用推荐的 F 列 tag。
    """
    dialog = tk.Toplevel(root)
    dialog.title("选择寝室号匹配方式")
    dialog.resizable(False, False)

    result = {"value": None}

    def pick(value):
        result["value"] = value
        dialog.destroy()

    frame = tk.Frame(dialog, padx=20, pady=16)
    frame.pack(fill="both", expand=True)

    tk.Label(frame, text="请选择怎么在总表里找寝室号：", font=FONT_UI,
             justify="left").pack(anchor="w")
    tk.Label(frame, text="使用 F 列 tag：靠表里已有的 tag 列匹配，最稳。\n"
                         "由程序读取寝室号：从 A 列“寝室号”里解析楼栋号+房号，不依赖 tag。",
             font=("Microsoft YaHei UI", 9), fg="#666", justify="left").pack(
        anchor="w", pady=(6, 14))

    btn_tag = tk.Button(frame, text="使用F列tag（推荐）", font=FONT_UI, width=26, height=2,
                        command=lambda: pick(SOURCE_TAG))
    btn_tag.pack(fill="x", pady=4)

    btn_room = tk.Button(frame, text="由程序读取寝室号（测试中）", font=FONT_UI,
                         width=26, height=2, command=lambda: pick(SOURCE_ROOM))
    btn_room.pack(fill="x", pady=4)

    def close(_event=None):
        pick(result["value"])       # 直接关窗视为取消

    dialog.bind("<Return>", lambda e: pick(SOURCE_TAG))   # 回车默认选推荐项
    dialog.bind("<Escape>", close)
    dialog.protocol("WM_DELETE_WINDOW", close)

    dialog.update_idletasks()
    w, h = dialog.winfo_reqwidth(), dialog.winfo_reqheight()
    x = max((dialog.winfo_screenwidth() - w) // 2, 0)
    y = max((dialog.winfo_screenheight() - h) // 2, 0)
    dialog.geometry("%dx%d+%d+%d" % (w, h, x, y))
    dialog.lift()
    dialog.focus_force()
    try:
        dialog.attributes("-topmost", True)
    except Exception:
        pass
    btn_tag.focus_set()

    dialog.grab_set()
    dialog.wait_window()
    try:
        dialog.attributes("-topmost", False)
    except Exception:
        pass
    return result["value"]


def choose_file():
    """弹出文件选择框，默认从程序所在目录开始找。"""
    start = os.path.dirname(os.path.abspath(sys.executable if getattr(sys, "frozen", False)
                                            else __file__))
    return filedialog.askopenfilename(
        title="请选择查寝总表原件（Excel）",
        initialdir=start,
        filetypes=[("Excel 工作簿", "*.xlsx"), ("所有文件", "*.*")])


def write_crash_log(exc_text):
    """把启动异常写到 exe 同目录的 出错日志.txt，避免窗口版程序“无声消失”。"""
    try:
        if getattr(sys, "frozen", False):
            folder = os.path.dirname(os.path.abspath(sys.executable))
        else:
            folder = os.path.dirname(os.path.abspath(__file__))
        path = os.path.join(folder, "出错日志.txt")
        with open(path, "a", encoding="utf-8") as f:
            f.write("=" * 60 + "\n")
            f.write("时间：%s\n" % datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
            f.write("程序：查寝总分计算器 %s\n\n" % VERSION)
            f.write(exc_text + "\n")
    except Exception:
        pass


def log_prepare_failure(stage, src_path, source, exc):
    """
    把一个"文件准备失败"的真实原因写进 出错日志.txt。

    为什么需要：以前只有程序崩溃才写日志，而且写的往往是崩溃时的二次异常
    （比如"窗口已经没了"），真正的原因（为什么读不了表）反而没留下。
    有了这一条，别人发日志过来就能直接看出卡在哪一步、报的是什么错。
    """
    try:
        import traceback
        lines = [
            "【文件准备失败】",
            "阶段：%s" % stage,
            "匹配方式：%s" % SOURCE_TITLES.get(source, source),
            "所选原件：%s" % src_path,
            "程序目录：%s" % app_dir(),
        ]
        if exc is not None:
            lines.append("异常类型：%s" % type(exc).__name__)
            lines.append("异常内容：%s" % exc)
            lines.append("")
            lines.append("".join(
                traceback.format_exception(type(exc), exc, exc.__traceback__)).rstrip())
        write_crash_log("\n".join(lines))
    except Exception:
        pass


def main():
    src_path = sys.argv[1] if len(sys.argv) > 1 else None
    if src_path and not os.path.exists(src_path):
        src_path = None
    # 可选：第二个参数指定匹配方式（tag / room），方便命令行或快捷方式直接指定
    cli_source = None
    if len(sys.argv) > 2 and sys.argv[2] in (SOURCE_TAG, SOURCE_ROOM):
        cli_source = sys.argv[2]
    source = cli_source or SOURCE_TAG

    root = tk.Tk()
    root.withdraw()          # 先不显示，等声明确认后再露面

    try:
        # 1) 免责声明
        show_disclaimer(root)

        # 2) 选匹配方式 → 3) 选原件 → 4) 准备文件（三步可能来回绕）
        #    如果选的方式在表里找不到寝室号（比如 tag 列为空），就回到第 2 步重选，
        #    刚才选好的文件沿用，不用重新挑。
        first_round = True
        while True:
            if first_round and cli_source:
                source = cli_source          # 命令行指定了就用它，跳过选择窗
            else:
                chosen = choose_source(root)
                if chosen is None:
                    root.destroy()
                    return
                source = chosen

            if not src_path:
                src_path = choose_file()
                if not src_path:
                    root.destroy()
                    return
            first_round = False

            application = ScoreApp(root, src_path, source)
            if application.ready:
                break
            if application.failed_stage == "no_room":
                continue                     # 换个匹配方式再来
            root.destroy()                   # 其它失败已经提示过了，直接收工
            return

        # 5) 显示主界面并抬到最前（此时窗口和界面都还没建，不会白闪）
        root.deiconify()
        # 窗口空白处按 Ctrl + 退格也能撤销上一格（与扣分框里的等效）
        root.bind("<Control-BackSpace>", application.on_undo_key)
        root.lift()
        root.focus_force()
        try:
            root.attributes("-topmost", True)
            root.after(400, lambda: root.attributes("-topmost", False))
        except Exception:
            pass
        root.mainloop()
    except Exception:
        import traceback
        detail = traceback.format_exc()
        write_crash_log(detail)
        try:
            messagebox.showerror(
                "程序出错",
                "程序启动/运行中出错，已在程序所在文件夹写了一份“出错日志.txt”，"
                "请把它发给作者（QQ %s）反馈。\n\n%s" % (AUTHOR_QQ, detail[-800:]))
        except Exception:
            pass
        raise


if __name__ == "__main__":
    main()
