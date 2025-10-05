import tkinter as tk
from tkinter import ttk
import tkinter.font as tkfont
import json
import os
import re
import math

# 状态文件路径
STATE_FILE = "state.json"

class VariableCalculator:
    def __init__(self, root):
        """
        初始化变量计算器
        :param root: tkinter根窗口
        """
        self.root = root
        self.root.title("VarCalc")
        self.font_family = "Consolas"
        self.font_size = 12
        self.default_font = (self.font_family, self.font_size)

        # ================= 自定义可滚动 Sheets 栏 =================
        # 创建用于显示sheet标签的画布
        self.sheet_canvas = tk.Canvas(root, height=30, bg="lightgray", highlightthickness=0)
        self.sheet_canvas.pack(fill="x", side="top")
        # 创建包含sheet按钮的框架
        self.sheet_frame = tk.Frame(self.sheet_canvas, bg="lightgray")
        self.sheet_window = self.sheet_canvas.create_window((0,0), window=self.sheet_frame, anchor="nw")
        # 绑定事件处理函数
        self.sheet_canvas.bind("<Configure>", self.update_sheet_scrollregion)
        self.sheet_canvas.bind_all("<MouseWheel>", self.scroll_sheets)

        # ================= Notebook =================
        # 创建主内容区域的notebook控件
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill=tk.BOTH, expand=True)
        # 隐藏默认的tab栏
        style = ttk.Style()
        style.layout("TNotebook.Tab", [])  # 隐藏默认 tab 栏

        # 存储sheet按钮的列表
        self.sheets = []  # 保存 Sheet 按钮
        # 添加新sheet的按钮
        self.add_sheet_button = tk.Button(self.sheet_frame, text="+", command=self.add_tab)
        self.add_sheet_button.pack(side="left", padx=2, pady=2)

        # 状态文件恢复
        # 绑定窗口关闭事件
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        # 如果存在状态文件，则加载状态，否则创建新标签页
        if os.path.exists(STATE_FILE):
            try:
                self.load_state()
            except Exception:
                self.add_tab()
        else:
            self.add_tab()

    # ================= Sheet 滚动逻辑 =================
    def update_sheet_scrollregion(self, event=None):
        """
        更新sheet栏的滚动区域
        """
        self.sheet_canvas.configure(scrollregion=self.sheet_canvas.bbox("all"))

    def scroll_sheets(self, event):
        """
        处理鼠标滚轮事件，滚动sheet栏
        :param event: 鼠标事件
        """
        if event.delta:
            # Windows系统滚动
            self.sheet_canvas.xview_scroll(-1*(event.delta//120), "units")
        else:
            # Linux系统滚动
            if event.num == 4:
                self.sheet_canvas.xview_scroll(-1, "units")
            elif event.num == 5:
                self.sheet_canvas.xview_scroll(1, "units")

    # ================= 添加新 Sheet =================
    def add_tab(self, title=None, contents=None):
        """
        添加新的标签页
        :param title: 标签页标题
        :param contents: 标签页内容
        """
        # 创建新的标签页框架
        tab = tk.Frame(self.notebook, bg="white")
        tab_name = title if title else f"Sheet{len(self.notebook.tabs())+1}"
        self.notebook.add(tab, text=tab_name)

        # 创建sheet栏按钮
        btn = tk.Button(self.sheet_frame, text=tab_name, relief="raised")
        btn.pack(side="left", padx=2, pady=2)
        # 绑定点击事件选择标签页
        btn.bind("<Button-1>", lambda e, t=tab: self.notebook.select(t))
        # 绑定双击事件重命名标签页
        btn.bind("<Double-1>", lambda e, b=btn: self.rename_sheet(b))
        # 绑定右键事件删除标签页
        btn.bind("<Button-3>", lambda e, b=btn, t=tab: self.delete_sheet(b, t))  # 右键删除

        self.sheets.append(btn)
        self.update_sheet_scrollregion()

        # ====== Tab 内容初始化 ======
        # 初始化标签页变量和条目
        tab.variables = {}
        tab.entries = []
        # 创建内容画布
        tab.canvas = tk.Canvas(tab, bg="white", highlightthickness=0)
        tab.canvas.pack(fill=tk.BOTH, expand=True)
        # 创建内容框架
        tab.frame = tk.Frame(tab.canvas, bg="white")
        tab.canvas.create_window((0,0), window=tab.frame, anchor="nw")
        # 绑定画布配置事件
        tab.canvas.bind("<Configure>", lambda e, t=tab: self.on_canvas_configure(t))
        # 添加新行按钮
        tab.add_button = tk.Button(tab.frame, text="+", command=lambda t=tab: self.add_input_row(t))
        tab.add_button.pack(pady=5)

        # 添加内容行
        if contents:
            for expr in contents:
                self.add_input_row(tab, expr)
        else:
            self.add_input_row(tab)

    def delete_sheet(self, button, tab):
        """
        删除标签页
        :param button: 标签页按钮
        :param tab: 标签页框架
        """
        # 至少保留一个标签页
        if len(self.sheets) <= 1:
            return
        idx = self.sheets.index(button)
        self.notebook.forget(tab)
        button.destroy()
        self.sheets.pop(idx)
        # 如果还有标签页，选择最后一个
        if self.notebook.tabs():
            self.notebook.select(self.notebook.tabs()[-1])

    def rename_sheet(self, button):
        """
        重命名标签页
        :param button: 标签页按钮
        """
        old_name = button['text']
        # 创建输入框用于重命名
        entry = tk.Entry(self.sheet_frame)
        entry.insert(0, old_name)
        entry.select_range(0, tk.END)
        entry.focus()
        # 定位输入框位置和大小
        entry.place(x=button.winfo_x(), y=button.winfo_y(), width=button.winfo_width(), height=button.winfo_height())

        def save_name(event=None):
            """
            保存新的标签页名称
            """
            new_name = entry.get().strip() or old_name
            idx = self.sheets.index(button)
            self.notebook.tab(idx, text=new_name)
            button.config(text=new_name)
            entry.destroy()

        # 绑定回车和失去焦点事件保存名称
        entry.bind("<Return>", save_name)
        entry.bind("<FocusOut>", save_name)

    # ================= 输入框逻辑 =================
    def add_input_row(self, tab, initial_text=""):
        """
        添加输入行
        :param tab: 标签页框架
        :param initial_text: 初始文本
        """
        # 创建行框架
        row_frame = tk.Frame(tab.frame, bg="white")
        row_frame.pack(fill="x", pady=3, anchor="w")

        # 创建文本输入框
        text = tk.Text(row_frame, height=1, wrap="word", font=self.default_font,
                       bd=1, relief="solid", padx=2, pady=2)
        text.insert("1.0", initial_text)
        text.pack(side="left", padx=(6,4))
        # 配置注释样式
        text.tag_configure("comment", foreground="green")

        # 创建结果显示标签
        result_label = tk.Label(row_frame, text="= ?", bg="white", fg="blue", font=self.default_font, anchor="w")
        result_label.pack(side="left", padx=(4,10))

        # 添加双击复制功能
        result_label.bind("<Double-Button-1>", lambda e, rl=result_label: self.copy_result(rl))

        # 将输入框、结果标签和行框架添加到标签页条目列表
        tab.entries.append((text, result_label, row_frame))

        # 重新放置添加按钮
        tab.add_button.pack_forget()
        tab.add_button.pack(pady=5)

        # ====== 新增：自定义粘贴逻辑，去掉 Excel 自带换行 ======
        def clean_paste(event, t=tab, tw=text, rl=result_label):
            """
            自定义粘贴逻辑，去除Excel自带换行和多余空格
            """
            try:
                pasted = tw.selection_get(selection='CLIPBOARD')
            except tk.TclError:
                return "break"  # 没有剪贴板内容

            # 去掉 Excel 自带的结尾换行
            pasted = pasted.rstrip("\r\n")

            # 去除前后空格
            pasted = pasted.strip()

            # 如果有选中内容，先删除
            if tw.tag_ranges("sel"):
                tw.delete("sel.first", "sel.last")

            # 插入清理过的文本
            tw.insert("insert", pasted)

            # 更新显示
            self.adjust_row_size(t, tw, rl)
            self.update_all(t)

            return "break"  # 阻止默认粘贴


        # 绑定粘贴事件
        text.bind("<<Paste>>", clean_paste)

        def on_double_click(event, text_widget):
            """
            处理双击事件，智能选择文本
            :param event: 鼠标事件
            :param text_widget: 文本控件
            """
            # 获取点击索引
            index = text_widget.index(f"@{event.x},{event.y}")
            line, char = map(int, index.split('.'))
            # 获取当前行内容
            line_text = text_widget.get(f"{line}.0", f"{line}.end")
            # 计算点击位置在行文本中的偏移
            offset = char

            # 先尝试处理引号内内容的选择
            quote_left = -1
            quote_right = -1

            # 检查是否在双引号内
            if offset < len(line_text) and line_text[offset] == '"':
                # 点击位置就是引号，查找匹配的引号
                quote_left = offset
                # 从当前位置向后查找下一个引号
                quote_right = line_text.find('"', offset + 1)
            elif offset > 0 and offset <= len(line_text):
                # 检查是否在两个引号之间
                # 查找左边最近的引号
                quote_left = line_text.rfind('"', 0, offset)
                if quote_left != -1:
                    # 查找右边对应的引号
                    quote_right = line_text.find('"', quote_left + 1)
                    # 确保点击位置在两个引号之间
                    if quote_right != -1 and quote_left < offset < quote_right:
                        pass  # 已经找到有效的引号对
                    else:
                        quote_left = -1
                        quote_right = -1

            # 如果找到有效的引号对，则选中引号内的内容
            if quote_left != -1 and quote_right != -1:
                text_widget.tag_remove("sel", f"{line}.0", f"{line}.end")
                # 选中引号之间的内容（不包括引号本身）
                text_widget.tag_add("sel", f"{line}.{quote_left + 1}", f"{line}.{quote_right}")
                return "break"  # 阻止默认双击行为

            # 否则执行原有的数字选择功能
            left = offset
            right = offset
            while left > 0 and (line_text[left - 1].isdigit() or line_text[left - 1] == "."):
                left -= 1
            while right < len(line_text) and (line_text[right].isdigit() or line_text[right] == "."):
                right += 1
            # 选中数字
            text_widget.tag_remove("sel", f"{line}.0", f"{line}.end")
            if left != right:
                text_widget.tag_add("sel", f"{line}.{left}", f"{line}.{right}")
            return "break"  # 阻止默认双击行为

        # 绑定双击事件
        text.bind("<Double-Button-1>", lambda e, tw=text: on_double_click(e, tw))

        # 绑定按键释放事件，调整大小并更新计算结果
        text.bind("<KeyRelease>", lambda e, t=tab: (self.adjust_row_size(t, text, result_label), self.update_all(t)))

        def on_backspace(event, t=tab, tw=text, rf=row_frame):
            """
            处理退格键事件，当文本框为空时删除行
            """
            content = tw.get("1.0", "end-1c")
            if content == "":
                idx = next((i for i, (tt, _, _) in enumerate(t.entries) if tt == tw), None)
                if idx is not None:
                    rf.destroy()
                    del t.entries[idx]
                    if idx-1 >= 0:
                        t.entries[idx-1][0].focus_set()
                    else:
                        if not t.entries:
                            self.add_input_row(t)
                return "break"
            return None

        # 绑定退格键事件
        text.bind("<KeyPress-BackSpace>", on_backspace)

        def on_return(event, t=tab, tw=text):
            """
            处理回车键事件，添加新行
            """
            idx = next((i for i, (tt, _, _) in enumerate(t.entries) if tt == tw), None)
            if idx is not None:
                self.add_input_row(t)
                t.entries[idx+1][0].focus_set()
            return "break"

        # 绑定回车键事件
        text.bind("<Return>", on_return)

        # 设置焦点并调整大小
        text.focus_set()
        self.root.update_idletasks()
        self.adjust_row_size(tab, text, result_label)
        self.update_all(tab)

    def copy_result(self, result_label):
        """
        双击复制结果到剪贴板
        :param result_label: 结果标签控件
        """
        # 获取标签文本并去除前缀"= "
        result_text = result_label.cget("text")
        if result_text.startswith("= "):
            result_text = result_text[2:]

        # 对于变量赋值格式 "= var_name → value"，只提取值部分
        if " → " in result_text:
            # 提取箭头后的值部分
            result_text = result_text.split(" → ", 1)[1]

        # 复制到剪贴板
        self.root.clipboard_clear()
        self.root.clipboard_append(result_text)
        self.root.update()


    def adjust_row_size(self, tab, text_widget, result_label):
        """
        调整输入行大小以适应内容
        :param tab: 标签页框架
        :param text_widget: 文本控件
        :param result_label: 结果标签
        """

        def pixel_to_chars(pixel, font):
            """
            把像素宽度转换为 Text.width 的字符数（基于 '0' 的像素宽做二分查找）
            :param pixel: 像素数
            :param font: 字体
            :return: 字符数
            """
            if pixel <= 0:
                return 1
            zero_w = font.measure("0") or 7
            # upper bound 初始估计
            hi = max(1, int(pixel / zero_w) + 3)
            # 保证 hi 足够大
            while font.measure("0" * hi) < pixel:
                hi *= 2
            lo = 1
            while lo < hi:
                mid = (lo + hi) // 2
                if font.measure("0" * mid) < pixel:
                    lo = mid + 1
                else:
                    hi = mid
            return max(1, lo)

        # 更新界面
        self.root.update_idletasks()
        font = tkfont.Font(font=text_widget['font'])

        # 获取文本内容
        content = text_widget.get("1.0", "end-1c")
        lines = content.split("\n") if content else [""]

        # === 临时关闭自动换行以测量逻辑行的完整像素宽 ===
        old_wrap = text_widget.cget("wrap")
        try:
            text_widget.config(wrap="none")
            self.root.update_idletasks()

            def measure_line_by_bbox(line_index, line_text):
                """
                尝试用 bbox 测量整行像素宽，失败则回退到逐字符累计
                :param line_index: 行索引
                :param line_text: 行文本
                :return: 像素宽度
                """
                if not line_text:
                    return font.measure("0")  # 空行设为一个最小宽度
                start_idx = f"{line_index}.0"
                end_idx = f"{line_index}.end-1c"
                try:
                    bbox_start = text_widget.bbox(start_idx)
                    bbox_end = text_widget.bbox(end_idx)
                    if bbox_start and bbox_end:
                        x_start = bbox_start[0]
                        x_end = bbox_end[0] + bbox_end[2]
                        # 有时 bbox 的 x 起点并非 0（存在内部 padding），直接取差值
                        return max(0, x_end - x_start)
                except Exception:
                    pass
                # 回退：逐字符测量（较慢，但稳妥）
                return sum(font.measure(ch) for ch in line_text)

            # 计算最长行的像素宽度
            longest_pixel = max((measure_line_by_bbox(i + 1, lines[i]) for i in range(len(lines))), default=0)
        finally:
            # 恢复原来的换行设置
            text_widget.config(wrap=old_wrap)
            self.root.update_idletasks()

        # 计算可用像素宽度 = 画布宽 - 结果标签宽 - padding
        canvas_w = tab.canvas.winfo_width()
        if canvas_w <= 1:
            canvas_w = max(self.root.winfo_width(), 300)
        result_req = result_label.winfo_reqwidth()
        padding = 30
        available_pixel = max(60, canvas_w - result_req - padding)

        # 给一点额外 margin 防止紧贴
        margin_pixel = 6
        needed_pixel = longest_pixel + margin_pixel

        # 根据内容长度决定是否换行
        if needed_pixel <= available_pixel:
            # 一行能完整显示，不需要换行
            width_chars = pixel_to_chars(needed_pixel, font)
            height_lines = max(1, len(lines))
        else:
            # 宽度受限，需要换行：以 available_pixel 为列宽换算列数，并估算需要的行数
            width_chars = pixel_to_chars(available_pixel - 2, font)  # 留一点余量
            # 需要的视觉行数 = ceil(最长行像素 / available_pixel)
            height_lines = max(1, math.ceil(longest_pixel / max(1, available_pixel)))

        # 设置文本控件的宽度和高度
        try:
            text_widget.config(width=width_chars, height=height_lines)
        except Exception:
            pass

    def on_canvas_configure(self, tab):
        """
        处理画布配置事件，更新滚动区域
        :param tab: 标签页框架
        """
        tab.canvas.configure(scrollregion=tab.canvas.bbox("all"))

    # ================= 计算逻辑 =================
    def update_all(self, tab):
        """
        更新所有计算结果
        :param tab: 标签页框架
        """
        # 清空变量字典
        tab.variables.clear()
        # 获取所有行的内容
        rows = [(text.get("1.0", "end-1c").strip(), text, label) for text, label, _ in tab.entries]
        # 定义注释模式
        comment_pattern = r'"[^"]*"'

        # 处理每一行
        for expr, text_widget, label in rows:
            # 清除旧的注释标记
            text_widget.tag_remove("comment", "1.0", "end")

            # 如果表达式为空，显示默认值
            if not expr:
                label.config(text="= ?")
                continue

            # 标记注释为绿色
            for m in re.finditer(comment_pattern, expr):
                start_idx = f"1.0+{m.start()}c"
                end_idx = f"1.0+{m.end()}c"
                text_widget.tag_add("comment", start_idx, end_idx)

            # 生成计算表达式：去掉注释
            calc_expr = re.sub(comment_pattern, "", expr)

            try:
                # 如果包含等号且不以问号结尾，则为变量赋值
                if "=" in calc_expr and not calc_expr.endswith("?") and not calc_expr.endswith("？"):
                    var, val = calc_expr.split("=",1)
                    var = var.strip()
                    val = val.strip()
                    calc_val = self.safe_eval(val, tab.variables)
                    if calc_val is not None:
                        tab.variables[var] = calc_val
                        label.config(text=f"= {var} → {calc_val}")
                    else:
                        label.config(text="= 错误")
                else:
                    # 否则为表达式计算
                    calc_expr = calc_expr.rstrip("?？").strip()
                    for var, val in tab.variables.items():
                        calc_expr = re.sub(r"\b"+re.escape(var)+r"\b", str(val), calc_expr)
                    calc_val = self.safe_eval(calc_expr, tab.variables)
                    if calc_val is not None:
                        label.config(text=f"= {calc_val}")
                    else:
                        label.config(text="= 错误")
            except Exception:
                label.config(text="= 错误")

        # 更新界面
        self.root.update_idletasks()
        for text_widget, label, _ in tab.entries:
            self.adjust_row_size(tab, text_widget, label)

    def safe_eval(self, expr, variables):
        """
        安全计算表达式
        :param expr: 表达式
        :param variables: 变量字典
        :return: 计算结果或None
        """
        try:
            expr = expr.strip()
            if expr == "":
                return None
            # 定义允许的字符模式
            allowed_re = re.compile(r'^[0-9+\-*/().\s]+$')
            if allowed_re.match(expr):
                # 只包含基本数学运算符的表达式
                return eval(expr, {"__builtins__": None}, {})
            else:
                # 包含变量的表达式
                return eval(expr, {"__builtins__": None}, variables)
        except Exception:
            return None

    # ================= 保存/恢复 =================
    def save_state(self):
        """
        保存程序状态到文件
        """
        # 构建状态数据结构
        state = {"window":{"geometry":self.root.winfo_geometry()}, "tabs":[]}
        for i, tab_id in enumerate(self.notebook.tabs()):
            tab_widget = self.notebook.nametowidget(tab_id)
            title = self.notebook.tab(tab_id, "text")
            contents = [text.get("1.0","end-1c") for text,_,_ in tab_widget.entries]
            state["tabs"].append({"title":title,"contents":contents})
        # 写入文件
        with open(STATE_FILE,"w",encoding="utf-8") as f:
            json.dump(state,f,ensure_ascii=False,indent=2)

    def load_state(self):
        """
        从文件加载程序状态
        """
        # 读取状态文件
        with open(STATE_FILE,"r",encoding="utf-8") as f:
            state = json.load(f)
        # 恢复窗口几何信息
        if "window" in state and "geometry" in state["window"]:
            try:
                self.root.geometry(state["window"]["geometry"])
            except Exception:
                pass
        # 恢复标签页
        for sheet in state.get("tabs",[]):
            self.add_tab(title=sheet["title"], contents=sheet["contents"])

    def on_close(self):
        """
        处理窗口关闭事件
        """
        try:
            self.save_state()
        except Exception:
            pass
        self.root.destroy()




if __name__ == "__main__":
    # 创建主窗口并运行程序
    root = tk.Tk()
    root.geometry("900x600")
    app = VariableCalculator(root)
    root.mainloop()
