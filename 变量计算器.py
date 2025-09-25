import tkinter as tk
from tkinter import ttk
import tkinter.font as tkfont
import json
import os
import re
import math

STATE_FILE = "state.json"

class VariableCalculator:
    def __init__(self, root):
        self.root = root
        self.root.title("变量计算器")
        self.font_family = "Consolas"
        self.font_size = 12
        self.default_font = (self.font_family, self.font_size)

        # ================= 自定义可滚动 Sheets 栏 =================
        self.sheet_canvas = tk.Canvas(root, height=30, bg="lightgray", highlightthickness=0)
        self.sheet_canvas.pack(fill="x", side="top")
        self.sheet_frame = tk.Frame(self.sheet_canvas, bg="lightgray")
        self.sheet_window = self.sheet_canvas.create_window((0,0), window=self.sheet_frame, anchor="nw")
        self.sheet_canvas.bind("<Configure>", self.update_sheet_scrollregion)
        self.sheet_canvas.bind_all("<MouseWheel>", self.scroll_sheets)

        # ================= Notebook =================
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill=tk.BOTH, expand=True)
        style = ttk.Style()
        style.layout("TNotebook.Tab", [])  # 隐藏默认 tab 栏

        self.sheets = []  # 保存 Sheet 按钮
        self.add_sheet_button = tk.Button(self.sheet_frame, text="+", command=self.add_tab)
        self.add_sheet_button.pack(side="left", padx=2, pady=2)

        # 状态文件恢复
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        if os.path.exists(STATE_FILE):
            try:
                self.load_state()
            except Exception:
                self.add_tab()
        else:
            self.add_tab()

    # ================= Sheet 滚动逻辑 =================
    def update_sheet_scrollregion(self, event=None):
        self.sheet_canvas.configure(scrollregion=self.sheet_canvas.bbox("all"))

    def scroll_sheets(self, event):
        if event.delta:
            self.sheet_canvas.xview_scroll(-1*(event.delta//120), "units")
        else:
            if event.num == 4:
                self.sheet_canvas.xview_scroll(-1, "units")
            elif event.num == 5:
                self.sheet_canvas.xview_scroll(1, "units")

    # ================= 添加新 Sheet =================
    def add_tab(self, title=None, contents=None):
        tab = tk.Frame(self.notebook, bg="white")
        tab_name = title if title else f"Sheet{len(self.notebook.tabs())+1}"
        self.notebook.add(tab, text=tab_name)

        # Sheet 栏按钮
        btn = tk.Button(self.sheet_frame, text=tab_name, relief="raised")
        btn.pack(side="left", padx=2, pady=2)
        btn.bind("<Button-1>", lambda e, t=tab: self.notebook.select(t))
        btn.bind("<Double-1>", lambda e, b=btn: self.rename_sheet(b))
        btn.bind("<Button-3>", lambda e, b=btn, t=tab: self.delete_sheet(b, t))  # 右键删除

        self.sheets.append(btn)
        self.update_sheet_scrollregion()

        # ====== Tab 内容初始化 ======
        tab.variables = {}
        tab.entries = []
        tab.canvas = tk.Canvas(tab, bg="white", highlightthickness=0)
        tab.canvas.pack(fill=tk.BOTH, expand=True)
        tab.frame = tk.Frame(tab.canvas, bg="white")
        tab.canvas.create_window((0,0), window=tab.frame, anchor="nw")
        tab.canvas.bind("<Configure>", lambda e, t=tab: self.on_canvas_configure(t))
        tab.add_button = tk.Button(tab.frame, text="+", command=lambda t=tab: self.add_input_row(t))
        tab.add_button.pack(pady=5)

        if contents:
            for expr in contents:
                self.add_input_row(tab, expr)
        else:
            self.add_input_row(tab)

    def delete_sheet(self, button, tab):
        if len(self.sheets) <= 1:
            return
        idx = self.sheets.index(button)
        self.notebook.forget(tab)
        button.destroy()
        self.sheets.pop(idx)
        if self.notebook.tabs():
            self.notebook.select(self.notebook.tabs()[-1])

    def rename_sheet(self, button):
        old_name = button['text']
        entry = tk.Entry(self.sheet_frame)
        entry.insert(0, old_name)
        entry.select_range(0, tk.END)
        entry.focus()
        entry.place(x=button.winfo_x(), y=button.winfo_y(), width=button.winfo_width(), height=button.winfo_height())
        def save_name(event=None):
            new_name = entry.get().strip() or old_name
            idx = self.sheets.index(button)
            self.notebook.tab(idx, text=new_name)
            button.config(text=new_name)
            entry.destroy()
        entry.bind("<Return>", save_name)
        entry.bind("<FocusOut>", save_name)

    # ================= 输入框逻辑 =================
    def add_input_row(self, tab, initial_text=""):
        row_frame = tk.Frame(tab.frame, bg="white")
        row_frame.pack(fill="x", pady=3, anchor="w")

        text = tk.Text(row_frame, height=1, wrap="word", font=self.default_font,
                       bd=1, relief="solid", padx=2, pady=2)
        text.insert("1.0", initial_text)
        text.pack(side="left", padx=(6,4))
        # 添加注释 tag
        text.tag_configure("comment", foreground="green")

        result_label = tk.Label(row_frame, text="= ?", bg="white", fg="blue", font=self.default_font, anchor="w")
        result_label.pack(side="left", padx=(4,10))

        tab.entries.append((text, result_label, row_frame))

        tab.add_button.pack_forget()
        tab.add_button.pack(pady=5)

        def on_double_click(event, text_widget):
            # 获取点击索引
            index = text_widget.index(f"@{event.x},{event.y}")
            line, char = map(int, index.split('.'))
            # 获取当前行内容
            line_text = text_widget.get(f"{line}.0", f"{line}.end")
            # 计算点击位置在行文本中的偏移
            offset = char
            # 左右扩展数字范围
            left = offset
            right = offset
            while left > 0 and line_text[left - 1].isdigit():
                left -= 1
            while right < len(line_text) and line_text[right].isdigit():
                right += 1
            # 选中数字
            text_widget.tag_remove("sel", f"{line}.0", f"{line}.end")
            if left != right:
                text_widget.tag_add("sel", f"{line}.{left}", f"{line}.{right}")
            return "break"  # 🔹关键，阻止默认双击行为

        text.bind("<Double-Button-1>", lambda e, tw=text: on_double_click(e, tw))

        text.bind("<KeyRelease>", lambda e, t=tab: (self.adjust_row_size(t, text, result_label), self.update_all(t)))
        text.bind("<<Paste>>", lambda e, t=tab: self.root.after(
            10, lambda: (self.adjust_row_size(t, text, result_label), self.update_all(t))))

        def on_backspace(event, t=tab, tw=text, rf=row_frame):
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

        text.bind("<KeyPress-BackSpace>", on_backspace)

        def on_return(event, t=tab, tw=text):
            idx = next((i for i, (tt, _, _) in enumerate(t.entries) if tt == tw), None)
            if idx is not None:
                self.add_input_row(t)
                t.entries[idx+1][0].focus_set()
            return "break"
        text.bind("<Return>", on_return)

        text.focus_set()
        self.root.update_idletasks()
        self.adjust_row_size(tab, text, result_label)
        self.update_all(tab)

    def adjust_row_size(self, tab, text_widget, result_label):
        self.root.update_idletasks()
        font = tkfont.Font(font=text_widget['font'])
        avg_char_w = font.measure('0') or 7
        content = text_widget.get("1.0", "end-1c")
        lines = content.split("\n") if content else [""]

        longest_pixel = max((font.measure(line) for line in lines), default=0)
        canvas_w = tab.canvas.winfo_width()
        if canvas_w <= 1:
            canvas_w = max(self.root.winfo_width(), 300)
        result_req = result_label.winfo_reqwidth()
        padding = 30
        available_pixel = max(60, canvas_w - result_req - padding)

        if longest_pixel <= available_pixel:
            max_chars = max((len(line) for line in lines))
            width_chars = max(1, max_chars + 1)
            height_lines = max(1, len(lines))
        else:
            width_chars = max(1, int(available_pixel / avg_char_w) - 1)
            height_lines = max(1, math.ceil(longest_pixel / available_pixel)) if available_pixel>0 else max(1,len(lines))

        try:
            text_widget.config(width=width_chars, height=height_lines)
        except Exception:
            pass

    def on_canvas_configure(self, tab):
        tab.canvas.configure(scrollregion=tab.canvas.bbox("all"))

    # ================= 计算逻辑 =================
    def update_all(self, tab):
        tab.variables.clear()
        rows = [(text.get("1.0", "end-1c").strip(), text, label) for text, label, _ in tab.entries]
        comment_pattern = r'"[^"]*"'

        for expr, text_widget, label in rows:
            # 清除旧 tag
            text_widget.tag_remove("comment", "1.0", "end")

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

        self.root.update_idletasks()
        for text_widget, label, _ in tab.entries:
            self.adjust_row_size(tab, text_widget, label)

    def safe_eval(self, expr, variables):
        try:
            expr = expr.strip()
            if expr == "":
                return None
            allowed_re = re.compile(r'^[0-9+\-*/().\s]+$')
            if allowed_re.match(expr):
                return eval(expr, {"__builtins__": None}, {})
            else:
                return eval(expr, {"__builtins__": None}, variables)
        except Exception:
            return None

    # ================= 保存/恢复 =================
    def save_state(self):
        state = {"window":{"geometry":self.root.winfo_geometry()}, "tabs":[]}
        for i, tab_id in enumerate(self.notebook.tabs()):
            tab_widget = self.notebook.nametowidget(tab_id)
            title = self.notebook.tab(tab_id, "text")
            contents = [text.get("1.0","end-1c") for text,_,_ in tab_widget.entries]
            state["tabs"].append({"title":title,"contents":contents})
        with open(STATE_FILE,"w",encoding="utf-8") as f:
            json.dump(state,f,ensure_ascii=False,indent=2)

    def load_state(self):
        with open(STATE_FILE,"r",encoding="utf-8") as f:
            state = json.load(f)
        if "window" in state and "geometry" in state["window"]:
            try:
                self.root.geometry(state["window"]["geometry"])
            except Exception:
                pass
        for sheet in state.get("tabs",[]):
            self.add_tab(title=sheet["title"], contents=sheet["contents"])

    def on_close(self):
        try:
            self.save_state()
        except Exception:
            pass
        self.root.destroy()




if __name__ == "__main__":
    root = tk.Tk()
    root.geometry("900x600")
    app = VariableCalculator(root)
    root.mainloop()
