import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import docx
import random
import re
import pickle # 用于保存和加载对象
import os     # 用于检查文件是否存在

# --- Question 类 (保持不变，确保它是可pickle的) ---
class Question:
    def __init__(self, q_type, original_num_text, text, options_text, answer_text, original_doc_order):
        self.q_type = q_type
        self.original_num_text = original_num_text 
        self.text = text 
        self.options_raw = options_text 
        self.answer_raw = answer_text 
        self.options = {} 
        self.answer = None 
        self.original_doc_order = original_doc_order
        self._parse_details()

    def _parse_details(self):
        match = re.match(r"^\s*(\d+[．.\s、]+)(.*)", self.original_num_text.strip())
        if match:
            self.original_num = match.group(1).strip()
            self.text = match.group(2).strip()
        else:
            self.original_num = "" 
            self.text = self.original_num_text.strip()

        if self.q_type in ["单选题", "多选题"]:
            for opt_line in self.options_raw:
                opt_match = re.match(r"^\s*([A-G])[\s.]+(.*)", opt_line.strip())
                if opt_match:
                    letter = opt_match.group(1)
                    opt_text = opt_match.group(2).strip()
                    self.options[letter] = opt_text
        
        cleaned_answer_text = self.answer_raw.replace("正确答案", "").strip("：: ").strip()
        if self.q_type == "单选题":
            self.answer = cleaned_answer_text[0] if cleaned_answer_text else None
        elif self.q_type == "多选题":
            self.answer = sorted([char for char in cleaned_answer_text.replace(" ", "") if char.isalpha()])
        elif self.q_type == "判断题":
            processed_ans_text = cleaned_answer_text.upper() # 转大写方便判断
            if "A" in processed_ans_text or "是" in processed_ans_text or "正确" in processed_ans_text:
                self.answer = "A"
            elif "B" in processed_ans_text or "否" in processed_ans_text or "错误" in processed_ans_text:
                self.answer = "B"
            else: # 降级处理
                self.answer = cleaned_answer_text[0] if cleaned_answer_text else None
        elif self.q_type == "填空题":
            parts = cleaned_answer_text.split()
            parsed_answers = []
            i = 0
            while i < len(parts):
                if parts[i].isdigit() and i + 1 < len(parts) and not parts[i+1].isdigit():
                    parsed_answers.append(parts[i+1])
                    i += 2
                else:
                    parsed_answers.append(parts[i])
                    i += 1
            self.answer = parsed_answers if parsed_answers else [cleaned_answer_text] 

    def get_display_text(self):
        return f"{self.original_num} {self.text}" if self.original_num else self.text

    def __repr__(self):
        return f"<{self.q_type} Q: {self.text[:20]}... A: {self.answer}>"

# --- QuizApp 类 ---
class QuizApp:
    SAVE_FILE_NAME = "quiz_progress.pkl" # 定义保存文件名

    # 定义统一的字体设置，方便修改
    QUESTION_FONT = ("微软雅黑", 18)
    OPTION_FONT = ("微软雅黑", 14)
    ANSWER_FONT = ("微软雅黑", 14)
    # QUESTION_FONT = ("Arial", 20)  # 题目正文和头部信息
    # OPTION_FONT = ("Arial", 20)    # 选项字体
    # ANSWER_FONT = ("Arial", 20)    # 显示答案的字体

    def __init__(self, master):
        self.master = master
        master.title("灵感菇")
        master.geometry("800x750") # 稍微调大一点高度给新按钮

        self.all_questions = []
        self.unanswered_questions = []
        self.answered_questions = []
        self.current_question_data = None
        self.user_answer_widgets = [] 
        self.last_imported_docx = None # 用于记录最后导入的docx路径，可选

        # --- Top Frame for File Import, Save/Load and Stats ---
        top_frame = tk.Frame(master, pady=10)
        top_frame.pack(fill=tk.X)

        self.btn_import = tk.Button(top_frame, text="导入新题库(Word)", command=self.import_word_file)
        self.btn_import.pack(side=tk.LEFT, padx=5)
        
        self.btn_save_progress = tk.Button(top_frame, text="保存进度", command=self.save_progress)
        self.btn_save_progress.pack(side=tk.LEFT, padx=5)

        # (加载按钮可选，因为我们会在启动时自动加载)
        # self.btn_load_progress = tk.Button(top_frame, text="加载进度", command=self.load_progress_manual)
        # self.btn_load_progress.pack(side=tk.LEFT, padx=5)

        self.stats_label = tk.Label(top_frame, text="未答题: 0 | 已答题: 0")
        self.stats_label.pack(side=tk.LEFT, padx=10)

        # --- Middle Frame (保持不变) ---
        self.question_frame = tk.LabelFrame(master, text="题目区域", padx=10, pady=10)
        self.question_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        self.question_header_label = tk.Label(self.question_frame, text="", justify=tk.LEFT, wraplength=750, font=self.QUESTION_FONT)
        self.question_header_label.pack(anchor="w")
        self.question_text_label = tk.Label(self.question_frame, text="请先导入题库或加载已有进度", justify=tk.LEFT, wraplength=750, font=self.QUESTION_FONT)
        self.question_text_label.pack(anchor="w", pady=5)
        self.options_frame = tk.Frame(self.question_frame) 
        self.options_frame.pack(anchor="w", fill=tk.X, pady=5)
        self.answer_display_label = tk.Label(self.question_frame, text="", justify=tk.LEFT, wraplength=750, fg="blue", font=self.ANSWER_FONT)
        self.answer_display_label.pack(anchor="w", pady=10)
        
        # --- Bottom Frame for Controls (保持不变) ---
        controls_frame = tk.Frame(master, pady=10)
        controls_frame.pack(fill=tk.X)
        self.btn_random_question = tk.Button(controls_frame, text="随机抽题", command=self.display_random_question, state=tk.DISABLED)
        self.btn_random_question.pack(side=tk.LEFT, padx=10)
        self.btn_show_answer = tk.Button(controls_frame, text="显示答案并移至已答", command=self.process_answer, state=tk.DISABLED)
        self.btn_show_answer.pack(side=tk.LEFT, padx=10)
        
        # --- Answered Questions Frame (保持不变) ---
        answered_frame = tk.LabelFrame(master, text="已答题目列表", padx=10, pady=10)
        answered_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        self.answered_listbox = tk.Listbox(answered_frame, selectmode=tk.EXTENDED, width=80, height=10)
        self.answered_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        answered_scrollbar = tk.Scrollbar(answered_frame, orient=tk.VERTICAL, command=self.answered_listbox.yview)
        answered_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.answered_listbox.config(yscrollcommand=answered_scrollbar.set)
        self.btn_move_back = tk.Button(answered_frame, text="移回未答列表", command=self.move_to_unanswered, state=tk.DISABLED)
        self.btn_move_back.pack(pady=5)

        # --- 自动加载进度 ---
        self.load_progress() 

        # --- 程序退出时自动保存 ---
        master.protocol("WM_DELETE_WINDOW", self.on_closing)

    def on_closing(self):
        if messagebox.askokcancel("退出", "确定要退出吗？将会自动保存当前进度。"):
            self.save_progress(silent=True) # 静默保存，不弹窗
            self.master.destroy()

    def save_progress(self, silent=False):
        if not self.all_questions: # 如果没有题目数据，不保存
            if not silent:
                messagebox.showinfo("保存", "没有题库数据可供保存。")
            return

        data_to_save = {
            "all_questions": self.all_questions,
            "unanswered_questions": self.unanswered_questions, # 直接保存列表
            "answered_questions": self.answered_questions,   # 直接保存列表
            "last_imported_docx": self.last_imported_docx
        }
        try:
            with open(self.SAVE_FILE_NAME, "wb") as f:
                pickle.dump(data_to_save, f)
            if not silent:
                messagebox.showinfo("保存成功", f"进度已保存到 {self.SAVE_FILE_NAME}")
        except Exception as e:
            if not silent:
                messagebox.showerror("保存失败", f"保存进度时发生错误: {e}")
            print(f"Error saving progress: {e}")

    def load_progress(self):
        if not os.path.exists(self.SAVE_FILE_NAME):
            self.question_text_label.config(text="未找到保存的进度文件。请导入新题库。")
            return

        try:
            with open(self.SAVE_FILE_NAME, "rb") as f:
                loaded_data = pickle.load(f)
            
            self.all_questions = loaded_data.get("all_questions", [])
            self.unanswered_questions = loaded_data.get("unanswered_questions", [])
            self.answered_questions = loaded_data.get("answered_questions", [])
            self.last_imported_docx = loaded_data.get("last_imported_docx")

            if not self.all_questions:
                 self.question_text_label.config(text="加载的进度为空。请导入新题库。")
                 return

            # 恢复UI状态
            self.answered_listbox.delete(0, tk.END)
            for q in self.answered_questions:
                q_preview = f"{q.q_type} (原序 {q.original_doc_order + 1}) {q.text[:40]}..."
                self.answered_listbox.insert(tk.END, q_preview)
            
            self.update_stats()
            self.clear_question_display() # 清空当前题目显示区
            self.question_text_label.config(text=f"成功加载 {len(self.all_questions)} 道题目。请点击“随机抽题”。")
            messagebox.showinfo("加载成功", f"已从 {self.SAVE_FILE_NAME} 加载进度。")

        except Exception as e:
            messagebox.showerror("加载失败", f"加载进度时发生错误: {e}\n可能需要重新导入题库。")
            print(f"Error loading progress: {e}")
            # 如果加载失败，清空数据以防万一
            self.all_questions = []
            self.unanswered_questions = []
            self.answered_questions = []
            self.update_stats()


    def parse_questions_from_docx(self, filepath):
        # ... (这个函数保持您上一版本中能工作的那个)
        # 我将使用您上一条回复中修正后的 flush_buffer_to_question 逻辑
        # print(f"--- 开始解析文档: {filepath} ---")
        doc = docx.Document(filepath)
        parsed_questions = []
        
        current_q_type = None
        question_buffer = [] 
        doc_line_counter = 0 

        def get_question_type(line_text):
            line_text = line_text.strip()
            if "单选题" in line_text: return "单选题"
            if "多选题" in line_text: return "多选题"
            if "填空题" in line_text: return "填空题"
            if "判断题" in line_text: return "判断题"
            return None

        def flush_buffer_to_question():
            nonlocal question_buffer, current_q_type, doc_line_counter
            # print(f"    尝试 flush_buffer_to_question. 当前类型: {current_q_type}, Buffer内容: {question_buffer}")
            if not question_buffer or not current_q_type:
                # print("      Buffer为空或当前类型未设置，清空并返回。")
                question_buffer = [] 
                return

            if len(question_buffer) == 1:
                # print("      Buffer中只有一个元素，尝试按换行符分割。")
                all_lines_in_block = question_buffer[0].splitlines() 
                if not all_lines_in_block:
                    # print("      分割后行列表为空，清空并返回。")
                    question_buffer = []
                    return
            else:
                # print(f"      Buffer中有多个元素 ({len(question_buffer)}个)，将它们合并后按换行符分割。")
                combined_text = "\n".join(question_buffer)
                all_lines_in_block = combined_text.splitlines()
            
            # print(f"      分割后的行列表 (all_lines_in_block): {all_lines_in_block}")

            answer_line_text = None
            answer_line_index_in_block = -1 

            for i, line_in_block in enumerate(all_lines_in_block):
                if "正确答案" in line_in_block:
                    answer_line_text = line_in_block
                    answer_line_index_in_block = i
                    # print(f"      在分割后的行列表中找到 '正确答案' 在索引 {i}: '{answer_line_text}'")
                    break
            
            if answer_line_text is None: 
                # print(f"      分割后的行列表中未找到 '正确答案'。行列表: {all_lines_in_block}")
                question_buffer = [] 
                return

            question_content_lines_from_block = all_lines_in_block[:answer_line_index_in_block]
            if not question_content_lines_from_block: 
                # print(f"      从分割后的行列表看，题干为空。答案行: {answer_line_text}")
                question_buffer = []
                return
            
            # print(f"      准备创建Question对象。题干行(来自block): {question_content_lines_from_block[0]}, 选项行(来自block): {question_content_lines_from_block[1:]}, 答案文本: {answer_line_text}")
            original_num_and_text = question_content_lines_from_block[0]
            options_lines = question_content_lines_from_block[1:]

            try:
                q_obj = Question(
                    q_type=current_q_type,
                    original_num_text=original_num_and_text.strip(), 
                    text="", 
                    options_text=[opt.strip() for opt in options_lines], 
                    answer_text=answer_line_text.strip(), 
                    original_doc_order=doc_line_counter 
                )
                parsed_questions.append(q_obj)
                doc_line_counter +=1
                # print(f"      成功创建并添加Question对象: {q_obj}")
            except Exception as e:
                print(f"      创建Question对象时出错: {e} -- 内容: {question_content_lines_from_block} | 答案: {answer_line_text}")
                import traceback
                traceback.print_exc()
            question_buffer = [] 

        for para_idx, para in enumerate(doc.paragraphs): 
            text = para.text.strip()
            # print(f"\n处理文档段落 {para_idx + 1}: '{text}'") 

            if not text: 
                # print("  空行，跳过。") 
                continue
            
            text = text.replace('↓', '').replace('←', '')
            new_q_type = get_question_type(text)
            # print(f"  识别到的新题型: {new_q_type}") 

            if new_q_type:
                # print(f"  遇到新题型 '{new_q_type}'。尝试清空旧buffer。") 
                flush_buffer_to_question()
                current_q_type = new_q_type
                # print(f"  当前题型更新为: {current_q_type}") 
            elif current_q_type: 
                # print(f"  非题型行，当前类型为 '{current_q_type}'，将 '{text}' 加入buffer。") 
                question_buffer.append(text)
                if "正确答案" in text:
                    # print(f"  当前行是 '正确答案' 行，尝试清空buffer。") 
                    flush_buffer_to_question()
            # else:
                # print(f"  非题型行，且当前题型未设置 (current_q_type is None)，忽略此行: '{text}'") 
        
        # print("\n--- 文档段落遍历结束 ---") 
        # print("尝试最后一次 flush_buffer_to_question (处理文档末尾可能剩余的题目)") 
        flush_buffer_to_question()
        
        # print(f"--- 解析完成，共解析出 {len(parsed_questions)} 个题目 ---") 
        # print(f"解析出的题目列表: {parsed_questions}") 
        return parsed_questions


    def import_word_file(self):
        filepath = filedialog.askopenfilename(
            title="选择Word题库文件",
            filetypes=(("Word documents", "*.docx"), ("All files", "*.*"))
        )
        if not filepath:
            return

        # 询问是否覆盖现有进度（如果已加载或已有题目）
        if self.all_questions:
            if not messagebox.askyesno("确认导入", "当前已有题库数据。导入新题库将覆盖现有数据和进度，确定吗？"):
                return
        
        try:
            self.all_questions = self.parse_questions_from_docx(filepath)
            
            if not self.all_questions:
                messagebox.showwarning("导入问题", "未能从文档中解析出任何题目。请检查文档格式。")
                return

            self.last_imported_docx = filepath # 记录文件路径
            self.unanswered_questions = list(self.all_questions) 
            random.shuffle(self.unanswered_questions)
            self.answered_questions = []
            self.answered_listbox.delete(0, tk.END)
            self.update_stats()
            self.clear_question_display()
            self.question_text_label.config(text=f"成功导入 {len(self.all_questions)} 道题目！请点击“随机抽题”。")
            messagebox.showinfo("成功", f"题库导入成功，共 {len(self.all_questions)} 道题目。")
            # 导入新题库后，也应该保存一下
            self.save_progress(silent=True)

        except Exception as e:
            messagebox.showerror("导入错误", f"无法解析Word文件或处理题目: {e}")
            print(f"导入或解析过程中发生错误: {e}") 
            import traceback
            traceback.print_exc() 

    # --- 其他方法 (update_stats, clear_question_display, display_random_question, process_answer, move_to_unanswered) ---
    # --- 保持与您上一版本能工作的代码一致 ---
    def update_stats(self):
        self.stats_label.config(text=f"未答题: {len(self.unanswered_questions)} | 已答题: {len(self.answered_questions)}")
        self.btn_random_question.config(state=tk.NORMAL if self.unanswered_questions else tk.DISABLED)
        self.btn_move_back.config(state=tk.NORMAL if self.answered_questions else tk.DISABLED)
        # 如果没有题目，禁用保存按钮可能也是个好主意
        self.btn_save_progress.config(state=tk.NORMAL if self.all_questions else tk.DISABLED)


    def clear_question_display(self):
        self.question_header_label.config(text="")
        self.question_text_label.config(text="请点击“随机抽题”或加载进度")
        self.answer_display_label.config(text="")
        for widget in self.options_frame.winfo_children():
            widget.destroy()
        self.user_answer_widgets = []
        self.current_question_data = None
        self.btn_show_answer.config(state=tk.DISABLED)


    def display_random_question(self):
        if not self.unanswered_questions:
            messagebox.showinfo("提示", "所有题目都已作答完毕！")
            self.clear_question_display()
            return

        self.current_question_data = random.choice(self.unanswered_questions)
        q = self.current_question_data 

        # 更新题目头部和正文的字体（如果它们是在这里重新配置的）
        self.question_header_label.config(text=f"{q.q_type} - (原序 {q.original_doc_order + 1})", font=self.QUESTION_FONT)
        self.question_text_label.config(text=q.get_display_text(), font=self.QUESTION_FONT)
        
        for widget in self.options_frame.winfo_children():
            widget.destroy()
        self.user_answer_widgets = []
        self.answer_display_label.config(text="")

        if q.q_type == "填空题":
            self.question_text_label.config(text=q.get_display_text())
            num_blanks = len(q.answer) if q.answer else 0
            if num_blanks == 0 and "___" in q.text:
                 num_blanks = q.text.count("___")
            if num_blanks == 0 and q.answer and q.answer[0] == q.answer_raw.replace("正确答案", "").strip("：: ").strip():
                 num_blanks = 1
            if num_blanks == 0 and not q.answer: # 真正没有答案解析出来
                 num_blanks = 1 #至少给一个空

            if num_blanks > 0:
                for i in range(num_blanks):
                    entry_label = tk.Label(self.options_frame, text=f"填空 {i+1}:", font=self.OPTION_FONT)
                    entry_label.pack(side=tk.LEFT, padx=(0,5))
                    entry = tk.Entry(self.options_frame, width=20, font=self.OPTION_FONT)
                    entry.pack(side=tk.LEFT, padx=(0,10))
                    self.user_answer_widgets.append(entry)
            else: # 理论上至少有一个空
                 tk.Label(self.options_frame, text="(请直接思考答案)").pack(anchor='w')


        elif q.q_type in ["单选题", "判断题"]:
            self.question_text_label.config(text=q.get_display_text())
            self.var_choice = tk.StringVar(value=None)
            options_to_display = q.options
            if q.q_type == "判断题": 
                options_to_display = {"A": "是 (正确)", "B": "否 (错误)"} 
                # 再次确保判断题答案与选项一致性
                if q.answer not in ["A", "B"]: # 如果解析的答案仍然不是A或B
                    cleaned_ans_text = q.answer_raw.replace("正确答案", "").strip("：: ").upper()
                    if "A" in cleaned_ans_text or "是" in cleaned_ans_text or "正确" in cleaned_ans_text: q.answer = "A"
                    elif "B" in cleaned_ans_text or "否" in cleaned_ans_text or "错误" in cleaned_ans_text: q.answer = "B"
                    else: q.answer = None # 无法确定


            for letter, opt_text in options_to_display.items():
                rb = tk.Radiobutton(self.options_frame, text=f"{letter}. {opt_text}", 
                                    variable=self.var_choice, value=letter, 
                                    wraplength=700, justify=tk.LEFT, 
                                    font=self.OPTION_FONT) # <--- 修改字体
                # rb = tk.Radiobutton(self.options_frame, text=f"{letter}. {opt_text}", variable=self.var_choice, value=letter, wraplength=700, justify=tk.LEFT)
                rb.pack(anchor="w")
                self.user_answer_widgets.append(rb) 

        elif q.q_type == "多选题":
            self.question_text_label.config(text=q.get_display_text())
            self.vars_multi_choice = {}
            for letter, opt_text in q.options.items():
                var = tk.BooleanVar()
                cb = tk.Checkbutton(self.options_frame, text=f"{letter}. {opt_text}", 
                                     variable=var, wraplength=700, 
                                     justify=tk.LEFT, 
                                     font=self.OPTION_FONT)
                # cb = tk.Checkbutton(self.options_frame, text=f"{letter}. {opt_text}", variable=var, wraplength=700, justify=tk.LEFT)
                cb.pack(anchor="w")
                self.vars_multi_choice[letter] = var
                self.user_answer_widgets.append(cb)

        else: 
            self.question_text_label.config(text=q.get_display_text())
            tk.Label(self.options_frame, text="(未知题型，请直接思考答案)").pack(anchor='w')

        self.btn_show_answer.config(state=tk.NORMAL)


    def process_answer(self):
        if not self.current_question_data:
            return

        q = self.current_question_data 
        correct_answer_str = ""
        user_answer_str = ""
        
        if q.q_type == "单选题":
            correct_answer_str = f"正确答案: {q.answer or '未提供'}. {q.options.get(q.answer, '') if q.answer else ''}"
            user_answer_val = self.var_choice.get()
            user_answer_str = f"您的选择: {user_answer_val or '未选择'}. {q.options.get(user_answer_val, '') if user_answer_val else ''}"
        
        elif q.q_type == "多选题":
            ans_list = q.answer if isinstance(q.answer, list) else []
            sorted_ans = sorted(ans_list)
            correct_answer_str = f"正确答案: {''.join(sorted_ans) or '未提供'}\n"
            for ans_letter in sorted_ans:
                 correct_answer_str += f"  {ans_letter}. {q.options.get(ans_letter, '')}\n"
            
            user_choices = sorted([letter for letter, var in self.vars_multi_choice.items() if var.get()])
            user_answer_str = f"您的选择: {''.join(user_choices) or '未选择'}\n"
            for choice_letter in user_choices:
                 user_answer_str += f"  {choice_letter}. {q.options.get(choice_letter, '')}\n"

        elif q.q_type == "判断题":
            display_options = {"A": "是 (正确)", "B": "否 (错误)"}
            correct_answer_str = f"正确答案: {q.answer or '未提供'}. {display_options.get(q.answer, '(答案解析可能不匹配)') if q.answer else ''}"
            user_answer_val = self.var_choice.get()
            user_answer_str = f"您的选择: {user_answer_val or '未选择'}. {display_options.get(user_answer_val, '') if user_answer_val else ''}"

        elif q.q_type == "填空题":
            ans_list = q.answer if isinstance(q.answer, list) else []
            correct_answer_str = f"正确答案: {' | '.join(ans_list or ['未提供'])}"
            user_fills = [widget.get() for widget in self.user_answer_widgets if isinstance(widget, tk.Entry)]
            user_answer_str = f"您的填写: {' | '.join(user_fills) or '未填写'}"
        else:
            correct_answer_str = "正确答案: (未知题型)"
            user_answer_str = "您的作答: (未知题型)"

        self.answer_display_label.config(text=f"{user_answer_str}\n{correct_answer_str}")
        
        if self.current_question_data in self.unanswered_questions: # 确保它确实还在未答列表
            self.unanswered_questions.remove(self.current_question_data)
            self.answered_questions.append(self.current_question_data)
            
            q_preview = f"{q.q_type} (原序 {q.original_doc_order + 1}) {q.text[:40]}..."
            self.answered_listbox.insert(tk.END, q_preview)
            
        self.update_stats()
        self.btn_show_answer.config(state=tk.DISABLED) 


    def move_to_unanswered(self):
        selected_indices = self.answered_listbox.curselection()
        if not selected_indices:
            messagebox.showwarning("提示", "请先在“已答题目列表”中选择要移回的题目。")
            return

        moved_count = 0
        # 从后往前删，避免索引问题
        for i in sorted(selected_indices, reverse=True):
            try:
                question_to_move = self.answered_questions.pop(i) 
                self.unanswered_questions.append(question_to_move)
                self.answered_listbox.delete(i) 
                moved_count +=1
            except IndexError:
                print(f"移回题目时索引错误: {i}") 

        if moved_count > 0:
            random.shuffle(self.unanswered_questions) 
            self.update_stats()
            messagebox.showinfo("成功", f"成功将 {moved_count} 道题目移回未答列表。")


# --- Main ---
if __name__ == "__main__":
    root = tk.Tk()
    app = QuizApp(root)
    root.mainloop()