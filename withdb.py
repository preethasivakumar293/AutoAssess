import os
import tkinter as tk
from tkinter import filedialog, messagebox, Toplevel, Text, Scrollbar, StringVar, ttk
from docx import Document
import re
import pandas as pd
import matplotlib.pyplot as plt
import sqlite3
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from datetime import datetime
from PIL import Image, ImageTk

# Directory containing subject files
DATA_FOLDER = os.path.dirname(os.path.abspath(__file__))

db_path = os.path.join(os.getcwd(), 'evaluation_results.db')

# Function to create the database and table if not exists
def create_db():
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()

    # Create the evaluation_results table if it doesn't exist
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS evaluation_results (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            student_name TEXT,
            total_marks REAL,
            submission_time TEXT
        )
    """)

    conn.commit()
    conn.close()

# Function to store evaluation results in the database
def store_results_in_db(results):
    try:
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()

        # Ensure the table exists before inserting
        create_db()

        # Insert data into the database
        for student_name, total_marks in results:
            submission_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            cursor.execute("""
                INSERT INTO evaluation_results (student_name, total_marks, submission_time)
                VALUES (?, ?, ?)
            """, (student_name, total_marks, submission_time))

        conn.commit()
    except sqlite3.Error as e:
        messagebox.showerror("Database Error", f"Failed to store results: {str(e)}")
    finally:
        if conn:
            conn.close()

# Function to validate login
def validate_login(username, password):
    return username == "admin" and password == "a"

# Function to fetch available subjects from DATA_FOLDER
def get_available_subjects():
    if not os.path.exists(DATA_FOLDER):
        return []
    files = os.listdir(DATA_FOLDER)
    subjects = [file.split("_")[0] for file in files if file.endswith("QuestionPaper.docx")]
    return list(set(subjects))

# Function to extract text from DOCX
def extract_text(file_path):
    doc = Document(file_path)
    return "\n".join([para.text.strip() for para in doc.paragraphs if para.text.strip()])

# Improved function to extract questions and keywords with special handling for tables
def extract_keywords(question_file):
    doc = Document(question_file)
    keyword_dict = {}
    
    # Extract text from document
    full_text = ""
    for para in doc.paragraphs:
        full_text += para.text + "\n"
    
    # Match question patterns: Q1, Q2, etc. followed by any text until the next question or end of file
    questions = re.findall(r'(Q\d+:.*?)(?=Q\d+:|$)', full_text, re.DOTALL)
    
    for q in questions:
        # Extract question number
        q_num_match = re.match(r'Q(\d+):', q)
        if q_num_match:
            q_num = f"Q{q_num_match.group(1)}"
            
            # Extract keywords section
            keyword_match = re.search(r'Keywords:(.*?)(?=Q\d+:|$)', q, re.DOTALL | re.IGNORECASE)
            if keyword_match:
                # Extract individual keywords
                keywords_text = keyword_match.group(1).lower()
                # Match keywords between ** marks or in bullet points
                keyword_items = re.findall(r'\*\*(.*?)\*\*|[-•]?\s+(.*?)(?=[-•]|\n|$)', keywords_text, re.DOTALL)
                
                # Process the matched items
                keywords = set()
                for item in keyword_items:
                    kw = item[0] if item[0] else item[1]
                    kw = kw.strip().lower()
                    if kw:
                        keywords.add(kw)
                
                keyword_dict[q_num] = keywords
   
    
    return keyword_dict

# Improved function to extract student answers
def extract_answers(answer_text):
    answers = {}
    
    # Match Q1.Answer, Q2.Answer, etc. followed by any text until the next answer
    qa_pattern = re.findall(r'Q(\d+)\.?\s*Answer:?(.*?)(?=Q\d+\.?\s*Answer:|$)', answer_text, re.DOTALL | re.IGNORECASE)
    
    for q_num, ans in qa_pattern:
        q_key = f"Q{q_num}"
        cleaned_answer = ans.strip().lower()
        answers[q_key] = cleaned_answer
    
    return answers

# Enhanced function to check keyword matches with specialized handling for different question types
def check_keyword_matches(student_answer, keywords, q_num):
    student_answer = student_answer.lower()
    matched_keywords = []
    
    # Question-specific handling
    if q_num == "Q4":
        table_terms = {
            "ram": ["volatile", "temporary"],
            "rom": ["non-volatile", "permanent"],
            "volatile": ["data is lost", "power off", "temporary"],
            "non-volatile": ["data is retained", "power off", "permanent"],
            "access time": ["faster", "slower"],
            "usage": ["temporary storage", "firmware", "boot"]
        }
        
        for term, related_terms in table_terms.items():
            if term in student_answer:
                matched_keywords.append(term)
                for related in related_terms:
                    if related in student_answer and related in keywords:
                        matched_keywords.append(related)
    
    elif q_num == "Q9":
        code_patterns = {
            "unsigned int": r'unsigned\s+int',
            "500": r'=\s*500',
            "printf": r'printf',
            "%u": r'%u',
            "#include": r'#include',
            "main": r'main\(\)',
            "return 0": r'return\s+0'
        }
        
        for keyword, pattern in code_patterns.items():
            if re.search(pattern, student_answer):
                matched_keywords.append(keyword)
    
    elif q_num == "Q10":
        bitwise_ops = ["&", "|", "^", "~", "<<", ">>"]
        logical_ops = ["&&", "||", "!"]
        
        bitwise_mentions = sum(1 for op in bitwise_ops if op in student_answer or f"bitwise {op}" in student_answer)
        logical_mentions = sum(1 for op in logical_ops if op in student_answer or f"logical {op}" in student_answer)
        
        has_definition = any(term in student_answer for term in ["definition", "defined as", "refers to"])
        has_examples = any(term in student_answer for term in ["example", "such as", "instance"])
        has_comparison = any(term in student_answer for term in ["compare", "different", "distinction", "versus", "vs"])
        
        if "bitwise" in student_answer or bitwise_mentions >= 2:
            matched_keywords.append("bitwise operators")
            for op_name, op_symbols in [("and", "&"), ("or", "|"), ("xor", "^"), ("not", "~"), 
                                      ("left shift", "<<"), ("right shift", ">>")]:
                if op_name in student_answer or op_symbols in student_answer:
                    matched_keywords.append(op_name)
                    if op_symbols in keywords:
                        matched_keywords.append(op_symbols)
        
        if "logical" in student_answer or logical_mentions >= 2:
            matched_keywords.append("logical operators")
            for op_name, op_symbols in [("logical and", "&&"), ("logical or", "||"), ("logical not", "!")]:
                if op_name in student_answer or op_symbols in student_answer:
                    matched_keywords.append(op_name.replace("logical ", ""))
                    matched_keywords.append(op_symbols)
        
        if any(term in student_answer for term in ["bit manipulation", "bit level", "bit operations"]):
            matched_keywords.append("bit manipulation")
        
        if any(term in student_answer for term in ["boolean", "true/false", "logical conditions"]):
            matched_keywords.append("boolean expressions")
        
        if has_definition and has_comparison:
            matched_keywords.append("definition")
        
        if "integer" in student_answer and "boolean" in student_answer:
            matched_keywords.append("operands")
    
    # General keyword matching for all questions
    for keyword in keywords:
        if keyword in matched_keywords:
            continue
        
        if keyword.lower() in student_answer:
            matched_keywords.append(keyword)
        else:
            keyword_words = set(re.findall(r'\b\w+\b', keyword.lower()))
            if len(keyword_words) > 1 and all(word in student_answer for word in keyword_words):
                matched_keywords.append(keyword)
            elif keyword in ["algorithm", "procedure", "computation"]:
                synonyms = {
                    "algorithm": ["method", "set of steps", "process", "rules", "instructions"],
                    "procedure": ["process", "method", "routine", "sequence", "steps"],
                    "computation": ["calculation", "processing", "computing", "compute", "solve"]
                }
                if keyword in synonyms and any(syn in student_answer for syn in synonyms[keyword]):
                    matched_keywords.append(keyword)
    
    return list(set(matched_keywords))

# Function to evaluate all students
def evaluate_all_students(subject):
    question_file = os.path.join(DATA_FOLDER, f"{subject}_QuestionPaper.docx")
    answers_folder = answer_folder_var.get()

    if not os.path.exists(question_file):
        messagebox.showerror("Error", f"Question paper for '{subject}' not found.")
        return None
    if not answers_folder or not os.path.exists(answers_folder):
        messagebox.showerror("Error", "Please select a valid answer sheets folder.")
        return None

    keyword_dict = extract_keywords(question_file)
    if not keyword_dict:
        messagebox.showerror("Error", "No keywords found in the question paper.")
        return None

    all_results = []
    all_students_data = []
    
    for file in os.listdir(answers_folder):
        if file.endswith(".docx"):
            student_name = file.split("_")[0]
            answer_path = os.path.join(answers_folder, file)
            answer_text = extract_text(answer_path)
            answers = extract_answers(answer_text)

            score = 0
            max_marks = len(keyword_dict) * 2
            results = []
            question_scores = {}

            for q_num, keywords in keyword_dict.items():
                student_answer = answers.get(q_num, "")
                matched_keywords = check_keyword_matches(student_answer, keywords, q_num)
                
                question_score = 0
                if keywords:
                    if q_num == "Q10":
                        has_bitwise = any(kw for kw in matched_keywords if "bitwise" in kw)
                        has_logical = any(kw for kw in matched_keywords if "logical" in kw)
                        has_operators = any(op in matched_keywords for op in ["&", "|", "^", "~", "<<", ">>", "&&", "||", "!"])
                        has_explanation = any(ex in matched_keywords for ex in ["definition", "bit manipulation", "boolean expressions", "operands"])
                        
                        if has_bitwise and has_logical and has_operators and has_explanation:
                            question_score = 2.0
                        elif (has_bitwise and has_logical) and (has_operators or has_explanation):
                            question_score = 1.5
                        elif has_bitwise or has_logical:
                            question_score = 1.0
                        else:
                            keyword_weight = 2.0 / len(keywords)
                            question_score = len(matched_keywords) * keyword_weight
                            question_score = min(question_score, 2.0)
                    else:
                        keyword_weight = 2.0 / len(keywords)
                        question_score = len(matched_keywords) * keyword_weight
                        question_score = min(question_score, 2.0)
                
                score += question_score
                question_scores[q_num] = question_score
                results.append(f"{q_num}: Marks {question_score:.2f}/2, Matched: {', '.join(matched_keywords) if matched_keywords else 'None'}")
            
            all_results.append((student_name, score, max_marks, results))
            
            student_row = {"Student": student_name, "Total": f"{score:.2f}/{max_marks}"}
            for q_num in keyword_dict.keys():
                student_row[q_num] = f"{question_scores.get(q_num, 0):.2f}/2"
            all_students_data.append(student_row)
    
    # Store all results in database
    store_results_in_db([(student_name, score) for student_name, score, _, _ in all_results])
    
    return all_results, all_students_data, keyword_dict

def visualize_summary_stats(stats_frame, results):
    total_scores = [score for _, score, _, _ in results]
    avg_score = sum(total_scores) / len(total_scores) if total_scores else 0
    max_score = max(total_scores) if total_scores else 0
    min_score = min(total_scores) if total_scores else 0

    top_performers = sorted(results, key=lambda x: x[1], reverse=True)[:3]
    top_performers_names = [x[0] for x in top_performers]
    top_performers_scores = [x[1] for x in top_performers]

    fig, ax = plt.subplots(1, 2, figsize=(10, 5))

    ax[0].bar(['Max', 'Min', 'Avg'], [max_score, min_score, avg_score], color=['green', 'red', 'blue'])
    ax[0].set_title('Highest, Lowest, and Average Scores')
    ax[0].set_ylabel('Score')

    ax[1].bar(top_performers_names, top_performers_scores, color='orange')
    ax[1].set_title('Top Performers')
    ax[1].set_ylabel('Score')

    canvas = FigureCanvasTkAgg(fig, stats_frame)
    canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
    canvas.draw()

def display_results(root, results_data):
    results, all_students_data, keyword_dict = results_data
    
    results_window = Toplevel(root)
    results_window.title("Evaluation Results")
    results_window.geometry("800x600")
    
    notebook = ttk.Notebook(results_window)
    notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
    
    # Tab 1: Detailed Results
    detailed_tab = ttk.Frame(notebook)
    notebook.add(detailed_tab, text="Detailed Results")
    
    export_frame = tk.Frame(detailed_tab)
    export_frame.pack(fill=tk.X, pady=5)
    
    def export_detailed_results():
        export_path = filedialog.asksaveasfilename(
            defaultextension=".txt", 
            filetypes=[("Text files", "*.txt"), ("Word Document", "*.docx")]
        )
        if export_path:
            if export_path.endswith(".txt"):
                with open(export_path, 'w') as f:
                    f.write(f"--- Evaluation Results: {datetime.now().strftime('%Y-%m-%d %H:%M')} ---\n\n")
                    for student_name, total_marks, max_marks, details in results:
                        f.write(f"Student: {student_name}, Total Marks: {total_marks:.2f}/{max_marks}\n\n")
                        for result in details:
                            f.write(result + "\n" + "-" * 40 + "\n")
                        f.write("\n")
                messagebox.showinfo("Export Successful", f"Detailed results exported to {export_path}")
            elif export_path.endswith(".docx"):
                doc = Document()
                doc.add_heading(f"Evaluation Results: {datetime.now().strftime('%Y-%m-%d %H:%M')}", 0)
                
                for student_name, total_marks, max_marks, details in results:
                    doc.add_heading(f"Student: {student_name}", 1)
                    doc.add_paragraph(f"Total Marks: {total_marks:.2f}/{max_marks}")
                    
                    for result in details:
                        doc.add_paragraph(result)
                        doc.add_paragraph("-" * 40)
                    
                    doc.add_paragraph("")
                
                doc.save(export_path)
                messagebox.showinfo("Export Successful", f"Detailed results exported to {export_path}")
    
    tk.Button(export_frame, text="Export Detailed Results", command=export_detailed_results).pack(pady=5)

    text_frame = tk.Frame(detailed_tab)
    text_frame.pack(fill=tk.BOTH, expand=True)
    
    text_area = Text(text_frame, wrap='word', font=("Arial", 12))
    scrollbar = Scrollbar(text_frame, command=text_area.yview)
    text_area.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    text_area.config(yscrollcommand=scrollbar.set)

    text_area.insert(tk.END, f"--- Evaluation Results: {datetime.now().strftime('%Y-%m-%d %H:%M')} ---\n\n")
    
    results.sort(key=lambda x: x[0])
    
    for student_name, total_marks, max_marks, details in results:
        text_area.insert(tk.END, f"Student: {student_name}, Total Marks: {total_marks:.2f}/{max_marks}\n\n")
        for result in details:
            text_area.insert(tk.END, result + "\n" + "-" * 40 + "\n")
        text_area.insert(tk.END, "\n")
    text_area.configure(state='disabled')
    
    # Tab 2: Summary Table
    summary_tab = ttk.Frame(notebook)
    notebook.add(summary_tab, text="Summary Table")
    
    df = pd.DataFrame(all_students_data)
    
    export_summary_frame = tk.Frame(summary_tab)
    export_summary_frame.pack(fill=tk.X, pady=5)
    
    def export_summary():
        export_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx", 
            filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv")]
        )
        if export_path:
            if export_path.endswith(".xlsx"):
                df.to_excel(export_path, index=False)
            elif export_path.endswith(".csv"):
                df.to_csv(export_path, index=False)
            messagebox.showinfo("Export Successful", f"Summary exported to {export_path}")
    
    tk.Button(export_summary_frame, text="Export Summary", command=export_summary).pack(pady=5)
    
    tree_frame = tk.Frame(summary_tab)
    tree_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
    
    y_scrollbar = ttk.Scrollbar(tree_frame)
    y_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    
    x_scrollbar = ttk.Scrollbar(tree_frame, orient="horizontal")
    x_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)
    
    columns = list(df.columns)
    tree = ttk.Treeview(tree_frame, columns=columns, show="headings", 
                        yscrollcommand=y_scrollbar.set, xscrollcommand=x_scrollbar.set)
    
    y_scrollbar.config(command=tree.yview)
    x_scrollbar.config(command=tree.xview)
    
    for col in columns:
        tree.heading(col, text=col)
        tree.column(col, width=100, anchor="center")
    
    for i, row in df.iterrows():
        values = [row[col] for col in columns]
        tree.insert("", "end", values=values)
    
    tree.pack(fill=tk.BOTH, expand=True)
    
    # Tab 3: Statistics
    stats_tab = ttk.Frame(notebook)
    notebook.add(stats_tab, text="Statistics")

    stats_frame = tk.Frame(stats_tab)
    stats_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

    total_scores = [score for _, score, _, _ in results]
    avg_score = sum(total_scores) / len(total_scores) if total_scores else 0
    max_score = max(total_scores) if total_scores else 0
    min_score = min(total_scores) if total_scores else 0

    top_performers = sorted(results, key=lambda x: x[1], reverse=True)[:3]

    text_frame = tk.Frame(stats_frame)
    text_frame.pack(fill=tk.BOTH, expand=True)

    stats_text = Text(text_frame, wrap='word', font=("Arial", 12), height=15)
    stats_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    scrollbar = Scrollbar(text_frame, command=stats_text.yview)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    stats_text.config(yscrollcommand=scrollbar.set)

    stats_text.insert(tk.END, "--- Class Statistics ---\n\n")
    stats_text.insert(tk.END, f"Total Students: {len(results)}\n")
    stats_text.insert(tk.END, f"Average Score: {avg_score:.2f}\n")
    stats_text.insert(tk.END, f"Highest Score: {max_score:.2f}\n")
    stats_text.insert(tk.END, f"Lowest Score: {min_score:.2f}\n\n")

    stats_text.insert(tk.END, "--- Top Performers ---\n\n")
    for i, (name, score, max_marks, _) in enumerate(top_performers, 1):
        stats_text.insert(tk.END, f"{i}. {name}: {score:.2f}/{max_marks}\n")

    stats_text.insert(tk.END, "\n--- Question Performance ---\n\n")

    question_avgs = {}
    for student_data in all_students_data:
        for q_num in student_data:
            if q_num != "Student" and q_num != "Total":
                score = float(student_data[q_num].split('/')[0])
                if q_num not in question_avgs:
                    question_avgs[q_num] = []
                question_avgs[q_num].append(score)

    for q_num, scores in sorted(question_avgs.items()):
        avg = sum(scores) / len(scores) if scores else 0
        stats_text.insert(tk.END, f"{q_num}: Average score {avg:.2f}/2\n")

    visualize_summary_stats(stats_frame, results)

    stats_text.configure(state='disabled')
    
    # Tab 4: Question 10 Specific Analysis
    q10_tab = ttk.Frame(notebook)
    notebook.add(q10_tab, text="Q10 Analysis")
    
    q10_frame = tk.Frame(q10_tab)
    q10_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
    
    q10_text = Text(q10_frame, wrap='word', font=("Arial", 12), height=20)
    q10_text.pack(fill=tk.BOTH, expand=True)
    
    q10_text.insert(tk.END, "--- Question 10 Analysis ---\n\n")
    q10_text.insert(tk.END, "Question: Distinguish between bitwise operators and logical operators.\n\n")
    
    q10_scores = []
    q10_details = []
    for student_name, total_marks, max_marks, details in results:
        for detail in details:
            if detail.startswith("Q10:"):
                q10_details.append((student_name, detail))
                score_part = detail.split("Marks ")[1].split("/")[0]
                q10_scores.append(float(score_part))
    
    q10_avg = sum(q10_scores) / len(q10_scores) if q10_scores else 0
    q10_max = max(q10_scores) if q10_scores else 0
    q10_min = min(q10_scores) if q10_scores else 0
    
    q10_text.insert(tk.END, f"Average Score: {q10_avg:.2f}/2\n")
    q10_text.insert(tk.END, f"Highest Score: {q10_max:.2f}/2\n")
    q10_text.insert(tk.END, f"Lowest Score: {q10_min:.2f}/2\n\n")
    
    keyword_counts = {}
    for _, detail in q10_details:
        matched_part = detail.split("Matched: ")[1]
        if matched_part != "None":
            keywords = [kw.strip() for kw in matched_part.split(",")]
            for kw in keywords:
                if kw not in keyword_counts:
                    keyword_counts[kw] = 0
                keyword_counts[kw] += 1
    
    q10_text.insert(tk.END, "Most Common Keywords Matched:\n")
    for kw, count in sorted(keyword_counts.items(), key=lambda x: x[1], reverse=True)[:10]:
        q10_text.insert(tk.END, f"- {kw}: {count} students\n")
    
    q10_text.insert(tk.END, "\nMost Common Missing Keywords:\n")
    all_q10_keywords = keyword_dict.get("Q10", set())
    missing_keywords = {kw: 0 for kw in all_q10_keywords}
    for _, detail in q10_details:
        matched_part = detail.split("Matched: ")[1]
        if matched_part != "None":
            matched_keywords = [kw.strip() for kw in matched_part.split(",")]
            for kw in all_q10_keywords:
                if kw not in matched_keywords:
                    missing_keywords[kw] += 1
    
    for kw, count in sorted(missing_keywords.items(), key=lambda x: x[1], reverse=True)[:10]:
        q10_text.insert(tk.END, f"- {kw}: missing in {count} students\n")
    
    q10_text.configure(state='disabled')

def main_app():
    global answer_folder_var
    def set_background(widget, image_path):
        image = Image.open(image_path)
        image = image.resize((850, 600), Image.Resampling.LANCZOS)
        bg_image = ImageTk.PhotoImage(image)
        canvas = tk.Canvas(widget, width=850, height=600)
        canvas.pack(fill="both", expand=True)
        canvas.create_image(0, 0, anchor="nw", image=bg_image)
        canvas.image = bg_image  # Keep a reference
        return canvas

    root = tk.Tk()
    root.title("Answer Evaluation System")
    root.geometry("850x600")
    root.resizable(False, False)

    style = ttk.Style(root)
    style.theme_use("clam")

    style.configure("TFrame", background="#f0f0f0")
    style.configure("TButton", background="#4CAF50", foreground="black", font=('Arial', 10, 'bold'))
    style.configure("TLabel", background="#f0f0f0", font=('Arial', 12))

    username_var = StringVar()
    password_var = StringVar()
    answer_folder_var = StringVar()

    def handle_login():
        if validate_login(username_var.get(), password_var.get()):
            show_subject_selection()
        else:
            messagebox.showerror("Login Failed", "Invalid username or password.")

    def show_subject_selection():
        for widget in root.winfo_children():
            widget.destroy()

        bg_canvas = set_background(root, "bg.jpg")

        subjects = get_available_subjects()
        if not subjects:
            messagebox.showerror("Error", "No question papers found.")
            return

        main_frame = ttk.Frame(root, padding="20 20 20 20")
        bg_canvas.create_window(440, 280, window=main_frame)

        ttk.Label(main_frame, text="Answer Evaluation System", font=("Times New Roman", 18, "bold")).pack(pady=10)

        subject_frame = ttk.Frame(main_frame)
        subject_frame.pack(fill=tk.X, pady=10)

        ttk.Label(subject_frame, text="Select Subject:", font=("Times New Roman", 12)).pack(side=tk.LEFT, padx=5)
        selected_subject_var = StringVar(value=subjects[0] if subjects else "Select Subject")
        subject_dropdown = ttk.Combobox(subject_frame, textvariable=selected_subject_var, values=subjects, state="readonly", width=20)
        subject_dropdown.pack(side=tk.LEFT, padx=5, expand=True, fill=tk.X)

        folder_frame = ttk.Frame(main_frame)
        folder_frame.pack(fill=tk.X, pady=10)

        ttk.Label(folder_frame, text="Answer Sheets Folder:", font=("Times New Roman", 12)).pack(anchor=tk.W, pady=5)

        folder_entry_frame = ttk.Frame(folder_frame)
        folder_entry_frame.pack(fill=tk.X)

        answer_folder_entry = ttk.Entry(folder_entry_frame, textvariable=answer_folder_var, width=30)
        answer_folder_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))

        ttk.Button(folder_entry_frame, text="Browse", command=lambda: answer_folder_var.set(filedialog.askdirectory())).pack(side=tk.RIGHT)

        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=20)

        def evaluate():
            subject = selected_subject_var.get()
            if subject == "Select Subject":
                messagebox.showerror("Error", "Please select a subject.")
                return

            status_var.set("Evaluating... Please wait")
            root.update_idletasks()

            results_data = evaluate_all_students(subject)

            status_var.set("Ready")

            if results_data:
                display_results(root, results_data)
            else:
                messagebox.showerror("Error", "Evaluation failed.")

        ttk.Button(button_frame, text="Evaluate", command=evaluate, style="TButton").pack(pady=10, fill=tk.X)

        ttk.Button(main_frame, text="Logout", command=lambda: [root.destroy(), main_app()]).pack(side=tk.BOTTOM, pady=10)

        help_frame = ttk.Frame(main_frame)
        help_frame.pack(fill=tk.X, side=tk.BOTTOM, pady=5)

        def show_help():
            help_window = Toplevel(root)
            help_window.title("Help")
            help_window.geometry("850x600")
            help_window.resizable(False, False)

            help_canvas = set_background(help_window, "bg.jpg")

            help_text = Text(help_window, wrap='word', font=("Times New Roman", 15))
           
            help_canvas.create_window(420, 300, window=help_text)
            
            help_content = """



                                              Answer Evaluation System - Help


                                    1. Login with your username and password
                                    2. Select a subject from the dropdown menu
                                    3. Click "Browse" to select the folder containing student
                                        answer sheets
                                    4. Click "Evaluate" to process all answer sheets
                                    5. View results in the detailed or summary tabs
                                    6. Export results using the export buttons
                                    

                                    Note: Answer sheets should be named with the student's name
                                    followed by "_AnswerSheet.docx"
                    """

            help_text.insert(tk.END, help_content)
            help_text.configure(state='disabled')

        ttk.Button(help_frame, text="Help", command=show_help).pack(side=tk.RIGHT)

    # LOGIN SCREEN
    bg_canvas = set_background(root, "bg1.jpg")

    login_frame = ttk.Frame(root, padding="20 20 20 20")
    bg_canvas.create_window(425, 220, window=login_frame)

    ttk.Label(login_frame, text="ANSWER  EVALUATION  SYSTEM", font=("Times New Roman", 17, "bold")).pack(pady=20)
    ttk.Label(login_frame, text="Login", font=("Times New Roman", 15, "bold")).pack(pady=20)

    username_frame = ttk.Frame(login_frame)
    username_frame.pack(fill=tk.X, pady=5)
    ttk.Label(username_frame, text="Username:", width=10, font=("Times New Roman", 14, "bold")).pack(side=tk.LEFT, padx=5)
    ttk.Entry(username_frame, textvariable=username_var, width=20, font=("Times New Roman", 14)).pack(side=tk.LEFT, padx=5, expand=True, fill=tk.X)

    password_frame = ttk.Frame(login_frame)
    password_frame.pack(fill=tk.X, pady=5)
    ttk.Label(password_frame, text="Password:", width=10, font=("Times New Roman", 14, "bold")).pack(side=tk.LEFT, padx=5)
    ttk.Entry(password_frame, textvariable=password_var, show="*", width=20, font=("Times New Roman", 14, "bold")).pack(side=tk.LEFT, padx=5, expand=True, fill=tk.X)

    ttk.Button(login_frame, text="Login", command=handle_login, style="TButton").pack(pady=20)

    status_var = StringVar(value="Ready")
    status_bar = ttk.Label(root, textvariable=status_var, relief=tk.SUNKEN, anchor=tk.W)
    status_bar.pack(side=tk.BOTTOM, fill=tk.X)

    root.mainloop()

if __name__ == "__main__":
    main_app()
