import os
import docx
import PyPDF2
import pdfplumber
import csv


# 读取 Word 文件并统计字数
def count_words_in_docx(file_path):
    doc = docx.Document(file_path)
    word_count = 0
    for para in doc.paragraphs:
        word_count += len(para.text)
    return word_count


# 读取 PDF 文件并统计字数
def count_words_in_pdf(file_path):
    word_count = 0
    with open(file_path, 'rb') as file:
        pdf = PyPDF2.PdfReader(file)
        for page in pdf.pages:
            word_count += len(page.extract_text())
    return word_count


# 读取 PDF 使用 pdfplumber（处理更复杂的布局）
def count_words_in_pdf_plumber(file_path):
    word_count = 0
    with pdfplumber.open(file_path) as pdf:
        for page in pdf.pages:
            word_count += len(page.extract_text())
    return word_count


# 获取文件夹内所有文件，筛选出 Word 和 PDF 文件
def get_files_in_directory(directory_path):
    files = []
    for filename in os.listdir(directory_path):
        file_path = os.path.join(directory_path, filename)
        if filename.endswith('.docx') or filename.endswith('.pdf'):
            files.append(file_path)
    return files


# 统计所有文件的字数
def count_words_in_files(directory_path):
    files = get_files_in_directory(directory_path)
    word_counts = []

    for file_path in files:
        file_name = os.path.basename(file_path)
        if file_path.endswith('.docx'):
            word_count = count_words_in_docx(file_path)
        elif file_path.endswith('.pdf'):
            try:
                word_count = count_words_in_pdf_plumber(file_path)  # 使用pdfplumber来更好地处理布局
            except Exception as e:
                print(f"Error processing {file_name}: {e}")
                word_count = 0

        word_counts.append((file_name, word_count))

    return word_counts


# 将统计结果保存为 CSV 文件
def save_word_counts_to_csv(word_counts, output_path):
    with open(output_path, mode='w', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        writer.writerow(['文件名', '字数统计'])
        for file_name, word_count in word_counts:
            writer.writerow([file_name, word_count])


# 主程序

input_folder = 'after'  # 替换为你的文件夹路径
output_file = '统计结果.csv'  # 输出文件路径
word_counts = count_words_in_files(input_folder)
save_word_counts_to_csv(word_counts, output_file)
print(f"统计结果已保存至 {output_file}")



