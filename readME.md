# 📘 Practical File PDF Generator  

This project automates the generation of **college practical file cover pages**.  

Instead of manually editing Word files for every subject, this script takes a **Word template** + **spreadsheet of subjects/teachers** + **student name input**, and generates **individual PDFs** and one **combined PDF** into a folder.  

---

## 📂 How It Works  
1. You prepare a `template.docx` (your cover page design with placeholders).  
2. You create a `subjects.xlsx` spreadsheet with **Subject_Name** and **Teacher_Name**.  
3. You run the script and enter your student name once.  
4. The script:  
   - Fills placeholders in the template for each subject.  
   - Converts them to PDFs.  
   - Deletes intermediate `.docx` files (only keeps PDFs).  
   - Copies your `index.pdf` (extra page) into the output folder.  
   - Generates a **combined_pages.pdf** that merges `index.pdf` + all subject PDFs.  

---

## 📑 Spreadsheet Format  

Your `subjects.xlsx` must have the following headers in the **first row**:  
(I already have the demo data, just adjust the names of teachers and subjects)  

| Subject_Name     | Teacher_Name    |
|------------------|----------------|
| Database System  | Dr. Sharma     |
| Networking       | Prof. Gurung   |
| AI               | Dr. Shrestha   |

---

## 📝 Word Template  

Leave the Word Template as is. Or if you want to change, just put:  

- `{{ Subject_Name }}` → for subject name  
- `{{ Teacher_Name }}` → for teacher’s name  
- `{{ Student_Name }}` → for your name  

---

## 🚀 How to Run  

1. Install Python and pip on your computer.  
2. Install dependencies:  

```bash
pip install pandas docxtpl docx2pdf PyPDF2 openpyxl
```

3. Run the program

```bash
python generate.py
```
4. You will be asked to enter your student name.
5. The files will be saved to /output folder along with index

---

