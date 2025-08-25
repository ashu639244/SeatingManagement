# Exam Seating Arrangement and Attendance System

## Tasks

### Task 1: Seating Arrangement for Exam
Design Python code to generate seating arrangement for exams.

- **Inputs (CSV with `;` separator):**
  - **ip_1**: roll, sem number, and registered courses (ignore schedule_sem).
  - **ip_2**: exam timetable.
  - **ip_3**: room capacity (can be updated over time; keep dynamic).
  - **ip_4**: roll–student name mapping.

- **Outputs:**
  - **op_1** and **op_2** (both in CSV and Excel format).

- **Constraints:**
  1. Fill large courses first.
  2. Allocate rooms in **Block 9 first**, then **LT**.
  3. Do not split an exam across Block 9 and LT.
  4. Keep a buffer of 5 students per classroom (make this configurable).
  5. Two modes:
     - **Dense**: Full capacity of room can be used for one course.
     - **Sparse**: Max half of room capacity can be used for one course.
  6. Optimize number of classrooms used and maintain room locality (avoid far-separated allocations like 101 and 401 unless necessary).

---

### Task 2: Class-wise Attendance Sheet
Design Python code to generate attendance sheet for each exam.

- **Details:**
  - Attendance sheet must contain **roll, name, signature**.
  - At the bottom: **5 rows × 2 columns blank** for invigilator and TA signatures.
  - Output file format:
    ```
    dd_mm_yyyy_sub_code_room_num_(morning/evening).xlsx
    ```
  - Use **ip_4** for roll–name mapping.

---

## File References
- [Input File (Google Sheets)](https://docs.google.com/spreadsheets/d/1yt3eDmftPLrKo5REC23qLISc4xNTTrmN6i5dIszEkvg/edit?usp=sharing)

---

## Deadlines
- **Group info submission:** 06/10/2024
- **Progress check:** 23/10/2024 (in lab)
- **Final check:** 13/11/2024

---

## Important Notes
- Inputs are **CSV files with `;` as separator**.
- Optimize seating allocation for room usage and locality.
- Plagiarism check will be applied. Identical submissions will get **zero marks**.
- This project carries significant grading weightage.

