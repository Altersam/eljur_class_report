import pymupdf
import json
import re
import os
import sys
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill


def resource_path(relative_path):
    """Get absolute path to resource, works for dev and for PyInstaller"""
    if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)


def parse_pdf(pdf_path):
    doc = pymupdf.open(pdf_path)

    students = []
    subjects = []
    teachers = {}
    details = {}
    all_dates = {}

    i = 0
    while i < doc.page_count:
        page = doc[i]
        text = page.get_text()

        if "Наименование предмета" in text:
            subj_name = extract_subject_name(text)
            dates, student_grades = extract_grades_from_page(text)

            if subj_name:
                if subj_name not in subjects:
                    subjects.append(subj_name)
                all_dates[subj_name] = dates

                for name, grades in student_grades.items():
                    if name not in students:
                        students.append(name)
                    if name not in details:
                        details[name] = {}

                    valid = [g for g in grades if g > 0]
                    absences = sum(1 for g in grades if g == 0)
                    avg_val = round(sum(valid) / len(valid), 2) if valid else 0.0

                    entries = []
                    for d, g in zip(dates, grades):
                        if g > 0:
                            entries.append({"date": d, "grade": g})

                    if subj_name in details[name]:
                        existing = details[name][subj_name]
                        existing["grades"].extend(valid)
                        existing["entries"].extend(entries)
                        existing["count"] = len(existing["grades"])
                        existing["absences"] += absences
                        g2 = existing["grades"]
                        existing["avg"] = round(sum(g2) / len(g2), 2) if g2 else 0.0
                        existing["predicted"] = predict(existing["avg"])
                    else:
                        details[name][subj_name] = {
                            "avg": avg_val,
                            "predicted": predict(avg_val),
                            "count": len(valid),
                            "absences": absences,
                            "grades": valid,
                            "teacher": "",
                            "entries": entries,
                        }

            i += 1
            continue

        if "Фамилия, имя, отчество учителя" in text:
            teacher_name = extract_teacher_name(text)
            if teacher_name and subjects:
                last_subj = subjects[-1]
                teachers[last_subj] = teacher_name
                for name in details:
                    if last_subj in details[name]:
                        details[name][last_subj]["teacher"] = teacher_name
            i += 1
            continue

        i += 1

    # Build averages
    averages = {}
    for name in students:
        averages[name] = {}
        if name in details:
            for subj in subjects:
                if subj in details[name]:
                    averages[name][subj] = details[name][subj]["avg"]

    return {
        "students": students,
        "subjects": subjects,
        "teachers": teachers,
        "averages": averages,
        "details": details,
    }


def extract_subject_name(text):
    lines = text.split("\n")
    found = False
    parts = []
    for line in lines:
        s = line.strip()
        if "Наименование предмета" in s:
            found = True
            continue
        if found:
            if s in ("*", "ЧИСЛО", "Список", "обучающихся", "Месяц", "Январь",
                     "Февраль", "Март", "Апрель", "Май", "Июнь", "Сентябрь",
                     "Октябрь", "Ноябрь", "Декабрь", "ПР", "КР", "СР",
                     "ДЗ", "Д/З", "Число", "и месяц"):
                continue
            if re.match(r'^\d+$', s):
                continue
            if s and not s.startswith("№"):
                parts.append(s)
                if len(parts) >= 2:
                    break
    if not parts:
        return None
    name = " ".join(parts).strip()
    name = re.sub(r'\s+', ' ', name)
    return name


def extract_teacher_name(text):
    lines = text.split("\n")
    for i, line in enumerate(lines):
        if "Фамилия, имя, отчество учителя" in line:
            if i + 1 < len(lines):
                return lines[i + 1].strip()
    return None


def extract_grades_from_page(text):
    lines = text.split("\n")

    date_start = None
    for idx, line in enumerate(lines):
        s = line.strip()
        if re.search(r'\d{1,2}\s+\d{1,2}', s) and idx > 5:
            date_start = idx
            break

    if date_start is None:
        return [], {}

    month_map = {
        "Январь": "01", "Февраль": "02", "Март": "03", "Апрель": "04",
        "Май": "05", "Июнь": "06", "Сентябрь": "09", "Октябрь": "10",
        "Ноябрь": "11", "Декабрь": "12",
    }

    # Find month headers before date_start
    months_found = []
    for idx in range(max(0, date_start - 10), date_start):
        s = lines[idx].strip()
        for m in month_map:
            if s == m:
                months_found.append((idx, m))

    first_month = months_found[0][1] if months_found else "Январь"
    second_month = months_found[1][1] if len(months_found) > 1 else None

    # Collect dates, assigning months based on ПР/КР boundaries
    all_dates = []
    current_month = first_month
    for idx in range(date_start, min(date_start + 5, len(lines))):
        s = lines[idx].strip()
        if s in ("ПР", "КР", "СР"):
            if second_month:
                current_month = second_month
            continue
        dates = re.findall(r'\d{1,2}', s)
        for d in dates:
            all_dates.append(f"{d.zfill(2)}.{month_map.get(current_month, '01')}")

    num_dates = len(all_dates)
    if num_dates == 0:
        return [], {}

    student_grades = {}
    idx = date_start + 5
    while idx < len(lines):
        line = lines[idx].strip()
        m = re.match(r'^(\d+)\s+([А-ЯЁ][а-яё]+\s+[А-ЯЁА-ЯЁ][а-яёА-ЯЁ]+)', line)
        if m:
            name = m.group(2)
            grade_tokens = []
            j = idx + 1
            while len(grade_tokens) < num_dates and j < len(lines):
                gl = lines[j].strip()
                if re.match(r'^\d+\s+[А-ЯЁ][а-яё]+\s+[А-ЯЁ]', gl):
                    break
                tokens = re.findall(r'[Н0-5]', gl)
                for t in tokens:
                    if len(grade_tokens) < num_dates:
                        if t == 'Н':
                            grade_tokens.append(0)
                        elif t.isdigit():
                            grade_tokens.append(int(t))
                j += 1
            if grade_tokens:
                student_grades[name] = grade_tokens
            idx = j
        else:
            idx += 1

    return all_dates, student_grades


def predict(avg):
    if avg >= 4.6:
        return 5
    elif avg >= 3.6:
        return 4
    elif avg >= 2.6:
        return 3
    else:
        return 2


def clean_subj(name):
    name = re.sub(r'\s*№\s*п/п\s*$', '', name).strip()
    return name


def generate_excel(data, output_path):
    wb = Workbook()
    ws1 = wb.active
    ws1.title = "Сводка"

    hf = Font(name='Calibri', bold=True, color='FFFFFF', size=11)
    hf_fill = PatternFill(start_color='1A237E', end_color='1A237E', fill_type='solid')

    students = data["students"]
    subjects = data["subjects"]
    averages = data["averages"]

    ws1.cell(row=1, column=1, value="№").font = hf
    ws1.cell(row=1, column=1).fill = hf_fill
    ws1.cell(row=1, column=2, value="Ученик").font = hf
    ws1.cell(row=1, column=2).fill = hf_fill

    col = 3
    for subj in subjects:
        ws1.cell(row=1, column=col, value=clean_subj(subj)).font = hf
        ws1.cell(row=1, column=col).fill = hf_fill
        col += 1

    ws1.cell(row=1, column=col, value="Средний балл").font = hf
    ws1.cell(row=1, column=col).fill = hf_fill
    col += 1
    ws1.cell(row=1, column=col, value="Итог").font = hf
    ws1.cell(row=1, column=col).fill = hf_fill

    for row_idx, student in enumerate(students, 2):
        ws1.cell(row=row_idx, column=1, value=row_idx - 1)
        ws1.cell(row=row_idx, column=2, value=student)

        avgs = averages.get(student, {})
        vals = [v for v in avgs.values() if v > 0]
        overall = round(sum(vals) / len(vals), 2) if vals else 0

        col = 3
        for subj in subjects:
            val = avgs.get(subj, None)
            if val is not None and val > 0:
                ws1.cell(row=row_idx, column=col, value=val)
            else:
                ws1.cell(row=row_idx, column=col, value="-")
            col += 1

        ws1.cell(row=row_idx, column=col, value=overall)
        col += 1
        ws1.cell(row=row_idx, column=col, value=predict(overall))

    ws2 = wb.create_sheet("Детально")

    row = 1
    for student in students:
        ws2.cell(row=row, column=1, value=student).font = Font(bold=True, size=14, color='1A237E')
        row += 1

        ws2.cell(row=row, column=1, value="Предмет").font = hf
        ws2.cell(row=row, column=1).fill = hf_fill
        ws2.cell(row=row, column=2, value="Учитель").font = hf
        ws2.cell(row=row, column=2).fill = hf_fill
        ws2.cell(row=row, column=3, value="Оценки").font = hf
        ws2.cell(row=row, column=3).fill = hf_fill
        ws2.cell(row=row, column=4, value="НКИ").font = hf
        ws2.cell(row=row, column=4).fill = hf_fill
        ws2.cell(row=row, column=5, value="Среднее").font = hf
        ws2.cell(row=row, column=5).fill = hf_fill
        ws2.cell(row=row, column=6, value="Итог").font = hf
        ws2.cell(row=row, column=6).fill = hf_fill
        row += 1

        det = data["details"].get(student, {})
        for subj in subjects:
            if subj in det:
                d = det[subj]
                ws2.cell(row=row, column=1, value=clean_subj(subj))
                ws2.cell(row=row, column=2, value=d["teacher"])
                ws2.cell(row=row, column=3, value=", ".join(str(g) for g in d["grades"]))
                ws2.cell(row=row, column=4, value=d.get("absences", 0))
                ws2.cell(row=row, column=5, value=d["avg"])
                ws2.cell(row=row, column=6, value=d["predicted"])
                row += 1

        row += 1

    for cl in ['A', 'B', 'C', 'D', 'E', 'F']:
        ws2.column_dimensions[cl].width = 35

    wb.save(output_path)


def generate_json(data, output_path):
    clean_data = {
        "students": data["students"],
        "subjects": [clean_subj(s) for s in data["subjects"]],
        "teachers": {clean_subj(k): v for k, v in data["teachers"].items()},
        "averages": {},
        "details": {},
        "student_subjects": {},
    }

    for student in data["students"]:
        clean_data["averages"][student] = {}
        clean_data["student_subjects"][student] = []
        for subj in data["subjects"]:
            cs = clean_subj(subj)
            if subj in data["averages"].get(student, {}):
                clean_data["averages"][student][cs] = data["averages"][student][subj]
                clean_data["student_subjects"][student].append(cs)

        clean_data["details"][student] = {}
        for subj in data["subjects"]:
            cs = clean_subj(subj)
            if subj in data["details"].get(student, {}):
                d = data["details"][student][subj]
                clean_data["details"][student][cs] = {
                    "avg": d["avg"],
                    "predicted": d["predicted"],
                    "count": d["count"],
                    "absences": d.get("absences", 0),
                    "grades": d["grades"],
                    "teacher": d["teacher"],
                    "entries": d["entries"],
                }

    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(clean_data, f, ensure_ascii=False, indent=2)

    return clean_data


def generate_html(data, output_path):
    clean_data = generate_json(data, output_path.replace('.html', '.json'))

    data_json = json.dumps(clean_data, ensure_ascii=False)

    template_path = resource_path("dashboard.html")
    if os.path.exists(template_path):
        with open(template_path, "r", encoding="utf-8") as f:
            html = f.read()
        # Replace the embedded data
        old_data_start = html.find("const D = ")
        old_data_end = html.find(";", old_data_start)
        if old_data_start >= 0 and old_data_end >= 0:
            html = html[:old_data_start] + "const D = " + data_json + ";" + html[old_data_end + 1:]
    else:
        html = build_html_from_scratch(clean_data)

    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html)


def build_html_from_scratch(clean_data):
    data_json = json.dumps(clean_data, ensure_ascii=False)
    with open(resource_path("dashboard.html"), "r", encoding="utf-8") as f:
        html = f.read()
    old_data_start = html.find("const D = ")
    old_data_end = html.find(";", old_data_start)
    if old_data_start >= 0 and old_data_end >= 0:
        html = html[:old_data_start] + "const D = " + data_json + ";" + html[old_data_end + 1:]
    return html


def main():
    if len(sys.argv) < 2:
        print("Использование: report.py <путь_к_журналу.pdf>")
        print("Или перетащите PDF-файл на этот exe")
        input("Нажмите Enter для выхода...")
        sys.exit(1)

    pdf_path = sys.argv[1]

    if not os.path.exists(pdf_path):
        print(f"Файл не найден: {pdf_path}")
        input("Нажмите Enter для выхода...")
        sys.exit(1)

    if not pdf_path.lower().endswith('.pdf'):
        print("Укажите PDF-файл")
        input("Нажмите Enter для выхода...")
        sys.exit(1)

    base = os.path.splitext(pdf_path)[0]
    xlsx_path = base + '.xlsx'
    json_path = base + '.json'
    html_path = base + '.html'

    print(f"Обработка: {pdf_path}")
    data = parse_pdf(pdf_path)

    print(f"  Учеников: {len(data['students'])}")
    print(f"  Предметов: {len(data['subjects'])}")

    print(f"  Генерация Excel: {xlsx_path}")
    generate_excel(data, xlsx_path)

    print(f"  Генерация JSON: {json_path}")
    generate_json(data, json_path)

    print(f"  Генерация HTML: {html_path}")
    generate_html(data, html_path)

    print("Готово!")


if __name__ == "__main__":
    main()
