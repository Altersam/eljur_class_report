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
            dates, student_grades = extract_grades_from_page(page)

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


def extract_grades_from_page(page):
    blocks = page.get_text("dict")["blocks"]

    month_map = {
        "Январь": "01", "Февраль": "02", "Март": "03", "Апрель": "04",
        "Май": "05", "Июнь": "06", "Сентябрь": "09", "Октябрь": "10",
        "Ноябрь": "11", "Декабрь": "12",
    }

    # Collect all spans with their coordinates
    spans = []
    for block in blocks:
        if "lines" in block:
            for line in block["lines"]:
                for span in line["spans"]:
                    text = span["text"].strip()
                    if text:
                        x0, y0, x1, y1 = span["bbox"]
                        spans.append({"text": text, "x0": x0, "y0": y0, "x1": x1, "y1": y1})

    # Find month headers (y around 65-80)
    month_spans = [s for s in spans if s["text"] in month_map and 65 < s["y0"] < 80]
    month_spans.sort(key=lambda s: s["x0"])

    # Find ПР/КР/СР boundaries (y around 95-110)
    boundary_spans = [s for s in spans if s["text"] in ("ПР", "КР", "СР") and 95 < s["y0"] < 110]
    boundary_spans.sort(key=lambda s: s["x0"])

    # Date spans are numbers with y around 85-100
    date_spans = [s for s in spans if re.match(r'^\d{1,2}$', s["text"]) and 85 < s["y0"] < 100]
    date_spans.sort(key=lambda s: s["x0"])

    # Build zones: each zone is (start_x, end_x, month_name)
    # A zone starts at a month header or after a boundary, and ends at the next month/boundary
    zones = []
    month_names = list(month_map.keys())

    for i, ms in enumerate(month_spans):
        # Find boundaries after this month
        boundaries_after = [b for b in boundary_spans if b["x0"] > ms["x0"]]

        # Zone for current month: from month_x to first boundary after it
        if boundaries_after:
            zones.append((ms["x0"], boundaries_after[0]["x0"], ms["text"]))
            # Zone after boundary = next month
            next_boundary_x = boundaries_after[0]["x0"]
            # Find end of this zone (next month or next boundary)
            if i + 1 < len(month_spans):
                zone_end = month_spans[i + 1]["x0"]
            else:
                # After last boundary, use next boundary or infinity
                if len(boundaries_after) > 1:
                    zone_end = boundaries_after[1]["x0"]
                else:
                    zone_end = float('inf')

            # Determine next month name
            if ms["text"] in month_names:
                idx = month_names.index(ms["text"])
                next_month = month_names[idx + 1] if idx + 1 < len(month_names) else ms["text"]
            else:
                next_month = ms["text"]

            zones.append((next_boundary_x, zone_end, next_month))
        else:
            # No boundary, zone extends to next month
            zone_end = month_spans[i + 1]["x0"] if i + 1 < len(month_spans) else float('inf')
            zones.append((ms["x0"], zone_end, ms["text"]))

    def get_month_for_date(date_x):
        for start, end, month in zones:
            if start <= date_x < end:
                return month_map[month]
        # Fallback
        return month_map[month_spans[-1]["text"]] if month_spans else "01"

    all_dates = []
    for ds in date_spans:
        month = get_month_for_date(ds["x0"])
        all_dates.append(f"{ds['text'].zfill(2)}.{month}")

    num_dates = len(all_dates)
    if num_dates == 0:
        return [], {}

    # Get date column x-centers
    date_x_centers = [ds["x0"] + (ds["x1"] - ds["x0"]) / 2 for ds in date_spans]

    # Find student rows and their grades
    student_grades = {}

    # Find all student name spans (number followed by name on same Y)
    # Use a wider Y tolerance for name matching
    name_spans = []
    for s in spans:
        if re.match(r'^\d+$', s["text"]) and 100 < s["y0"] < 600:
            for s2 in spans:
                if (abs(s2["y0"] - s["y0"]) < 5 and
                        s2["x0"] > s["x1"] and
                        re.match(r'^[А-ЯЁ][а-яё]+', s2["text"])):
                    name_spans.append({"num": s, "name": s2, "y": s["y0"]})
                    break

    for ns in name_spans:
        student_y = ns["y"]
        student_name = ns["name"]["text"]

        # Initialize grades array with None
        grades = [None] * num_dates

        # Find grade spans on same Y row (within tolerance of 8)
        for s in spans:
            if abs(s["y0"] - student_y) < 8 and s["x0"] > 170:
                s_center = s["x0"] + (s["x1"] - s["x0"]) / 2
                best_col = -1
                best_dist = 15
                for col_idx, cx in enumerate(date_x_centers):
                    dist = abs(s_center - cx)
                    if dist < best_dist:
                        best_dist = dist
                        best_col = col_idx

                if best_col >= 0:
                    txt = s["text"]
                    if txt == "Н":
                        val = 0
                    elif txt.isdigit():
                        val = int(txt)
                        if not (1 <= val <= 5):
                            val = None
                    else:
                        val = None

                    if val is not None:
                        if grades[best_col] is None or (grades[best_col] == 0 and val > 0):
                            grades[best_col] = val

        # Fill None with 0
        final_grades = [g if g is not None else 0 for g in grades]
        student_grades[student_name] = final_grades

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
