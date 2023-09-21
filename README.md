# Python_Excel
Extracting specific columns to be converted into a CSV file for sharepoint migration 
import re
from openpyxl import load_workbook
from datetime import datetime
from pathlib import Path
from csv import DictWriter

FIELD_NAME_MAP = {
    re.compile(r"\bname\b(?! of interviewer)"): "Name",
    re.compile(r"\bdate of interview\b"): "Date of Interview",
    re.compile(r"\bemail\b"): "Email",
    re.compile(
        r"\bclient\b(?!.*\b(?:role|category|location|requirements|description)\b)"
    ): "Client",
    re.compile(r"\brole category\b"): "Role Category",
    re.compile(r"\blocation of role\b"): "Location of Role",
    re.compile(r"\bround\b"): "Round",
    re.compile(r"\bduration\b"): "Duration",
    re.compile(
        r"\b(interviewer|name of interviewer|interviewer \/ hiring manager)\b"
    ): "Hiring Manager",
    re.compile(r"\baccount manager\b"): "Account Manager",
    re.compile(r"\bwhat went well\b"): "Self Reflection - What went well?",
    re.compile(r"\bwhat did not go well\b"): "Self Reflection - What did not go well?",
    re.compile(
        r"\bwhat could you have done better\b"
    ): "Self Reflection - What could you have done better?",
    re.compile(r"\bother comments\b"): "Self Reflection - Other comments?",
}


def get_field_name(cell_value):
    cell_value = cell_value.lower()
    for pattern, field_name in FIELD_NAME_MAP.items():
        if pattern.search(cell_value):
            return field_name
    return None


def extract_data_from_excel(file_path):
    wb = load_workbook(file_path, data_only=True)
    ws = wb.active

    data = {
        "Name": "",
        "Date of Interview": "",
        "Email": "",
        "Client": "",
        "Role Category": "",
        "Location of Role": "",
        "Round": "",
        "Duration": "",
        "Hiring Manager": "",
        "Account Manager": "",
        "Questions": [],
        "Categories": [],
        "Ratings": [],
        "Self Reflection - What went well?": "",
        "Self Reflection - What did not go well?": "",
        "Self Reflection - What could you have done better?": "",
        "Self Reflection - Other comments?": "",
    }

    questions_start_idx = None
    self_reflection_start_idx = None
    email_domain_part = ""
    questions_header_skipped = False
    self_reflection_questions_captured = 0

    for idx, row in enumerate(ws.iter_rows(values_only=True), start=1):
        cell_value = row[0]
        if cell_value is None or self_reflection_questions_captured >= 4:
            continue

        cell_value_str = str(cell_value).strip().lower()

        if "key questions you were asked" in cell_value_str:
            questions_start_idx = idx + 1
        if "self reflection" in cell_value_str:
            self_reflection_start_idx = idx + 1

        field_name = get_field_name(cell_value_str)
        if field_name:
            if field_name.startswith("Self Reflection"):
                self_reflection_questions_captured += 1

            if field_name == "Client":
                data[field_name] = str(row[1]).split("\n")[0] if row[1] else ""
            elif field_name == "Email" and "@" not in str(row[1]):
                email_domain_part = row[2] if row[2] else ""
                data[field_name] = f"{row[1]}{email_domain_part}"
            elif field_name == "Date of Interview" and isinstance(row[1], datetime):
                data[field_name] = row[1].strftime("%d-%m-%Y")
            else:
                data[field_name] = row[1]

        if questions_start_idx and idx >= questions_start_idx:
            if self_reflection_start_idx and idx >= self_reflection_start_idx:
                continue
            if not cell_value or "self reflection" in cell_value_str:
                continue
            if not questions_header_skipped:
                if "question" in cell_value_str.lower():
                    questions_header_skipped = True
                    continue
            data["Questions"].append(row[0])
            data["Categories"].append(row[2])
            data["Ratings"].append(row[3])

    return data


def save_data_to_csv(data_list, output_csv_path):
    columns = [
        "Name",
        "Date of Interview",
        "Email",
        "Client",
        "Role Category",
        "Location of Role",
        "Round",
        "Duration",
        "Hiring Manager",
        "Account Manager",
        "Questions",
        "Categories",
        "Ratings",
        "Self Reflection - What went well?",
        "Self Reflection - What did not go well?",
        "Self Reflection - What could you have done better?",
        "Self Reflection - Other comments?",
    ]

    with open(output_csv_path, "w", newline="", encoding="utf-8") as f:
        writer = DictWriter(f, fieldnames=columns)
        writer.writeheader()
        writer.writerows(data_list)


def extract_data_from_directory(directory_path, output_csv_path):
    data_list = []
    for file_path in Path(directory_path).rglob("*.xls*"):
        try:
            data_list.append(extract_data_from_excel(file_path))
        except Exception as e:
            print(f"Could not process file {file_path}: {e}")

    save_data_to_csv(data_list, output_csv_path)


if __name__ == "__main__":
    directory_path = r"c:\Users\61490\Desktop\FDM\Power Apps Project\OneDrive_1_14-09-2023"
    output_csv_path = (
        r"c:\Users\61490\Desktop\CSVfile.csv"
    )

    extract_data_from_directory(directory_path, output_csv_path)
