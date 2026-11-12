import pandas as pd
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime

# ==================================
# SCHOOL INFORMATION
# ==================================

school_name = "ABDULLAHI DEMO SECONDARY SCHOOL"
term = "FIRST TERM"
class_name = "SS 1 SCIENCE"
session = "2026/2027"
class_teacher = "Mr. Abdullahi Umar Ismail"

# ==================================
# LOAD DATA
# ==================================

df = pd.read_csv("students.csv")

subjects = [
    "Mathematics",
    "English",
    "Physics",
    "Chemistry",
    "Biology"
]

# ==================================
# CALCULATIONS
# ==================================

df["Total"] = df[subjects].sum(axis=1)
df["Average"] = df["Total"] / len(subjects)

# ==================================
# GRADES
# ==================================

def grade(avg):
    if avg >= 80:
        return "A"
    elif avg >= 70:
        return "B"
    elif avg >= 60:
        return "C"
    elif avg >= 50:
        return "D"
    else:
        return "F"

# ==================================
# REMARKS
# ==================================

def remark(avg):
    if avg >= 80:
        return "Excellent"
    elif avg >= 70:
        return "Very Good"
    elif avg >= 60:
        return "Good"
    elif avg >= 50:
        return "Fair"
    else:
        return "Fail"

df["Grade"] = df["Average"].apply(grade)
df["Remark"] = df["Average"].apply(remark)

# ==================================
# POSITIONS
# ==================================

def position_suffix(position):
    if position == 1:
        return "1st"
    elif position == 2:
        return "2nd"
    elif position == 3:
        return "3rd"
    else:
        return f"{position}th"

df["Position"] = (
    df["Average"]
    .rank(ascending=False, method="min")
    .astype(int)
)

df["Position"] = df["Position"].apply(position_suffix)

df = df.sort_values(
    by="Average",
    ascending=False
)

# ==================================
# DISPLAY RESULTS
# ==================================

print("\nSTUDENT RESULT SHEET\n")

print(
    df[
        [
            "Name",
            "Total",
            "Average",
            "Grade",
            "Remark",
            "Position"
        ]
    ]
)

# ==================================
# SAVE CSV
# ==================================

df.to_csv(
    "result_sheet.csv",
    index=False
)

# ==================================
# CLASS STATISTICS
# ==================================

best_student = df.iloc[0]["Name"]

class_average = round(
    df["Average"].mean(),
    2
)

total_students = len(df)

pass_rate = round(
    (
        len(df[df["Average"] >= 50])
        / total_students
    ) * 100,
    2
)

# ==================================
# EXCEL REPORT
# ==================================

with pd.ExcelWriter(
    "result_sheet.xlsx",
    engine="openpyxl"
) as writer:

    df.to_excel(
        writer,
        sheet_name="Results",
        startrow=7,
        index=False
    )

    worksheet = writer.sheets["Results"]

    # ==================================
    # MERGED SCHOOL HEADER
    # ==================================

    worksheet.merge_cells("A1:K1")
    worksheet.merge_cells("A2:K2")
    worksheet.merge_cells("A3:K3")
    worksheet.merge_cells("A4:K4")
    worksheet.merge_cells("A5:K5")

    worksheet["A1"] = school_name
    worksheet["A2"] = f"{term} EXAMINATION RESULT"
    worksheet["A3"] = class_name
    worksheet["A4"] = session
    worksheet["A5"] = f"Class Teacher: {class_teacher}"

    # ==================================
    # GENERATED DATE
    # ==================================

    worksheet["M7"] = "Generated On"
    worksheet["N7"] = datetime.now().strftime("%d/%m/%Y")

    # ==================================
    # STATISTICS
    # ==================================

    worksheet["M1"] = "CLASS STATISTICS"

    worksheet["M2"] = "Best Student"
    worksheet["N2"] = best_student

    worksheet["M3"] = "Class Average"
    worksheet["N3"] = class_average

    worksheet["M4"] = "Total Students"
    worksheet["N4"] = total_students

    worksheet["M5"] = "Pass Rate (%)"
    worksheet["N5"] = pass_rate

    # ==================================
    # SIGNATURES
    # ==================================

    worksheet["A18"] = "____________________"
    worksheet["A19"] = "Class Teacher"

    worksheet["F18"] = "____________________"
    worksheet["F19"] = "Principal"

    worksheet["J18"] = "____________________"
    worksheet["J19"] = "Result Officer"

    # ==================================
    # STYLING
    # ==================================

    school_fill = PatternFill(
        start_color="1F4E78",
        end_color="1F4E78",
        fill_type="solid"
    )

    white_font = Font(
        bold=True,
        color="FFFFFF",
        size=14
    )

    header_font = Font(
        bold=True,
        color="FFFFFF"
    )

    blue_fill = PatternFill(
        start_color="4472C4",
        end_color="4472C4",
        fill_type="solid"
    )

    center = Alignment(
        horizontal="center",
        vertical="center"
    )

    thin = Side(style="thin")

    border = Border(
        left=thin,
        right=thin,
        top=thin,
        bottom=thin
    )

    # School Header Formatting
    for cell in ["A1", "A2", "A3", "A4", "A5"]:
        worksheet[cell].font = white_font
        worksheet[cell].fill = school_fill
        worksheet[cell].alignment = center

    # Statistics Formatting
    worksheet["M1"].font = header_font
    worksheet["M1"].fill = blue_fill
    worksheet["M1"].alignment = center

    for row in range(1, 6):
        worksheet[f"M{row}"].border = border
        worksheet[f"N{row}"].border = border

    # Table Header Formatting
    for cell in worksheet[8]:
        cell.font = header_font
        cell.fill = blue_fill
        cell.alignment = center
        cell.border = border

    # Data Borders
    for row in worksheet.iter_rows(
        min_row=9,
        max_row=8 + len(df),
        min_col=1,
        max_col=11
    ):
        for cell in row:
            cell.border = border

    # Auto Column Width
    for column_cells in worksheet.columns:

        max_length = 0

        column_letter = get_column_letter(
            column_cells[0].column
        )

        for cell in column_cells:
            try:
                if cell.value:
                    max_length = max(
                        max_length,
                        len(str(cell.value))
                    )
            except:
                pass

        worksheet.column_dimensions[
            column_letter
        ].width = max_length + 5

print("\nResult sheet saved as result_sheet.csv")
print("Excel result saved as result_sheet.xlsx")