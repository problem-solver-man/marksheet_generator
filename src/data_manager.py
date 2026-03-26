from openpyxl import load_workbook


# Generator function to provide single worksheet at a time
def load_data(excel_path, selected_class):
    wb = load_workbook(excel_path, read_only=True, data_only=True)

    if selected_class == "All Class":
        for ws in wb.worksheets:
            yield ws
    else:
        yield wb[selected_class]


# Function to prepare marksheet values for all students of a class
def prepare_data(ws) -> list:
    wslist = []     # Worksheet-as-a-list
    for row in ws.iter_rows(min_row=2, values_only=True):
        row = list(row)
        # Handle empty values
        for i in range(len(row)):
            if row[i] is None:
                row[i] = 0

        wslist.append([row[:2], row[2:9], row[9:16], row[16:23],
                       row[23:30], row[30:37], row[37:44], row[44:]])

    # Add different total marks in each student list
    for stud in wslist:
        # Total secured subjective marks (tssm)
        tssm = [sum(x for x in stud[i][:2]
                    if isinstance(x, (int, float))) for i in range(1, 7)]
        tssm.append(stud[7][0])

        # Number of failed subjects
        failed_subjects = sum(1 for marks in tssm if marks < 30)

        # Total secured marks (tsm)
        tsm = [sum(x for x in stud[i] 
                   if isinstance(x, (int, float))) for i in range(1, 8)]

        # Calculate total marks, percentage, and result (Pass or Fail)
        sum_tssm = sum(tssm)
        sum_tsm = sum(tsm)
        result = "Pass" if failed_subjects == 0 else "Fail"
        percentage = f"{sum_tsm / 7:.2f}%"

        # Append [tssm], [tsm], and [sum of tssm, tsm] into wslist
        stud.append(tssm)
        stud.append(tsm)
        stud.append([failed_subjects, sum_tssm, sum_tsm, result, percentage])

    wslist.sort(key=lambda x: (x[10][0], -x[10][2], -x[10][1]))
    return wslist


# Function to convert all values to type str
def polish_data(student) -> list:
    tssm = student[8]

    for index, value in enumerate(tssm):
        if isinstance(value, (int, float)) and value < 30:
            tssm[index] = "*" + str(round(value))

    for item in student:
        for index, value in enumerate(item):
            if isinstance(value, float):
                item[index] = str(round(value))
            elif not isinstance(value, str):
                item[index] = str(value)

    return student


# Function to return the rank of a student
def get_rank(index):
    if 10 <= index % 100 <= 13:
        suffix = "th"
    else:
        suffix = {1: "st", 2: "nd", 3: "rd"}.get(index % 10, "th")
        
    return f"{index}<sup>{suffix}</sup>"

