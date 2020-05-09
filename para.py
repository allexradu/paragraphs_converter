import excel

# assigning the fist column of the workbook to the cells list
cells = excel.read_excel_first_column()

for each_cell in cells:
    # splitting the text inside the multiple-line-cell into a list of individual lines
    all_the_lines = each_cell.splitlines()

    # adding paragraphs to each line
    for i in range(len(all_the_lines) - 1):
        all_the_lines[i] = '<p>' + all_the_lines[i] + '</p>\n'

    # joining all the lines into one multi-line string
    excel.descriptions.append(''.join(all_the_lines))

# adding all the data to the excel workbook
excel.add_data_to_workbook()
