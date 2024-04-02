import openpyxl
import os
from colorist import red, effect_underline, Color, effect_bold

# dictionary where the attendance files at
directory = "D:\\astro\\Astro\\uni tutoring\\M110\\attendance"

# getting week number to record attendance for the corresponding week
target_week = int(input("Enter week number: "))

# using this list for displaying purposes once the students is recorded as attended
# it is also used to write the data to a temp file in case of corruption
# it stores a tuple(studentId, sectionNumber, StudentName)
recorded_students = []

while True:

    target_id = int(input("Enter student id: "))

    if target_id == -1:
        break

    # using this counter to indicate that the student id wasn't found in
    # 4 excell files if this counter value is 4
    notFoundCounter = 0

    # looping through files within the directory and opening only xlsx files
    for filename in os.listdir(directory):
        if filename.endswith("xlsx"):
            # getting different information about the file first
            # all the section except section 62 has different format
            # so there are different values for section 62 specifically
            isSection62 = '62' in filename
            id_row = 10 if isSection62 else 6
            id_column = 3
            current_week = target_week - 1 if isSection62 else target_week
            week_column = id_column + 3 + current_week if isSection62 else id_column + 3 + current_week
            name_column = id_column + 1
            section_number = filename[10:13]

            # opening the file and activating the current sheet
            wb = openpyxl.load_workbook(os.path.join(directory, filename))
            sheet = wb.active

            # this indicator used to exit the loops once student's id is found and marked attended
            # it is useful to immediately save the changes once the id found to avoid redundancy
            changeMade = False

            # looping through each record within the table
            # starting from the rows and columns where student ids are found
            # again to avoid redundancy
            for record in sheet.iter_rows(min_row=id_row, min_col=id_column, max_col=id_column):

                # looping through each cell
                # and comparing its value with the target student's id
                for cell in record:
                    # once found we get the current row and column to move from there
                    # to get more information like the name
                    if cell.value == target_id:
                        row = cell.row
                        column = cell.column

                        # marking the student as attended for the specific week
                        sheet.cell(row=row, column=week_column, value="p")
                        # getting students name
                        name = sheet.cell(row, name_column).value
                        # displaying to console that the operation was successful with bold text
                        effect_bold(
                            f"Name: {name}\nSection: {section_number}\nhas been successfully marked as attended for "
                            f"week {target_week}\n\n")
                        # here we set the changeMade to true, so we can immediately stop looping
                        # and ask for the next student's id
                        changeMade = True
                        # adding the record to list for the reasons set earlier
                        recorded_students.append((target_id, section_number, name))
                        break
                if changeMade:
                    break
            if changeMade:
                # saving the file and resetting the counter, so it doesn't reach 4
                # indicating that student's id was found and changes has been made
                wb.save(os.path.join(directory, filename))
                notFoundCounter = 0
                break
            else:
                # if the id is not found within a file then it will be incremented
                # until it reaches 4 indicating that the id is not within any file
                # then we counter
                notFoundCounter += 1

        # if id is not found within the 4 files
        if notFoundCounter == 4:
            red(f"Student with id {target_id} is not found")
            notFoundCounter = 0


def formatRecord(currentIndex):
    """Formatting information saved within the recorded students list"""
    return f"{currentIndex + 1}) {recorded_students[currentIndex][0]} - {recorded_students[currentIndex][1]} -" \
           f" {recorded_students[currentIndex][2]} "


def writeRecordToFile():
    """Saving information recorded in recorded_students list into a file in case of corruption & checking"""
    with open(os.path.join(directory + "\\temp", f"Week {target_week}.txt"), 'w') as f:
        for index in range(len(recorded_students)):
            f.write(formatRecord(index) + "\n")


# displaying how many students where recorded and saving info into a file
if recorded_students:
    writeRecordToFile()
    effect_underline("Recorded students", Color.GREEN)
    for i in range(len(recorded_students)):
        record = formatRecord(i)
        effect_underline(record, Color.GREEN)
    effect_underline(f"Recorded {len(recorded_students)} students successfully", Color.GREEN)
