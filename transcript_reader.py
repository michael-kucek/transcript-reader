import string
from openpyxl import load_workbook
import time
import re
import sys
import csv
import math

smillis = int(round(time.time() * 1000))

# Create a global to hold the requests
smart_requests = []

# Simple progress bar output to IDE console
def update_progress(progress, text):
    barLength = 20 # Modify this to change the length of the progress bar
    status = ""
    if isinstance(progress, int):
        progress = float(progress)
    if not isinstance(progress, float):
        progress = 0
        status = "error: progress var must be float\r\n"
    if progress < 0:
        progress = 0
        status = "Halt...\r\n"
    if progress >= 1:
        progress = 1
        status = "Done...\r\n"
    block = int(round(barLength*progress))
    text = "\r" + text + ": [{0}] {1}% {2}".format( "#"*block + "-"*(barLength-block), round(progress*100), status)
    sys.stdout.write(text)
    sys.stdout.flush()

# Writes csvs
def csv_writer(data, path):
    with open(path, "w", newline="\n") as f:
        writer = csv.writer(f)
        for i in data:
            writer.writerow(i)

# uses openpyxl to read the transcripts
def load_transcripts(excel_file):
    print("loading " + excel_file)
    wb = load_workbook(filename = excel_file, read_only=True)
    print("loading complete")
    sheet_names = wb.get_sheet_names()
    converted_transcripts = []
    # read each of the sheets in the xlsx
    for sheet in sheet_names:
        ws = wb[sheet]
        sheet_length = ws._max_row
        i = 0
        for row in ws:
            i += 1
            temp_row = []
            for cell in row:
                if cell.value != None:
                    temp_row.append(cell.value)
            # find only the important iformation save it
            if len(temp_row) > 3 and temp_row[0] != 'Year' and temp_row[0] != "Fed Ethnicity & Race: "\
                    and not re.search('ALGEBRA I|BIOLOGY|ENGLISH I|ENGLISH II|US HISTORY', temp_row[0]):
                converted_transcripts.append(temp_row)
            update_progress(i/sheet_length, "Converting transcripts on " + sheet)
    return converted_transcripts

# Checks to see if there is a valid id and returns it
def get_id(row):
    return row[1].split()[2]
 
# Gets the name of the student
def get_name(row):
    return row[0]

# a lot of stuff to parse the credit and grade
def get_course_credit_and_grade(row):
    row[:] = [x for x in row if x != ' ' and x not in ['10/11' '11/12', '12/13', '13/14', '14/15', '15/16']]
    # catches index errors
    try:
        # a 'Q' indicates a PreAP class
        if 'Q' in row[0].replace(" ","").split(':')[1]:
            going_to_return = [row[0].replace(" ","").split(':')[0] + "PreAP", row[2][0:3], row[1].replace(' ', "")]
        else:
            going_to_return = [row[0].replace(" ","").split(':')[0], row[2][0:3], row[1].replace(' ', "")]
    except IndexError:
        going_to_return = [row[0].replace(" ","").split(':')[0], row[2][0:3], row[1].replace(' ', "")]
    return going_to_return


def parse_transcripts(transcripts):
    all_students = []
    student_index = -1
    j = 0
    max = len(transcripts)
    for row in transcripts:
        j +=  1
        if type(row[1]) is str and re.search('Student No: ', row[1]):
            student_index += 1 # updates the student index
            all_students.append([get_id(row), get_name(row)]) # puts the id in slot 0, name in slot 1
        if len(row) == 5:
            all_students[student_index].append(get_course_credit_and_grade(row)) # gives course and credit to student
        # !!!!!!!!!!!!!!!!!!!!!!!!!!!  uncomment if and indent all_students... to fix change
        update_progress(j/max, "Parsing transcripts")
    return all_students

# Finds all of the unique courses, does not count frequency
def find_unique(list):
    uniques = []
    unique_dict = {}
    for row in list:
        for column in row:
            if len(column) == 3 and column[0] not in unique_dict:
                unique_dict[column[0]] = 1
            elif len(column) == 3:
                unique_dict[column[0]] += 1
    for entry in unique_dict:
        uniques.append([entry, unique_dict[entry]])
    # csv_writer(uniques, 'unique_classes.csv')
    return uniques

# counts how much credit each student has. Written after the other reading functions were done, so not as effient as getting it directly from the transcripts
def count_credit_earned(transcripts, uniques):
    credit_counter = {}
    for course in uniques:
        credit_counter[course] = 0
    for student_courses in transcripts:
        for course in student_courses:
            if len(course) == 2 and course[1] == '0.5' or course[1] == '1.0' or course[1] == '0.0':
                credit_counter[course[0]] += 1
    print(len(transcripts), "total students in grade")
    # print(credit_counter['ALG1A'], " students have credit for Alg 1 A")
    return credit_counter, len(transcripts)

# Creates course requests in the format required for the third party program ScheduleSmart
def course_requests(classes, transcripts, uniques, class_size):
    credit_dict, kids = count_credit_earned(transcripts, uniques)
    requests = []
    for course in classes:
        try:
            earned = credit_dict[course]
            needed = kids - earned
            classes_needed = math.ceil(float(needed / class_size))
            requests.append([course, earned, needed, classes_needed])
        except KeyError:
            requests.append([course, "no credits earned", '', ''])
    for row in requests:
        print(row[2], "kids need", row[0], "in", row[3], "classes of size", class_size)

section_predictor = ['ALG1A', 'GEOMA', 'error', 'ALG2A', 'PRECALCA', 'ENG1A', 'ENG2A', 'ENG3A', 'ENG4A', 'BIOA',
                     'CHEMA', 'PHYSICSA', 'WGEOA', 'WHISTB', 'USHISTA']


def course_requests(transcripts, department_courses):
    ss_index = 0
    schedule_smart_crs = []
    for student in transcripts:
        course_index = 2
        temp_student_index = 0
        temp_student = [student[0], student[1]]
        while temp_student_index < len(department_courses):
            temp_student.append(0)
            temp_student_index += 1
        schedule_smart_crs.append(temp_student) # append new student
        for course in department_courses:
            for specific_course in course:
                for credit in student:
                    if credit[0] == specific_course and credit[1] != '0.0':
                        schedule_smart_crs[ss_index][course_index] = 1
                        break
            course_index += 1
        ss_index += 1
    return schedule_smart_crs


def find_next_course(students, courses):
    j = 0
    while j < len(students):
        i = len(courses)
        while i >= 1:
            if students[j][i] == 1:
                currents = students[j]
                inserting = courses[i-1][0]
                students[j].insert(1, courses[i-1][0])
                break
            else:
                i -= 1
        if i == 0:
            # students[j].insert(1, courses[i][0])
            del students[j]
            j -= 1
        j += 1
    for student in students:
        smart_requests.append([student[0], student[2], student[1]])

def calc_credits(kid):
    i = 2
    credits = 0
    while i < len(kid):
        credits += float(kid[i][1])
        i += 1
    return credits

# need to add more courses
# group unaccounted for credits at the end
# this function makes a file for mail merging into the district's graduation plan word doc
def compute_grad_requirements(transcripts):
    all_kids = []
    unique_courses = find_unique(transcripts)
    # possible_classes = ['null', 'null', 'null']
    possible_classes = []
    for c in unique_courses:
        possible_classes.append(c[0])
    possible_classes.sort()
    possible_classes.insert(0, 'ID')
    possible_classes.insert(1, "Name")
    possible_classes.insert(2, "Credits")
    print(possible_classes)
    for kid in transcripts:
        full_name = kid[1]
        new_name = full_name.replace("  ", ", ")
        tempkid = [kid[0], new_name, 0.0]
        for j in range(len(possible_classes) - 3):
            tempkid.insert(3, '')
        tempkid[2] = str(calc_credits(kid))
        tempcourses = {}
        for course in kid[2:]:
            tempcourses[course[0]] = {'grade': course[2], 'credits': course[1]}
        for i in range(len(possible_classes)):
            pci = possible_classes[i]
            testvar = kid[2:]
            if pci in tempcourses:
                try:
                    course_grade = int(tempcourses[pci]['grade'])
                    if tempcourses[pci]['credits'] == '0.0':
                        tempkid[i] = str(course_grade) + ' X'
                    else:
                        tempkid[i] = str(course_grade)
                except ValueError:
                    tempkid[i] = 'NG X'
        all_kids.append(tempkid)
    all_kids.insert(0, possible_classes)
    csv_writer(all_kids, 'first draft of credit mail merge thing.csv')
    print()




def complete_grade_level(excel_file, section_predictor, max_class_size):
    transcripts = load_transcripts(excel_file)
    student_credit_list = parse_transcripts(transcripts)
    unique_courses = find_unique(student_credit_list)
    course_requests(section_predictor, student_credit_list, unique_courses, max_class_size)

# excel_file = '3students.xlsx'
# excel_file = 'all12thtranscripts.xlsx'
# excel_file = 'all11thtranscripts.xlsx'
# excel_file = 'all10thtranscripts.xlsx'
# transcripts = load_transcripts(excel_file)
# student_credit_list = parse_transcripts(transcripts)
# unique_courses = find_unique(student_credit_list)
# csv_writer(unique_courses, 'unique with preap maybe.csv')
# print('done')
# # count_credit_earned(student_credit_list, unique_courses, "12")
# course_requests(section_predictor, student_credit_list, unique_courses, 35)

# test for mail merge
# compute_grad_requirements(student_credit_list)

all_transcripts = ['all10thtranscripts.xlsx', 'all11thtranscripts.xlsx', 'all12thtranscripts.xlsx']

section_predictor_10th = ['ALG1A', 'GEOMA', 'ENG1A', 'ENG2A', 'BIOA', 'CHEMA', 'WGEOA', 'WHISTA']
section_predictor_11th = ['ALG1A', 'ALG1B', 'GEOMA', 'ALG2A', 'PRECALCA', 'ENG1A', 'ENG1B', 'ENG2A', 'ENG3A', 'BIOA',
                          'BIOB', 'CHEMA', 'PHYSICSA', 'WGEOA', 'WGEOB', 'WHISTA', 'USHISTA']
section_predictor_12th = ['ALG1A', 'ALG1B', 'GEOMA', 'GEOMB', 'ALG2A', 'PRECALCA', 'ENG1A', 'ENG1B', 'ENG2A', 'ENG2B',
                          'ENG3A', 'BIOA', 'BIOB', 'CHEMA', 'CHEMB', 'PHYSICSA', 'WGEOA', 'WGEOB', 'WHISTA', 'WHISTB',
                          'USHISTA']

# complete_grade_level('all9thtranscripts.xlsx', section_predictor_10th, 30)
# complete_grade_level('all10thtranscripts.xlsx', section_predictor_11th, 30)
# complete_grade_level('all11thtranscripts.xlsx', section_predictor_12th, 30)
# print('how many classes do we need for rising 10th graders?')

course_sequences = [[['ENG1A', 'ENG1SOLA'], ['ENG2A', 'ENG2SOLA'], ['ENG3A', 'APENGLANA'], ['ENG4A', 'APENGLITA']],
                    [['ALG1A'], ['GEOMA'], ['ALG2A', 'ALG2AM', 'ALG2AT'], ['PRECALCA', 'APCALCABA', 'APSTATSA']],
                    [['BIOA', 'BIOLOGY'], ['CHEMA', 'CHEM', 'AP-CHEMA'], ['PHYSICSA', 'APPHYS1A'], ['AQUASCIA', 'AP-BIOA', 'AP-CHEMA', 'ANATPHYSA']],
                    [['WGEOA', 'APHUMGEOA'], ['WHISTA', 'APWHISTA'], ['USHISTA', 'APUSHISTA'], ['GOVT', 'GOVTT'], ['APMACECO', 'CONS-ECO', 'ECO-FE', 'ECO-FET']],
                    [['ART1A'], ['ART2A', 'ART2CRMCA', 'ART2DGMDA', 'ART2DRAWA', 'ART2JWLRA', 'ART2PHTOA', 'ART2SCLPA']],
                    [['DANCE1A'], ['DANCE2A'], ['DANCE3A']],
                    [['PEFOUNDA', 'ROTC1A', 'PEIS', 'PEITS'], ['PERHLTH1A']],
                    [['SPAN1A', 'SSSPAN1', 'SSSPAN1A'], ['SPAN2A', 'SSSPAN2', 'SSSPAN2A'], ['APSPANLITA','SPAN3A', 'SSSPAN3A'], ['APSPANLANGA','SPAN4A', 'SSSPAN4A']],
                    [['FREN1A'], ['FREN2A'], ['FREN3A'], ['FREN4A']],
                    [['MUS1BANDA'], ['MUS2BANDA'], ['MUS3BANDA']],
                    [['MUS1THYA'], ['MUS2THYA'], ['MUS3THYA'], ['APMUSTHYA']],
                    [['MUS1CHORA'], ['MUS2CHORA'], ['MUS3CHORA']],
                    [['TH1A', 'TH1PRODA', 'TH1TECHA'], ['TH2A', 'TH2PRODA', 'TH2TECHA'], ['TH3A'], ['TH4A']],
                    [['DEBATE1A'], ['DEBATE2A'], ['DEBATE3A']],
                    [['INTCOSMOA'], ['COSMET1A'], ['COSMET2A'], ['PRACTHLSCA']],
                    [['PRINITA'], ['SECOND PIT COURSE', 'COMPMTNA', 'COMPPROGA', 'DIMEDIAA'], ['THIRD PIT COURSE', 'ADVCOMPPA', 'WEBTECHA'], ['FOURTH PIT COURSE']],
                    [['PRINMANA'], ['FLEXIBLE MANUFACTURING']],
                    [['CONCENGTA'], ['ENGPRSA'], ['ROBOTAA'], ['PRACTICUM']],
                    [['PRINTDLA'], ['AUTOTECH?']],
                    [['ENG1APreAP'], ['ENG2APreAP'], ['ENG3APreAP'], ['ENG4APreAP']],
                    [['ALG1APreAP'], ['GEOMAPreAP'], ['ALG2APreAP'], ['PRECALCAPreAP']],
                    [['BIOAPreAP'], ['CHEMAPreAP'], ['PHYSICSAPreAP'], ['AP-BIOA']],
                    [['WGEOAPreAP'], ['WHISTAPreAP'], ['USHISTAPreAP'], ['GOVTPreAP'], ['ECO-FEPreAP']]]

# crs = course_requests('3students.xlsx', ['ENG1A', 'ENG2A', 'ENG3A', 'ENG4A'])
# find_next_course(crs, ["ENG1", "ENG2", "ENG3", "ENG4"])
# for grade in ['9th transcripts 4 4.xlsx', '10th transcripts 4 4.xlsx', '11th transcripts 4 4.xlsx']:
#     transcripts = load_transcripts(grade)
#     student_credit_list = parse_transcripts(transcripts)
#     for seq in course_sequences:
#         crs = course_requests(student_credit_list, seq)
#         find_next_course(crs, seq)

transcripts = load_transcripts('3students.xlsx')
student_credit_list = parse_transcripts(transcripts)
for seq in course_sequences:
    crs = course_requests(student_credit_list, seq)
    find_next_course(crs, seq)

smart_requests.sort()
csv_writer(smart_requests, 'all reqs with names.csv')
for line in smart_requests:
    print(line)

emillis = int(round(time.time() * 1000))
dif = float((emillis - smillis)/1000)
print('Seconds elapsed:',dif)
