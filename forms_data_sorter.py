import os
import csv
import shutil
import glob
from datetime import datetime
from smtplib import SMTP as SMTP
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import COMMASPACE, formatdate
from email.utils import make_msgid
import logging

import openpyxl
from jinja2 import Environment, FileSystemLoader, select_autoescape
env = Environment(
    loader = FileSystemLoader('template'),
    autoescape=select_autoescape(['html'])
)


def classlist_creation(term: str, course: str) -> (dict, dict, dict, list):
    """
    purpose:    create the classlist by sorting through student names and email addresses
    this function returns two dictionaries: classlist and feedback
    classlist_dict stores the name and email information for each student
    feedback_dict is instantiated here, and will store the assessed students' emails and peer feedback data
    """
    filename = os.path.join('classlists', f'{term}-{course}-fullclasslist.csv')
    classlist_dict: dict = {}
    feedback_dict: dict = {}
    studentcode_dict: dict = {}
    id_list: list = []
    # open classlist with ISO-8859-1 to read special characters
    with open(filename, encoding="ISO-8859-1", newline='') as csvfile:
        reader = csv.reader(csvfile)
        # skip the column headers in the first row
        next(reader, None)
        for row in reader:
            # store email addresses in feedback dictionary
            # store student information in classlist and student code dictionaries
            feedback_dict[row[2]] = []
            classlist_dict[row[0]] = [row[2], row[3]]
            studentcode_dict[row[3]] = row[2]
            id_list.append(row[0])
    return classlist_dict, feedback_dict, studentcode_dict, id_list


def assessment_data(term: str, course: str, instructor_list: list, file: str, studentcodecols = [-4]) -> (dict, dict):
    """
    purpose:        to process the feedback data generated by microsoft forms
    keyword arguments:
    namecolumns:    the columns corresponding to the 3-digit codes of the assessed students
                    default values for namecolumns are 8 (only one assessment per form)
    this function returns the feedback dictionary, filled with the peer assessment data
    """
    participation_grade = {}
    classlist_dict, feedback_dict, studentcode_dict, id_list = classlist_creation(term, course)

    wb = openpyxl.load_workbook(file)
    sheet = wb.active
    # process student codes from peer assessment spreadsheet
    for i, row in enumerate(sheet.iter_rows(min_row=2, values_only=True)):
        # columns containing emailaddress and studentID are hard-coded based on form template
        emailaddress = row[3]
        studentID = row[-7].strip()
        if emailaddress in instructor_list:
            instructor_flag = True
        else:
            instructor_flag = False
        if studentID in id_list:
            participation_grade[studentID] = 1
        for studentcodecolumn in studentcodecols:
            studentcode = row[studentcodecolumn]
            # the following if statement allows us to pass by blank rows
            studentcode = str(studentcode).strip()
            try:
                int(studentcode)
            except ValueError as e:
                raise ValueError(f'''Student code was not a numerical value in row {i+2},
                form submitted by {emailaddress}''')
            try:
                studentemail = studentcode_dict[studentcode]
            except KeyError:
                raise KeyError(f'''Student code {studentcode} was not in the list of codes in row {i+2}, 
                form submitted by {emailaddress}''')
            try:
                feedback_dict[studentemail].append([row[studentcodecolumn+1], row[studentcodecolumn+2],
                                                    row[studentcodecolumn+3], instructor_flag])
            except KeyError:
                raise KeyError(f'Email address {emailaddress} does not have a match in the list of student emails')

    # make a list of students who did not submit a peer feedback form
    for IDnum in id_list:
        if IDnum not in participation_grade:
            participation_grade[IDnum] = 0
    return feedback_dict, participation_grade


def feedback_data_shuttle(oldpath: str, newpath: str) -> None:
    shutil.move(oldpath, newpath)


def makecsv_tutpartic(participation_grade: dict, tutorialdate: str, course: str) -> None:
    """
    purpose:    to create a csv file for upload to myCourses gradebook with ID numbers and grades
    input:      date of tutorial and participation grade
                date will be input into filename and column header to allow myC upload to correct column
    note:       this function is built to work with D2L-based LMS and may need to be adjusted for other LMS
    """
    filepath = os.path.join('gradebook_upload', f'{course}-{tutorialdate}-upload.csv')
    with open(filepath, 'w', newline='') as csvfile:
        csvwriter = csv.writer(csvfile)
        csvwriter.writerow(['OrgDefinedId', '{} Points Grade'.format(tutorialdate), 'End-of-Line Indicator'])
        csvwriter = csv.writer(csvfile, lineterminator=',\n')
        for studentID, grade in participation_grade.items():
            csvwriter.writerow([studentID, grade])


def email(send_from: str, send_to: str, subject: str, text: str, server: str, port: str) -> None:
    """
    purpose:    to send out feedback emails
    input:      email addresses of sender and recipient, subject line, email content
    note:       this function may have to be adjusted based on your institution's email server
    """
    assert isinstance(send_to, list)
    msg = MIMEMultipart()

    msg['From'] = send_from
    msg['Subject'] = subject
    msg['Message-Id'] = make_msgid()
    msg['Date'] = formatdate(localtime=True)

    # record the MIME types of both parts
    message = MIMEText(text, 'html')

    # attach parts into message container
    msg.attach(message)

    smtp = SMTP(f'{server}:{port}')
    smtp.ehlo()
    smtp.starttls()
    smtp.sendmail(send_from, send_to, 'Subject: ' + subject + '\n' + msg.as_string())
    smtp.close()


def email_prep(feedback_dict: dict, tutorialdate: str, emailfrom: str, course: str, testrunemail: str, server: str, port: str, testrun=True) -> None:
    """
    purpose:    prepare email by loading html template and setting template variables
    note:       default mode is a test run which will send the complete feedback emails to a specified account
    final run:  make sure to set testrun = False
    """
    for key, item in feedback_dict.items():
        if item:
            template = env.get_template('feedbackemail.html')
            if testrun:
                emailaddress = testrunemail
            else:
                emailaddress = key
            emailmessage = template.render(tutorialdate=tutorialdate, evaluations=item)
            email(emailfrom, [emailaddress], f'CHEM {course} Tutorial Peer Assessment Feedback', emailmessage, server, port)
            print(key)
            logging.info(f'Email sent to: {key}. Testrun: {testrun}')


if __name__ == '__main__':
    print('Enter tutorial date:')
    tutorialdate = input()
    parsed_date = datetime.strptime(tutorialdate, '%b %d')
    formatted_tutorial_date = parsed_date.strftime('%m-%d')
    print('Enter course number:')
    course = input()
    print('Is this a test run? Y/N')
    testrun_input = input()
    if testrun_input.lower() == 'n':
        testrun = False
    else:
        testrun = True
    term = '2024s'
    emailfrom = 'chem2X2.chemistry@mcgill.ca'
    testrunemail = 'danielle.vlaho@mcgill.ca'
    server = 'mailhost.mcgill.ca'
    port = '25'
    xlsx_files = glob.glob('/Users/deevlaho/Downloads/*.xlsx')
    most_recent_file = max(xlsx_files, key=os.path.getmtime)
    filetime = datetime.fromtimestamp(os.path.getmtime(most_recent_file))
    formatted_time = filetime.strftime('%Y-%m-%d')
    print(most_recent_file)
    print(formatted_time)
    print('Press enter to continue or CRTL+C to exit')
    input()
    feedback_dict, participation_grade = assessment_data(term, course, ['danielle.vlaho@mcgill.ca'], most_recent_file)
    email_prep(feedback_dict, tutorialdate, emailfrom, course, testrunemail, server, port, testrun=testrun)
    makecsv_tutpartic(participation_grade, tutorialdate, course)
    newpath = f'/Users/deevlaho/Library/CloudStorage/OneDrive-McGillUniversity/Documents/PycharmProjects/peer-assessment/feedback_data_files/{term}-{course}-peer_feedback-{formatted_tutorial_date}.xlsx'
    if not testrun:
        feedback_data_shuttle(most_recent_file, newpath)


