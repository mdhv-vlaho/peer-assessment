#############################################
PEER FEEDBACK COMPILER
Developed by Danielle Vlaho and Mitchell Huot
#############################################

Notes:
This program was developed to sort, anonymise, and distribute peer feedback data from tutorial sessions in large-enrolment classes.
It was designed for use in the Department of Chemistry at McGill University.

You will need to have the following to run this program successfully:
	- Access to Microsoft forms or other application that can output data in .csv format
	- Forms_data_sorter file (this is the main program file)
	- Python 3.7 or later with module jinja2 installed 
	- Classlist file (stores student data)
	- Feedback data (downloaded from the form and saved as .csv, link below to form template)
	- Optional: peer_feedback_gui file (if you would like a more user-friendly version of the program)

This program is built based on output generated using Microsoft forms.  
Any adjustments to the order of questions (and the questions themselves) within the form will require adjusting the code of the "assessment_data" function to match, otherwise it may not compile. 
Link to form template: https://forms.office.com/Pages/ShareFormPage.aspx?id=cZYxzedSaEqvqfz4-J8J6sfnxKA-efJOvZUfC2t1h_RUOUdTQjFKMkI2RU03NEpPSERIQVlCRjk3TC4u&sharetoken=4tVKDbivWDGYEidt58xQ

0. Prepare folder structure: 
	a. peer_assessment
		i. classlists (archive)
		ii. gradebook_upload (archive)
		iii. template (contains html templates for email)
		iv. venv (this will be set up automatically with the python virtual environment when a new project is created)
		v. assessment_data_files (archived, future_stash)
	- All second-level folders should also contain an "archive" to store old files

0. Prepare classlist file:
	a. Download all classlist files (four in total: 211-001, 211-002, 212-001, 212-002)
	b. Create a new .csv file with the following headings, or open and edit the "fullclasslist_template" file
		- ID number (listed as "OrgDefinedId")
		- Last name, First name (in same column)
		- Email
		- ID code (For class sizes of 100-899, use three digit numbers, starting at 101 and counting up. For larger classes, use 4-digit numbers starting at 1001)
		- Group (this is a placeholder)
		- Tutorial time (this is a placeholder)
		- TA (this is a placeholder)
		- room (this is a placeholder)
	c. Combine classlists together if necessary (e.g. two course codes which are cross listed, e.g. CHEM 212/211/242 and 222/234/252 at McGill) and rearrange information in the appropriate format, as in point b. above
	d. Save the file with filename {term}-{course}-fullclasslist.csv, e.g. "2024s-212-fullclasslist.csv"
		- "term" and "course" are input values that you will specify when the program is run
	Note: placeholders can be incorporated if desired. These are left in to allow flexibility in the program output. 

Y. Prepare weekly data:
	- On the forms page, download the week's data: 
		- "responses" tab > "download a copy"
	- Once the program has been run, the feedback data will automatically be renamed and stored in the "assessment_data_files" folder
	- Once you have downloaded the weekly data, you should scrub the form 
		- click Responses > 3 dots next to "open in excel" > Delete all responses

X. Run the program:
	- To run a program, right-click anywhere on code > run file in console
	- You can run either the "peer_feedback_gui.py" file or the "forms_data_sorter.py" file
		- The "peer_feedback_gui.py" file is more user-friendly, while the "forms_data_sorter.py" file will give more useful error messages if the gui is not working properly
	- You will be prompted to enter a date and course
		- For the date, use the format MMM XX, e.g. "Jan 01" or "May 23". This date will be incorporated as-is in emails to students, e.g. "Feedback data for May 23 Tutorial" 
		- This date will be reformatted automatically as mm/dd to rename files
	- To do a test run, check the "Is test run?" box and enter an email for the feedback to be sent to
		- This is a failsafe run that will send all of the compiled feedback to the specified address, to ensure that everything is working correctly
	- You will need to click the "Browse" button and navigate to the .xlsx feedback dataset to use for the run


