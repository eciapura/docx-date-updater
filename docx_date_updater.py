import datetime, sys, os, re

try:
    from docx import Document
except ModuleNotFoundError:
	print("python-docx is not installed on this machine. Please install it and try again. You may install it by accessing the terminal or command prompt and trying pip install python-docx. For more information, visit https://python-docx.readthedocs.io/en/latest/user/install.html.")
	sys.exit()

try:
	os.chdir(sys.argv[1])
except IndexError:
	print("You did not input a valid directory as the first argument. Please run the script again with a valid directory.")
	sys.exit()
except (FileNotFoundError, NotADirectoryError):
	print("That directory could not be accessed. Please run the script again with a valid directory.")
	sys.exit()

all_logs = [single_log for single_log in os.listdir() if ".docx" in single_log and "~$" not in single_log]

if len(all_logs) == 0:
	print("Note: There are no .docx files in the selected directory.")
	sys.exit()

string_potential_month_date = r"(January|February|March|April|May|June|July|August|September|October|November|December)\s\d{4}"
string_current_month_date = str(datetime.datetime.now().strftime("%B %Y"))

def ReplaceDate(log, outdate, newdate):
	log_paragraphs = list(log.paragraphs)
	for single_para in log_paragraphs:
		if not re.search(outdate, single_para.text):
			continue
		if single_para.runs:
			text_bulk = ''.join([characters.text for characters in single_para.runs])
			prospective_new_text = re.sub(outdate, newdate, text_bulk)
			single_para.runs[0].text = prospective_new_text
			for leftovers in single_para.runs[1:]:
				leftovers.text = ''

for logs in all_logs:
	current_log  = Document(logs)
	ReplaceDate(current_log, string_potential_month_date, string_current_month_date)
	current_log.save(logs)
