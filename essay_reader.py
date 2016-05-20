import win32com.client as win32
import glob, os, shutil, fnmatch, re, time
from lxml import etree
import zipfile
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfpage import PDFPage
from cStringIO import StringIO
import cx_Oracle

#------GLOBAL VARIABLES---------

#variable for text file processing
WORD_NAMESPACE = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
PARA = WORD_NAMESPACE + 'p'
TEXT = WORD_NAMESPACE + 't'
wdFormatUnicodeText = 7
word = win32.gencache.EnsureDispatch('Word.Application')
word.Visible = False

#Directories to search and store
_src_dir = 'R:\\Assessments\\EARLY ADMISSION\\Essays\\OriginalEssays\\PZ'
_copy_dir = 'R:\\Assessments\\EARLY ADMISSION\\Essays\\Copies'
_clean_dir = 'R:\\Assessments\\EARLY ADMISSION\\Essays\\CleanCopies'
_trg_dir = 'R:\\Assessments\\EARLY ADMISSION\\Essays\\Schools'
start_time = time.time()

#expression match for MPS Student Ids
idexp = re.compile('[8-9]\d{6,6}')

#database connections and dictionaries
db = cx_Oracle.connect('k12intel_metadata', 'javelin1912', 'ex01-scan.milwaukee.k12.wi.us:1521/RUNDWPRODIC.world')

cursor = db.cursor()
named_params = {'grade':'08', 'status':'Enrolled', 'activity':'Active', 'school_status':'Y'}
cursor.execute('SELECT student_id, student_name FROM K12INTEL_DW.DTBL_STUDENTS INNER JOIN K12INTEL_DW.DTBL_SCHOOLS on dtbl_schools.school_code = dtbl_students.student_next_year_school_code WHERE student_current_grade_code = :grade and student_status = :status and student_activity_indicator = :activity and reporting_school_ind = :school_status', named_params)
student_names_dict = dict(cursor.fetchall())

cursor.execute('SELECT student_id, school_name FROM K12INTEL_DW.DTBL_STUDENTS INNER JOIN K12INTEL_DW.DTBL_SCHOOLS on dtbl_schools.school_code = dtbl_students.student_next_year_school_code WHERE student_current_grade_code = :grade and student_status = :status and student_activity_indicator = :activity and reporting_school_ind = :school_status', named_params)
student_sch_dict = dict(cursor.fetchall())

cursor.execute('SELECT distinct school_name, school_code FROM K12INTEL_DW.DTBL_STUDENTS INNER JOIN K12INTEL_DW.DTBL_SCHOOLS on dtbl_schools.school_code = dtbl_students.student_next_year_school_code WHERE student_current_grade_code = :grade and student_status = :status and student_activity_indicator = :activity and reporting_school_ind = :school_status', named_params)
schools_list = [item[0] for item in cursor.fetchall()]



#-------FUNCTIONS--------------

####--PROCESSING FILES---##

#Function to walk through all subdirectories in a root directory and make a file list of all 
#word docs and pdfs 
def get_files(drcty):
	file_list = []
	for root, dirs, files in os.walk(drcty):
		for file in files:
			if doc_pdf:
				doc = root + '\\' + file
				file_list.append(doc)
	return file_list

#Function to save copies of list of files to a copy directory'''
def save_copies(files, copy_dir):
	for file in files:
		if doc_pdf(file) in ['pdf', 'doc']:
			try:
				shutil.copy(file, copy_dir)
			except: 
				print('File did not copy: ' + file)
			
	
#Function to check if file is a word doc or a pdf	
def doc_pdf(file):
	if fnmatch.fnmatch(file, '*.doc?'):
		return 'doc'
	elif fnmatch.fnmatch(file, '*.pdf'):
		return 'pdf'
	else:
		None

		
#Function to get the ID from the original filename if it is present
def get_id(text):
	match = idexp.search(text)
	if match:
		id = match.group()
	else:
		id = None
	return id

	
#Function to get student ID if it exists in the text of the document
def parse_doc_id(path):	
	id = get_id(path)
	if not id:
		document = zipfile.ZipFile(path)
		xml_content = document.read('word/document.xml')
		tree = etree.fromstring(xml_content)
		id = None
		for node in tree.iter(etree.Element):
			if check_element_is(node, 't'):
				try:
					id = get_id(node.text)
				except:
					id = None
			if id:
				break
	return id

	
#Function to parse a PDF and return a tuple of (Student ID, text of PDF)	
def parse_pdf(path):
	id = get_id(path)
	txt = None
	rscmgr = PDFResourceManager()
	retstr = StringIO()
	codec = 'ascii'
	laparams = LAParams()
	device = TextConverter(rscmgr, retstr, codec=codec, laparams=laparams)
	fp = open(path, 'rb')
	interpreter = PDFPageInterpreter(rscmgr, device)
	password = ""
	maxpages = 0
	caching = True
	pagenos=set()
	for page in PDFPage.get_pages(fp, pagenos, maxpages=maxpages, password=password, caching=caching, check_extractable=True):
		interpreter.process_page(page)
	fp.close()
	device.close()
	txt = retstr.getvalue()
	retstr.close()
	if not id:
		id = get_id(txt)
	return (id, txt)

	
def save_doc_as_text(filepath, new_name, trg_dir):
	try:
		doc = word.Documents.Open(filepath)
		doc.SaveAs(trg_dir + '\\' + new_name + '.txt', wdFormatUnicodeText)
		doc.Close()
	except:
		print syserr
		print 'Could not save word file as text: ' + filepath


def save_idtxt_to_file(id_txt, trg_dir):
	newfile = open(trg_dir + '\\' + id_txt[0] + '.txt', 'w')
	newfile.write(id_txt[1])
	newfile.close()
		
#Function to check if xml element is text from word doc	
def check_element_is(element, type_char):
	return element.tag == TEXT	

		
#Function to take file list of docs and pdfs, check for ID,
#create new text file named with ID and removed old file
def clean_store(src_files, clean_dst):
	for file in src_files:
		try:
			id = None
			txt = None
			if doc_pdf(file) == 'doc':
				id = parse_doc_id(file)
				if id:
					save_doc_as_text(file, id, clean_dst)
					os.remove(file)
			elif doc_pdf(file) == 'pdf':
				id_txt = parse_pdf(file)
				if id_txt[0]:
					save_idtxt_to_file(id_txt, clean_dst)
					os.remove(file)
		except:
			print 'File could not be processed: ' + file			

			
			
##---PARSING FILES AND STORING CLEAN ESSAYS---##
			
#Function to create empty directories for holding each schools essays
def create_school_folders(school_list, trg_dir):
	for school in school_list:
		if not os.path.exists(trg_dir + '\\' + school):
			os.makedirs(trg_dir + '\\' + school)

			
#Read text from a file parsing, store id, essay, stu name and school into new file
def store_essay(fname, trg_dir):
		with open(fname) as file:
			base_fname = os.path.split(fname)[1]
			id = base_fname[0:7]
			essay = ''
			run = 0
			for line in file:
				if line not in ['\n', ' \n']:
					if re.search('Please type your essay into this',line):
						run = 1
					elif re.search('When you have completed your essay',line):
						run = 0
					if run == 1:
						essay = essay + line
		try:
			stu_name = student_names_dict[id]
			stu_school = student_sch_dict[id]
			new_file = open(trg_dir + '\\' + stu_school + '\\' + stu_name + '_' + id + '.txt', 'w')
			new_file.write(essay)
			os.remove(fname)
		except:
			pass

		
#main method to run to walk through files, get id, save copy with new name to appropriate folder		
def main(src_dir, copy_dir, clean_dir, target_dir):
	
	src_files = get_files(src_dir)
	if len(src_files) > 0:
		save_copies(src_files, copy_dir)
		copy_files = get_files(copy_dir)
		clean_store(copy_files, clean_dir)	
	
	create_school_folders(schools_list, target_dir)
	
	clean_files = get_files(clean_dir)
	if len(clean_files) > 0:
		for file in clean_files:
			store_essay(file, target_dir)
		
	

#-------PROGRAM EXECUTION----------

if __name__=='__main__':
	main(_src_dir, _copy_dir, _clean_dir, _trg_dir)

word.Quit()

#--------PROGRAM STATISTICS-------------
		
#Print Runtime to Console	
print("----{} seconds ----").format(int(time.time() - start_time))	
print("----{} minutes, {} seconds ----").format(int((time.time() - start_time) / 60), (time.time() - start_time) % 60)
				

#TESTS
#print(get_files(drct))
#print(get_id('donald 8738728 897903892')) # '8738728'
#print(get_id('samueal Jackson')) #none	
#print(doc_pdf('testone.docx')) #return doc
#print(doc_pdf('testtwo.pdf')) # return pdf
#print(doc_pdf('testthree.db')) # return None	
#print student_names_dict['8637406'] # Watksins, Keanna N.
#print student_sch_dict['8637406'] # OBAMA SCTE
#print schools_list