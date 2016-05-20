import win32com.client as win32
import glob, os, shutil, fnmatch, re, time
from lxml import etree
import zipfile
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfpage import PDFPage
from cStringIO import StringIO

#------GLOBAL VARIABLES---------

#variable for MS word file processing
WORD_NAMESPACE = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
PARA = WORD_NAMESPACE + 'p'
TEXT = WORD_NAMESPACE + 't'
wdFormatUnicodeText = 7
word = win32.gencache.EnsureDispatch('Word.Application')
word.Visible = False

#Directories to search and store
_src_dir = 'R:\\Personal\\Blake\\Essays\\2015'
_copy_dir = 'R:\\Personal\\Blake\\Essays\\Copies'
_clean_dir = 'R:\\Personal\\Blake\\Essays\\CleanCopies'
start_time = time.time()

#expression match for MPS Student Ids
idexp = re.compile('[8-9]\d{6,6}')


#-------FUNCTIONS--------------

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
		if doc_pdf(file) == 'pdf' or doc_pdf(file) == 'doc':
			shutil.copy(file, copy_dir)
		
			
#Function to check if file is a word doc or a pdf	
def doc_pdf(file):
	if fnmatch.fnmatch(file, '*.doc?'):
		return 'doc'
	elif fnmatch.fnmatch(file, '*.pdf'):
		return 'pdf'
	else:
		return None
		
		
#Function to get the ID from the original filename if it is present
def get_id(text):
	match = idexp.search(text)
	if match:
		id = match.group()
	else:
		id = None
	return id

#Function to check if xml element is text from word doc	
def check_element_is(element, type_char):
	return element.tag == TEXT	

	
#Function to get student ID if it exists in the text of the document
def parse_doc_id(path):	
	id = get_id(path)
	if not id:
		document = zipfile.ZipFile(path)
		xml_content = document.read('word/document.xml')
		tree = etree.fromstring(xml_content)
		#ftree = open('ftree.txt', 'w')
		#ftree.write(etree.tostring(tree, pretty_print = True))
		#ftree.close()
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
		print "none"


def save_idtxt_to_file(id_txt, trg_dir):
	newfile = open(trg_dir + '\\' + id_txt[0] + '.txt', 'w')
	newfile.write(id_txt[1])
	newfile.close()
		
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
			
			
#main method to run to walk through files, get id, save copy with new name to appropriate folder		
def main(src_dir, copy_dir, clean_dir):
	src_files = get_files(src_dir)
	if len(src_files) > 0:
		save_copies(src_files, copy_dir)
		copy_files = get_files(copy_dir)
		clean_store(copy_files, clean_dir)		
		

#-------PROGRAM EXECUTION----------

if __name__== '__main__':
	main(_src_dir, _copy_dir, _clean_dir)

word.Quit()
	
#--------PROGRAM STATISTICS-------------
		
#Print Runtime to Console			
print("----{} minutes, {} seconds ----").format(int((time.time() - start_time) / 60), (time.time() - start_time) % 60)
				

#TESTS
#print(get_files(drct))
#print(get_id('donald 8738728 897903892')) # '8738728'
#print(get_id('samueal Jackson')) #none	
#print(doc_pdf('testone.docx')) #return doc
#print(doc_pdf('testtwo.pdf')) # return pdf
#print(doc_pdf('testthree.db')) # return None	