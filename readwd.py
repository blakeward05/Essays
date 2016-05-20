from lxml import etree
import re
import zipfile

WORD_NAMESPACE = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
PARA = WORD_NAMESPACE + 'p'
TEXT = WORD_NAMESPACE + 't'
idexp = re.compile('[8-9]\d{6,6}')

def getid_fname(text):
	match = idexp.search(text)
	if match:
		id = match.group()
	else:
		id = None
	return id

def get_docx_idtxt(path):	
	document = zipfile.ZipFile(path)
	xml_content = document.read('word/document.xml')
	tree = etree.fromstring(xml_content)
	id = None
	for node in tree.iter(etree.Element):
		if check_element_is(node, 't'):
			try:
				id = getid_fname(node.text)
			except:
				id = None
		if id:
			print id
			break
	# 
	
def check_element_is(element, type_char):
	return element.tag == TEXT
	
get_docx_idtxt('C:\\Users\\wardb\\Ahmed.docx')