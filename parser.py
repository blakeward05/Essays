import re
import os
import json

#Variables
essays_dict = {}

#Directory of files to parse
_src_dir = 'R:\\Personal\\Blake\\Essays\\CleanCopies'
_trg_dir = 'R:\\Personal\\Blake\\Essays\\Essays'

def get_files(drcty):
	file_list = []
	for root, dirs, files in os.walk(drcty):
		for file in files:
			doc = root + '\\' + file
			file_list.append(doc)
	return file_list


#Read text from a file parsing for the start and end of the essay section
def get_essay(fname):
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
		essays_dict[id] = essay
		
		# new_file = open(_trg_dir + '\\' + base_fname, 'w')
		# new_file.write(essay)	

def merge_essays(id, essay):
	None#all_essays.update[id, essay]

files = get_files(_src_dir)
for file in files:
	get_essay(file)

print(essays_dict)	