import argparse
import re
import xml.etree.ElementTree as ET
import zipfile
import os
import sys

nsmap = {
'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}


# pass file name
def process_args():
	parser = argparse.ArgumentParser()
	parser.add_argument("docx", help="path of the docx file")
	parser.add_argument('txt', help="text file of replace values")
	
	args = parser.parse_args()
	if not os.path.exists(args.docx):
		print 'docx File Not Found...'
		sys.exit(1)
	
	if args.txt is not None:
		 if not os.path.exists(args.txt):
			print 'text File Not Found...'
			sys.exit(1)
	return args

# creating a json that reads from file that contains 
# replace data
def create_json(val):
	#print val
	infile = open(val, 'r')
	lines = infile.readlines()
	d = {}
	for line in lines:
		#print line
		sp = line.split(':')
		d[sp[0]] = sp[1]

	return d


# qualified name
def qn(tag):
    	prefix, tagroot = tag.split(':')
    	uri = nsmap[prefix]
    	
	return '{{{}}}{}'.format(uri, tagroot)


# process xml
def xml_process(xml, j):
	print 'xml--> '+xml
	text = u''
	tree = ET.parse(xml)
        root = tree.getroot()
        for child in root.iter():
                if child.tag == qn('w:t'):
                        t_text = child.text
                        #print t_text
			temp = t_text.replace('*','')
			for key in j:
				if key == temp:
					child.text = j[key]
			#print temp
                elif child.tag == qn('w:tab'):
                        continue
                elif child.tag in (qn('w:br'), qn('w:cr')):
                        continue
                elif child.tag == qn("w:p"):
                        continue
	
	tree.write(xml)
	print 'docx xml value replacement complete...'

# extract the docx achive
def docx_extract(docx):
	output = 'temp/'
	zipf = zipfile.ZipFile(docx, 'r')
	zipf.extractall(output)
	zipf.close()
	
	return output+'word/document.xml'


def new_docx(xml):
    outfile = 'output.docx'
    zipf = zipfile.ZipFile(outfile, 'w')
    sp = xml.split('/')
    path = sp[0]
    # Iterate all the directories and files
    for root, dirs, files in os.walk(path):
        # Create a prefix variable with the folder structure inside the path folder. 
        # So if a file is at the path directory will be at the root directory of the zip file
        if root.replace(path,'') == '':
                prefix = ''
        else:
		# folder structure like
                # folder1/folder2/file.txt
                prefix = root.replace(path, '') + '/'
                if (prefix[0] == '/'):
                        prefix = prefix[1:]
        for filename in files:
                actual_file_path = root + '/' + filename
                zipped_file_path = prefix + filename
                zipf.write( actual_file_path, zipped_file_path)

    zipf.close()
    print "Done writing new docx --> "+outfile

if __name__ == '__main__':
	args = process_args() # process the user input
	j = create_json(args.txt) # create json
	xml = docx_extract(args.docx) # extract the docx files
	xml_process(xml, j) # change the docx xml entries
	new_docx(xml) # create new docx with new entry
