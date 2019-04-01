from tkinter import filedialog
from tkinter import *
import os
import csv
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfpage import PDFPage
from io import StringIO


root = Tk()
root.title("Result")
root.geometry("600x370")
root.configure(background='#9400D3')

text = Text(root)
n_name_index=0
subjects=[]
subjectCount=0
sl=[]

def convert_pdf_to_txt(path):
    print("Started Converting pdf to text")
    rsrcmgr = PDFResourceManager()
    retstr = StringIO()
    codec = 'utf-8'
    laparams = LAParams()
    device = TextConverter(rsrcmgr, retstr, codec=codec, laparams=laparams)
    fp = open(path, 'rb')
    interpreter = PDFPageInterpreter(rsrcmgr, device)
    password = ""
    maxpages = 0
    caching = True
    pagenos=set()

    for page in PDFPage.get_pages(fp, pagenos, maxpages=maxpages, password=password,caching=caching, check_extractable=True):
        interpreter.process_page(page)

    text = retstr.getvalue()

    fp.close()
    device.close()
    retstr.close()
    print("PDF To Text Success!")
    return text

def getfile():
	global raw_data
	text.insert(INSERT, "Started...PDF 2 TEXT\n")
	filename =  filedialog.askopenfilename(initialdir = "{}".format(os.getcwd()),title = "Select file",filetypes = (("pdf files","*.pdf"),("all files","*.*")))

	raw_data=convert_pdf_to_txt(filename)
	text.insert(INSERT, "Completed...PDF 2 TEXT\n")
	preprocess(raw_data)

def savefile():
	global sl
	options = {'filetypes':[('CSV','.csv')]}
	outfile = filedialog.asksaveasfilename(**options)
	print(outfile)
	with open(outfile, 'w') as writeFile:
		writer = csv.writer(writeFile)
		writer.writerows(sl)


def getSubjectList(ndata):
	text.insert(INSERT, "Preprocessing data : TRYING TO FETCH SUBJECT LISTS\n")
	array=ndata[1:n_name_index-1]
	global subjects
	global subjectCount
	for a in array:
		subject_id=a.strip().split(" ")[0]
		subjects.append("{}:{}".format(subject_id,"OE"))
		subjects.append("{}:{}".format(subject_id,"TH"))	
		subjects.append("{}:{}".format(subject_id,"OE+TH"))	
		subjects.append("{}:{}".format(subject_id,"TW"))
		subjects.append("{}:{}".format(subject_id,"PR"))
		subjects.append("{}:{}".format(subject_id,"OR"))
		subjectCount+=1
	#text.insert(INSERT, "Preprocessing data : Subject list fetched \n{}\n".format(subjects))
	print(subjects)
	
def getName(n,n_name_index,ndata):
	a=ndata[n_name_index*n]
	name=a.split(' ')
	name=list(filter(('').__ne__,name))
	return name[1]+" "+name[2]+" "+name[3]

def getPRN(n,n_name_index,ndata):
	a=ndata[n_name_index*n]
	name=a.split(' ')
	name=list(filter(('').__ne__,name))
	return name[0]

def getSubjectData(sno,subjectNo,n_name_index,ndata):
	index=subjectNo+(sno*n_name_index)+1
	#print("index {}".format(subjectNo+(sno*n_name_index)+1))
	sub1=ndata[index].split(' ')
	sub1=list(filter(('').__ne__,sub1))
	return sub1	

def getStudentCount(ndata):
	count=1
	a=getName(0,n_name_index,ndata)
	while a!="------- ------- -------":
		try:
			a=getName(count,n_name_index,ndata)
			count+=1
		except:
			break
	count-=2
	return count

def preprocess(data):
	global subjectCount
	text.insert(INSERT, "Preprocessing data : Splitting based on new line\n")
	ndata=data.split("\n")
	space=ndata[0]
	uni=ndata[1]
	college=ndata[3]
	branch=ndata[4]
	date=ndata[5]
	dots=ndata[6]
	grd=ndata[9]
	header='               IN       TH     [IN+TH]     TW       PR       OR    Tot% Crd  Grd  Pts   Pts'
	dashes=ndata[11]
	special='\x0c '
	header1='               OE       TH     [OE+TH]     TW       PR       OR    Tot% Crd  Grd  Pts   Pts'
	special2='\x0c'
	blank=''
	sem1=' SEM.:1'
	sem2=' SEM.:2'
	reservedBacklog="RESULT RESERVED FOR BACKLOGS."
	text.insert(INSERT, "Preprocessing data : Removal of unwanted data\n")
	for i in range (20):
		for d in ndata:
			if (d==header) or (d==space) or (d==uni) or (d==sem1) or (d==header1): 
				ndata.remove(d)
			if (d==college) or(d==branch) or (d==date) or (d==sem2) or (d==reservedBacklog):
				ndata.remove(d)  
			if(d==dots) or (d==grd) or (d==dashes) or (d==special) or (d==blank) or (d==special2):
				ndata.remove(d)
	text.insert(INSERT, "Preprocessing data : Removal of unwanted data complete\n")
	#ndata
	with open("output.txt", 'w') as file_handler:
		for item in ndata:
		    file_handler.write("{}\n".format(item))
	text.insert(INSERT, "Preprocessing data : File Generated as output.txt\n")
	determinNameIndex(ndata)
	getSubjectList(ndata)
	global sl
	subjects.insert(0,"Exam Roll No")
	subjects.insert(1,"Name")
	subjects.append("sgpa")
	subjects.append("total credits earned")
	sl.append(subjects)
	for i in range(getStudentCount(ndata)+2):

		#print(getSubjectScore(i,j,n_name_index,ndata))
		sr=[]
		prn=getPRN(i,n_name_index,ndata)
		sr.append(prn)
		if(getName(i,n_name_index,ndata) =="------- ------- -------"):
			text.insert(INSERT, "Skipped useless data\n")
			continue
		sr.append(getName(i,n_name_index,ndata))
		n=3
		for n in range(1,subjectCount):
			data=getSubjectData(i,n,n_name_index,ndata)
			try:
				data.remove("*")
			except:
				print("* found")			
			print(len(data))
			try:
				if "T" in prn or "S" in prn:
					sr.append(data[1].split('/')[0])
					sr.append(data[2].split('/')[0])
					sr.append(data[3].split('/')[0])
					sr.append(data[4].split('/')[0])
					sr.append(data[5].split('/')[0])
					sr.append(getSubjectData(i,n,n_name_index,ndata)[6].split('/')[0])
				else:
					sr.append(data[2].split('/')[0])
					sr.append(data[3].split('/')[0])
					sr.append(data[4].split('/')[0])
					sr.append(data[5].split('/')[0])
					sr.append(data[6].split('/')[0])
					sr.append(getSubjectData(i,n,n_name_index,ndata)[7].split('/')[0])
			except:
				continue
		sr.append("")
		sr.append("")
		sr.append("")
		sr.append("")
		sr.append("")
		sr.append("")
		if "THIRD" in getSubjectData(i,subjectCount,n_name_index,ndata):
			print(getSubjectData(i,subjectCount,n_name_index,ndata))
			sr.append(getSubjectData(i,subjectCount,n_name_index,ndata)[4]) #sgpa
			sr.append(getSubjectData(i,subjectCount,n_name_index,ndata)[-1]) #sgpa
		elif "SECOND" in getSubjectData(i,subjectCount,n_name_index,ndata):
			sr.append(getSubjectData(i,subjectCount,n_name_index,ndata)[4]) #sgpa
			sr.append(getSubjectData(i,subjectCount,n_name_index,ndata)[-1]) #sgpa
		else:
			sr.append(getSubjectData(i,subjectCount,n_name_index,ndata)[2]) #sgpa
			sr.append(getSubjectData(i,subjectCount,n_name_index,ndata)[-1]) #sgpa
		sl.append(sr)

	button2 = Button(frame, text="Save File",bg="white", fg='#9400D3',padx=12,pady=12,font='Helvetica 18 bold', command=savefile)
	button2.pack( side = BOTTOM)
	print("Complete")
def determinNameIndex(ndata):
	global n_name_index
	text.insert(INSERT, "Preprocessing data : Name Index determination\n")
	for i in range(30):
		if "SGPA" in ndata[i]:
			n_name_index=i+1
			break
	text.insert(INSERT, "Preprocessing data : Name Index determination complete found at {} index \n".format(n_name_index))
	text.insert(INSERT, "Preprocessing data : Test - Name Index determination check get name roll no 3 {} \n".format(getName(3,n_name_index,ndata)))


frame = Frame(root)
frame.pack()
bottomframe = Frame(root)
bottomframe.pack( side = BOTTOM )
button = Button(frame, text="Choose File", bg="white", fg='#9400D3',padx=12,pady=12,font='Helvetica 18 bold', command=getfile)


button.pack( side = BOTTOM)

text.pack()


root.mainloop()
