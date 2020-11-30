import PyPDF2
from os import path
import tkinter
from tkinter import *
from tkinter.ttk import *
from tkinter import filedialog, Checkbutton
from pathlib import Path
from time import sleep

""" Merges two sets of pages corresponding to the same document, with
one set scanned in the forward order and the other in reverse """

def in_path_key(e):
	pass

def in_path_focusout(e):
	if path.exists(in_path.get()):
		out_name.delete(0, 'end')
		out_name.insert(0, 'new_'+str(path.basename(in_path.get())))
	if str(out_path['state']) == 'disabled':
			out_path.configure(state='normal')
			out_path.delete(0, 'end')
			out_path.insert(0, Path(in_path.get()).parent)
			out_path.configure(state='disabled')

def in_browse_callback():
	name = filedialog.askopenfilename()
	if name != '':
		in_path.delete(0, 'end')
		in_path.insert(0, name)
		out_name.delete(0, 'end')
		out_name.insert(0, 'new_'+str(path.basename(name)))
		if str(out_path['state']) == 'disabled':
			out_path.configure(state='normal')
			out_path.delete(0, 'end')
			out_path.insert(0, Path(name).parent)
			out_path.configure(state='disabled')

def in_select_all(e):
	if parent.focus_get() != e.widget:
	    # select text after 50ms
	    in_path.after(50, select_all, e.widget)

def select_all(widget):
	# select text
	widget.select_range(0, 'end')
	# move cursor to the end
	widget.icursor('end')

def out_browse_callback():
	name = filedialog.askdirectory()
	if name != '':
		out_path.delete(0, 'end')
		out_path.insert(0, name)

def out_select_all(e):
	if parent.focus_get() != e.widget:
	    # select text after 50ms
	    out_path.after(50, select_all, e.widget)

def outname_select_all(e):
	if parent.focus_get() != e.widget:
	    # select text after 50ms
	    out_path.after(50, select_all, e.widget)

def state_changed():
	state = out_path['state']
	if str(state) == 'normal':
		temp = Path(in_path.get()).parent
		if str(temp) == '.':
			temp = in_path.get()
		out_path.delete(0, 'end')
		out_path.insert(0, temp)
		out_path.configure(state='disabled')
	else:
		out_path.configure(state='normal')


def merge():
	status_lab['text'] = '>> Merging (0%)...'

	# Create File objects for forward_pass.pdf and reverse_pass.pdf
	in_filename = in_path.get()
	in_file = open(in_filename, 'rb')

	# Create PdfFileReader objects for the forward and reverse files
	in_PDF = PyPDF2.PdfFileReader(in_file)	

	# Get the number of pages
	num_pages = in_PDF.getNumPages()

	popup = Toplevel()
	popup.title('Merging (0%)...')
	prog_label = Label(popup, text="Merging...\n")
	prog_label.grid(row=0,column=0, pady = 50, padx = 50)
	progress = Progressbar(popup, orient=HORIZONTAL, length=200, maximum = 100, mode='determinate')
	progress.grid(row=1, column = 0, pady = 20, padx = 50)

	# Create a PdfFileWriter object and write the pages from both
	# forward and reverse PDFs to it
	pdfWriter = PyPDF2.PdfFileWriter()
	for page_num in range(int(num_pages/2)):
		popup.update()
		pdfWriter.addPage(in_PDF.getPage(page_num))
		pdfWriter.addPage(in_PDF.getPage(num_pages-1-page_num))
		progress_val = int((page_num*200.0)/num_pages)
		progress['value'] = progress_val
		status_lab['text'] = '>> Merging ('+str(progress_val)+'%)...'
		popup.title('Merging ('+str(progress_val)+'%)...')

	# Create the out file
	out_filename = out_path.get()+'/'+out_name.get()+'.pdf'
	append_count = 1
	while path.exists(out_filename):
		# If the filename already exists, append a count to it
		out_filename = out_path.get()+'/'+out_name.get()+'('+str(append_count)+').pdf'
		append_count += 1
	out_file = open(out_filename, 'wb')

	# Write pages from pdfWriter to the out file
	pdfWriter.write(out_file)

	# Close all opened files
	out_file.close()
	in_file.close()

	popup.title('Success')
	prog_label['text'] = 'Success:\n'+str(out_name.get())+'.pdf created'
	progress['value'] = 100
	status_lab['text'] = '>> \''+str(out_name.get())+'.pdf\' created'
	popup.update()
	sleep(2)
	popup.destroy()

parent = Tk()
parent.title("Merge Scanned PDFs")

frame1 = Frame(parent)
frame2 = Frame(parent)
frame3 = Frame(parent)
frame4 = Frame(parent)
frame1.grid(row = 0, column = 0, sticky = W, pady = 15, padx = 30)
frame2.grid(row = 1, column = 0, sticky = W, pady = 30, padx = 30)
frame3.grid(row = 2, column = 0, sticky = W, pady = (30, 10), padx = 30)
frame4.grid(row = 3, column = 0, sticky = W, pady = (10, 30), padx = 30)

in_label = Label(frame1, text='Select Input File')
in_path = Entry(frame1, width = 60)
in_path.bind("<Button-1>", in_select_all)
in_path.bind("<Key>", in_path_key)
in_path.bind("<FocusOut>", in_path_focusout)
in_path.focus()
in_browse = Button(frame1, text='Browse', command=in_browse_callback)
in_label.grid(row = 0, column = 0, sticky=W, pady = 2, padx = 10)
in_path.grid(row = 1, column = 0, sticky=W, pady = 2, padx = 10, columnspan=14)
in_browse.grid(row = 1, column = 14, sticky=W, pady = 2, padx = 0)

out_label1 = Label(frame2, text='Select Output Directory')
out_path = Entry(frame2, width = 60)
out_path.bind("<Button-1>", out_select_all)
out_browse = Button(frame2, text='Browse', command=out_browse_callback)
out_checkbox = Checkbutton(frame2, text='Is the output path the same as the input path?', command=state_changed, onvalue=1, offvalue=0)
out_label2 = Label(frame2, text='Enter Output Filename')
out_name = Entry(frame2, width = 30)
out_name.insert(0, 'merged')
out_label1.grid(row = 0, column = 0, sticky=W, pady = 2, padx = 10)
out_path.grid(row = 1, column = 0, sticky=W, pady = 2, padx = 10, columnspan=14)
out_browse.grid(row = 1, column = 14, sticky=W, pady = 2, padx = 0)
out_checkbox.grid(row = 2, column = 0, sticky=W, pady = 2, padx = 10, columnspan = 2)
out_checkbox.select()
state_changed()
out_label2.grid(row = 3, column = 0, sticky=W, pady = 2, padx = 10)
out_name.grid(row = 3, column = 1, sticky=W, pady = 2, padx = 0)
out_name.bind("<Button-1>", outname_select_all)

filler_lab = Label(frame3, text=' ', width=45)
filler_lab.grid(row=0, column=0)
merge = Button(frame3, text='Merge', command=merge, width = 30)
merge.grid(row = 0, column = 1, sticky=E)
status_lab = Label(frame4, text='>> Written by Saud', foreground="#999999")
status_lab.grid(row=1, column=0)

parent.mainloop()