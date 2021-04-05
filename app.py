from tkinter import *
from tkinter import messagebox,ttk
import os
import webbrowser
import mailchimp_python_function1 as mpf1
import mailchimp_python_function2 as mpf2
import newsletter_template
from PIL import ImageTk,Image
from tkinter import filedialog as fd
import xlrd
import io
import time

root = Tk()
# root.geometry('1350x700+0+0')

# full screen (hides taskbar and title bar)
# root.attributes('-fullscreen', True)

# determines screen height and width and adjusts by itself
w, h = root.winfo_screenwidth(), root.winfo_screenheight()
root.geometry("%dx%d+0+0" % (w, h))

# opens in maxmized window
root.state('zoomed')

root.title('E-mail promotion')
root.iconbitmap('icon.ico')
bg_color = 'gray20'
choice = IntVar()
choice.set(1)
temp_list =[]
audience_id = "05dfb02cbf"
preview_text = ""

# ======================================================================================================
# functions
# ======================================================================================================

def add_text():
	global preview_text
	preview_text = prev_textarea.get("1.0", "end-1c")
	# preview_text.rstrip('\n')
	# print(preview_text+"something")
	preview_text_win.destroy()
	root.attributes('-disabled', False)
	root.wm_deiconify()

def quit_ptext():
	preview_text_win.destroy()
	root.attributes('-disabled', False)
	root.wm_deiconify()

def quit_add():
	adder.destroy()
	root.attributes('-disabled', False)
	root.wm_deiconify()

def quit_show():
	contacts_window.destroy()
	root.attributes('-disabled', False)
	root.wm_deiconify()

def disable_import():
	cont_input_area.config(state="normal")
	# scroll_y.config(state="normal")
	# open_file_text.delete(0,END)
	open_file_text.config(state="disabled")
	browse_btn.config(state="disabled")
	choice.set(1)

def disable_type():
	open_file_text.config(state="normal")
	browse_btn.config(state="normal")
	# cont_input_area.delete(0,END)
	cont_input_area.config(state="disabled")
	# scroll_y.config(state="disabled")
	choice.set(2)

'''def root_enable():
	cname_text.config(state="normal")
	name_text.config(state="normal")
	rto_text.config(state="normal")
	add_cont_btn.config(state="normal")
	show_cont_btn.config(state="normal")
	show_temp_btn.config(state="normal")
	send_btn.config(state="normal")
	fb_btn.config(state="normal")
	insta_btn.config(state="normal")
	in_btn.config(state="normal")
	github_btn.config(state="normal")
	textarea.config(state="normal")
	render_temp_btn.config(state="normal")

def root_disable():
	cname_text.config(state="disabled")
	name_text.config(state="disabled")
	rto_text.config(state="disabled")
	add_cont_btn.config(state="disabled")
	show_cont_btn.config(state="disabled")
	show_temp_btn.config(state="disabled")
	send_btn.config(state="disabled")
	fb_btn.config(state="disabled")
	insta_btn.config(state="disabled")
	in_btn.config(state="disabled")
	github_btn.config(state="disabled")
	textarea.config(state="disabled")
	render_temp_btn.config(state="disabled")'''

# function to open file
def browse():
	global location
	location= fd.askopenfilename()
	adder.wm_deiconify()
	open_file_text.delete(0,END)
	open_file_text.insert(0, location)

def append_cont():
	global email_list
	if (choice.get()==1):
		email_list = cont_input_area.get("1.0", END).splitlines()

		# stripping the whitespaces if any
		email_list = [email.strip() for email in email_list]

		email_list = list(set(email_list))
		email_list.sort()
		# with open("email_list.txt", "a") as f:
		# 	for email in email_list:
		# 		f.write(email + os.linesep)
		mpf1.add_members_to_audience_function(audience_id = audience_id, email_list = email_list)
	if (choice.get()==2):
		try:
			wb = xlrd.open_workbook(location)
		except:
			messagebox.showerror("Unable to open file", "Invalid directory, unsupported format or corrupt file. Make sure you selected a file with the extention (.xlsx).")
			open_file_text.delete(0,END)
			adder.wm_deiconify()
			return
		sheet = wb.sheet_by_index(0)
		sheet.cell_value(0, 0)
		for i in range(sheet.nrows):
			temp_list.append(sheet.cell_value(i, 0))
		temp_set = set(temp_list)
		email_list = list(temp_set)

		# stripping the whitespaces if any
		email_list = [email.strip() for email in email_list]
		
		email_list.sort()
		# with open("email_list.txt", "a") as f:
		# 	for email in email_list:
		# 		f.write(email + os.linesep)
		mpf1.add_members_to_audience_function(audience_id = audience_id, email_list = email_list)
	# root_enable()
	root.attributes('-disabled', False)
	adder.destroy()
		
def nameText(event):
	name_text.delete(0,END)

def rtoText(event):
	rto_text.delete(0,END)

def contText(event):
	cont_input_area.delete(1.0,END)

def render_temp():
	html = textarea.get("1.0", END)
	with io.open('newsletter_template.py','w',encoding='utf8') as f:
		text = "html_code = \"\"\"\\"+os.linesep+html+"\"\"\"#end"
		f.write(text)
	messagebox.showinfo("Template Rendered", "New template added successfully, click on 'Show Rendered Template' to view.")
	textarea.delete("1.0",END)

def show_template():
	with io.open('newsletter_template.py','r',encoding='utf8') as f1:
		text = f1.readlines()
		for line in text:
			if line.startswith("html_code") or line.endswith("#end"):
				continue
			with io.open('template.html','a',encoding='utf8') as f2:
				f2.write(line)

	# open template
	url = "file:///D:/mailchimp/template.html"
	webbrowser.open(url, new=2)

	time.sleep(4)
	# delete template instance
	os.remove("template.html")

# wrapper function to delete selected contacts
def del_selected():
	item = cont_tree.selection()
	del_contacts(item)

# wrapper function to delete all contacts
def del_all():
	item = cont_tree.get_children()
	del_contacts(item)

# function to delete contact
def del_contacts(item):
	del_email_tuple = []
	del_email_list = []
	for i in item:
		del_email_tuple.append(cont_tree.item(i,'values'[0]))

	for i in range(len(del_email_tuple)):
		del_email_list.append(del_email_tuple[i][1])
	if len(del_email_list)==0:
		messagebox.showinfo("No item selected","Select and item to proceed.")
		contacts_window.wm_deiconify()
		return
	elif len(del_email_list)==1:
		response = messagebox.askyesno("Are you sure?","Do you want to delete \""+del_email_list[0]+"\" from the audience? You won't be able to add it later. If you want to keep but prevent the contact from getting new emails, unsubscribe instead.")
	else:
		response = messagebox.askyesno("Are you sure?","Do you want to delete these contacts from the audience? You won't be able to add them later to the audience. If you want to keep but prevent them from getting new emails, unsubscribe instead.")
	contacts_window.wm_deiconify()
	if response:
		for email in del_email_list:
			errorcheck = mpf2.del_member(list_id=audience_id,email=email)
			if errorcheck == 'ok':
				contacts_window.wm_deiconify()
				return
			for x in item:
				cont_tree.delete(x)
		contacts_window.wm_deiconify()
	else:
		contacts_window.wm_deiconify()

def sub_contacts():
	items = cont_tree.selection()
	for item in items:
		values = cont_tree.item(item,'values')
		sr_no = values[0]
		email = values[1]
		status = "subscribed"

		response = mpf2.subscribe(list_id=audience_id,email=email)
		if response=='ok':
			continue
		cont_tree.item(item, text="", values=(sr_no, email, status))
	contacts_window.wm_deiconify()

def unsub_contacts():
	items = cont_tree.selection()
	for item in items:
		values = cont_tree.item(item,'values')
		sr_no = values[0]
		email = values[1]
		status = "unsubscribed"

		response = mpf2.unsubscribe(list_id=audience_id,email=email)
		if response=='ok':
			continue
		cont_tree.item(item, text="", values=(sr_no, email, status))
	contacts_window.wm_deiconify()

# function to add contact
def add_contacts():
	# root_disable()
	root.attributes('-disabled', True)
	global adder
	global cont_input_area
	# global scroll_y
	global open_file_text
	global browse_btn
	adder = Tk()
	adder.overrideredirect(True)
	adder.title('Add Contacts')
	adder.wm_iconbitmap('icon.ico')
	a_w,a_h = 400,550
	x = (w/2-a_w/2)
	y= (h/2-a_h/2)
	adder.geometry(f'{a_w}x{a_h}+{int(x)}+{int(y)}')
	adder.wm_deiconify()

	boundary = Frame(adder,bd=10,relief=RAISED, width=a_w,height=a_h)
	boundary.place(x=0,y=0)

	# radiobuttons
	type_rbtn = Radiobutton(adder, text='input contacts',variable=choice,value=1, font=('times new roman', 13, 'bold'),cursor='hand2',command=disable_import)
	type_rbtn.place(x=20,y=13)
	type_rbtn.select()

	import_rbtn = Radiobutton(adder, text='import contacts',variable=choice,value=2, font=('times new roman', 13, 'bold'),cursor='hand2',command=disable_type)
	import_rbtn.place(x=20,y=360)

	# contact input frame
	cont_input_frame =Frame(adder, bd=10, relief=GROOVE)
	cont_input_frame.place(x=20, y=40, width=360, height=300)

	scroll_y=Scrollbar(cont_input_frame,orient=VERTICAL)
	cont_input_area=Text(cont_input_frame,yscrollcommand=scroll_y.set,)
	scroll_y.pack(side=RIGHT,fill=Y)
	scroll_y.config(command=cont_input_area.yview)
	cont_input_area.pack(fill=BOTH,expand=1)
	cont_input_area.insert(1.0, 'example1@gmail.com\nexample2@gmail.com\nexample3@gmail.com\nexample4@gmail.com\nexample5@gmail.com')			# Placeholder
	cont_input_area.bind("<Button>",contText)

	lbl = Label(adder, text="or",font=('arial',11,'italic'),fg='red')
	lbl.place(x=183,y=345)

	open_file_text=Entry(adder,width=35,font=('arial',10),bd=2,state="disabled")
	open_file_text.place(x=20,y=400)

	# browse button
	browse_btn=Button(adder,text="Browse",command=browse,width=10,bd=4,font=('arial',9),state="disabled")
	browse_btn.place(x=285,y=395)

	# add button
	add_btn=Button(adder,text='Add Contacts',command=append_cont, bg=bg_color,fg='white', width=12,bd=3,
	font=('arial', 9, 'bold')).place(x=150,y=460)

	close_btn=Button(adder,text='Close',bg=bg_color,fg='white',pady=0,bd=3,
						 font=('arial', 8, 'bold'),command=quit_add)
	close_btn.place(x=325,y=500)

# function to show contact
def show_contacts():
	global cont_tree
	global contacts_window
	email_list = mpf2.members_info(audience_id)
	if email_list=='ok':
		return
	email_list_sorted = sorted(email_list)

	root.attributes('-disabled', True)                    # disable the root window

	contacts_window = Tk()
	contacts_window.title('Contacts')
	contacts_window.wm_iconbitmap('icon.ico')
	show_w = 500
	show_h = 525
	x = (w/2)-(show_w/2)
	y = (h/2)-(show_h/2)
	contacts_window.geometry(f'{show_w}x{show_h}+{int(x)}+{int(y)}')
	contacts_window.overrideredirect(True)

	# border frame
	boundary = Frame(contacts_window,bd=10,relief=RAISED, width=show_w,height=show_h)
	boundary.place(x=0,y=0)

	# treeview frame
	treeview_frame = LabelFrame(contacts_window, bd=10, relief=GROOVE, text="Contact list", font=('times new roman',13,'bold'))
	treeview_frame.place(x=20,y=20)

	#treeview scrollbar
	tree_scroll = Scrollbar(treeview_frame)
	tree_scroll.pack(side=RIGHT,fill=Y)

	# create treeview
	cont_tree = ttk.Treeview(treeview_frame,yscrollcommand=tree_scroll.set)
	cont_tree.pack(fill=BOTH,expand=1)

	# configure the scrollbar
	tree_scroll.config(command=cont_tree.yview)

	# define columns
	cont_tree['columns'] = ("","Email","Status")

	# format columns
	cont_tree.column("#0",width=0, stretch=NO)
	cont_tree.column("",anchor=CENTER, width=40,minwidth=30)
	cont_tree.column("Email",anchor=W,width=250, minwidth=180)
	cont_tree.column("Status",anchor=CENTER,width=130,minwidth=90)

	# Create heading
	cont_tree.heading("#0",text="")
	cont_tree.heading("",text="",anchor=CENTER)
	cont_tree.heading("Email", text="Email", anchor=W)
	cont_tree.heading("Status", text="Status", anchor=CENTER)

	# add data
	count = 0
	for email in email_list_sorted:
		cont_tree.insert(parent='',index='end',iid=count, text="", values=(count+1,email,email_list[email]))
		count+=1

	# contact function frame
	cont_func_frame=LabelFrame(contacts_window,bd=10,text="Actions to perform",font=('times new roman',13,'bold'))
	cont_func_frame.place(x=20,y=295,width=460,height=180)

	sub_btn=Button(contacts_window,text='Subscribe', bg=bg_color,fg='white', command=sub_contacts,width=12,height=2,bd=3,
	font=('arial', 10, 'bold')).place(x=95,y=330)

	unsub_btn=Button(contacts_window,text='Unsubscribe', bg=bg_color,fg='white', command= unsub_contacts, width=12,height=2,bd=3,
	font=('arial', 10, 'bold')).place(x=295,y=330)

	del_btn=Button(contacts_window,text='Delete', bg=bg_color,fg='white', command=del_selected, width=12,height=2,bd=3,
	font=('arial', 10, 'bold')).place(x=95,y=390)

	del_all_btn=Button(contacts_window,text='Delete All', bg=bg_color,fg='white', command=del_all, width=12,height=2,bd=3,
	font=('arial', 10, 'bold')).place(x=295,y=390)

	close_btn=Button(contacts_window,text='Close',bg=bg_color,fg='white', command=quit_show,pady=0,bd=3,
						 font=('arial', 8, 'bold'))
	close_btn.place(x=438,y=483)

	
	'''# contacts showing textbox
				cont_showing_frame =Frame(contacts_window, bd=10, relief=GROOVE)
				cont_showing_frame.place(x=20, y=40, width=360, height=350)
			
				scroll_y=Scrollbar(cont_showing_frame,orient=VERTICAL)
				cont_showing_area=Text(cont_showing_frame,yscrollcommand=scroll_y.set,)
				scroll_y.pack(side=RIGHT,fill=Y)
				scroll_y.config(command=cont_showing_area.yview)
				cont_showing_area.pack(fill=BOTH,expand=1)
			
				# insert all the emails in text box
				counter = 0
				for email in response:
					counter = counter + 1
					cont_showing_area.insert(END, str(counter) + ". " + email + os.linesep)
			
				cont_showing_area.configure(state='disabled')
			'''

# preview text function
def add_preview_text():
	global preview_text_win
	global prev_textarea
	root.attributes('-disabled', True)
	preview_text_win = Tk()
	preview_text_win.overrideredirect(True)
	preview_text_win.title('Add Preview Text')
	preview_text_win.wm_iconbitmap('icon.ico')
	p_w,p_h = 500,310
	x = (w/2-p_w/2)
	y= (h/2-p_h/2)
	preview_text_win.geometry(f'{p_w}x{p_h}+{int(x)}+{int(y)}')
	preview_text_win.focus_force()

	# border frame
	boundary = Frame(preview_text_win,bd=10,relief=RAISED, width=p_w,height=p_h)
	boundary.place(x=0,y=0)

	# preview text input area
	preview_text_frame =Frame(preview_text_win, bd=10, relief=GROOVE)
	preview_text_frame.place(x=30, y=30, width=440, height=200)

	preview_text_title=Label(preview_text_frame,text='Preview Text', font=('times new roman', 14, 'bold'),bd=7,relief=GROOVE).pack(fill=X)

	scroll_y=Scrollbar(preview_text_frame,orient=VERTICAL)
	prev_textarea=Text(preview_text_frame,font=('calibri', 11),yscrollcommand=scroll_y.set,)
	scroll_y.pack(side=RIGHT,fill=Y)
	scroll_y.config(command=prev_textarea.yview)
	prev_textarea.pack(fill=BOTH,expand=1)

	prev_textarea.insert(1.0, preview_text)
	prev_textarea.focus()

	# add button
	add_text_btn=Button(preview_text_win,text='Add',command=add_text, bg=bg_color,fg='white', width=12,bd=3,
	font=('arial', 9, 'bold')).place(x=195,y=240)

	close_btn=Button(preview_text_win,text='Close',bg=bg_color,fg='white',pady=0,bd=3,
						 font=('arial', 8, 'bold'),command=quit_ptext)
	close_btn.place(x=440,y=270)

# function to send mail
def send():
	global preview_text
	# =============================================================================
	# campaign creation
	# =============================================================================
	campaign_name = cname_text.get()
	from_name = name_text.get()
	reply_to = rto_text.get()

	campaign = mpf1.campaign_creation_function(campaign_name=campaign_name,
                                      audience_id=audience_id,
                                      from_name=from_name,
                                      reply_to=reply_to,
                                      preview_text=preview_text)
	if campaign=="ok":
		return

	# =============================================================================
	# template creation
	# =============================================================================

	html_code = newsletter_template.html_code           

	response = mpf1.customized_template(html_code=html_code, campaign_id=campaign['id'])

	if response=="ok":
		return

	# =============================================================================
	# send the mail campaign
	# =============================================================================

	response = mpf1.send_mail(campaign_id=campaign['id'])
	if response=="ok":
		return

	preview_text = ""
	print("Emails sent")

# functions for social platforms
# function to open facebook profile
def social_fb():
	webbrowser.open("https://www.facebook.com/profile.php?id=100009259126329", new=2)

# function to open instagram profile
def social_insta():
	webbrowser.open("https://instagram.com/ashutosh_rai.05?igshid=170b58r9gjasq", new=2)

# function to open lnkedin profile
def social_in():
	webbrowser.open("www.linkedin.com/in/ashutoshrai142", new=2)

# function to open github profile
def social_git():
	webbrowser.open("https://github.com/cosmicbanana8342", new=2)

# ==============================================================================================================
# Driver code
# ==============================================================================================================

title=Label(root,text="E-Mail App",font=('times new roman',30,'bold'),bg=bg_color,fg='gold',bd=12,relief=GROOVE,pady=2).pack(fill=X)

# campaign name frame
cname_frame=LabelFrame(root,bd=10,relief=GROOVE,font=('times new roman',15,'bold'),bg=bg_color,fg='gold')
cname_frame.place(x=10,y=80,width=740,height=90)

cname_label=Label(cname_frame,text='Campaign Name',font=('times new roman',18,'bold'),bg=bg_color,fg='white').grid(row=0,column=0,padx=20,pady=2)
cname_text=Entry(cname_frame,width=30,font=('arial',15),bd=7,relief=SUNKEN)
cname_text.focus()
cname_text.grid(row=0,column=1,pady=15,padx=10)

# from name frame
name_frame=LabelFrame(root,bd=10,relief=GROOVE,font=('times new roman',15,'bold'),bg=bg_color,fg='gold')
name_frame.place(x=10,y=180,width=740,height=90)

name_label=Label(name_frame,text='Name',font=('times new roman',18,'bold'),bg=bg_color,fg='white').grid(row=0,column=0,padx=20,pady=2)
name_text=Entry(name_frame,width=30,font=('arial',15),bd=7,relief=SUNKEN)
name_text.insert(0, 'Ashutosh Rai')			# Placeholder
name_text.bind("<Button>",nameText)
name_text.grid(row=0,column=1,pady=15,padx=121)

# reply to frame
rto_frame=LabelFrame(root,bd=10,relief=GROOVE,font=('times new roman',15,'bold'),bg=bg_color,fg='gold')
rto_frame.place(x=10,y=280,width=740,height=90)

rto_label=Label(rto_frame,text='Reply To',font=('times new roman',18,'bold'),bg=bg_color,fg='white').grid(row=0,column=0,padx=20,pady=2)
rto_text=Entry(rto_frame,width=30,font=('arial',15),bd=7,relief=SUNKEN)
rto_text.insert(0, 'ashutoshrai403@gmail.com')			# Placeholder
rto_text.bind("<Button>",rtoText)
rto_text.grid(row=0,column=1,pady=15,padx=85)

# contact buttons frame
contacts_frame=LabelFrame(root,bd=10,relief=GROOVE,font=('times new roman',15,'bold'),bg=bg_color,fg='gold')
contacts_frame.place(x=10,y=380,width=740,height=90)

# add contacts button
add_cont_btn=Button(contacts_frame,text="Add Contacts",command=add_contacts,width=30,bd=7,font=('arial',12,'bold'))
add_cont_btn.grid(row=0,column=0,padx=18,pady=12)

# show contacts button
show_cont_btn=Button(contacts_frame,text="Show Contacts",command=show_contacts,width=30,bd=7,font=('arial',12,'bold'))
show_cont_btn.grid(row=0,column=1,padx=32,pady=12)

# preview text frame
prev_text_frame=LabelFrame(root,bd=10,relief=GROOVE,font=('times new roman',15,'bold'),bg=bg_color,fg='gold')
prev_text_frame.place(x=10,y=480,width=740,height=90)

# preview text button
preview_text_btn=Button(prev_text_frame,text="Add Preview Text",command=add_preview_text,width=30,bd=7,font=('arial',12,'bold'))
preview_text_btn.pack(pady=12)

# send button
send_btn=Button(root,text='Send',bg=bg_color,fg='white',pady=15, width=10,bd=3,
						 font=('arial', 15, 'bold'),command=send)
send_btn.place(x=300,y=575)

# credits frame
credits_frame=LabelFrame(root,bd=10,relief=GROOVE,font=('times new roman',15,'bold'),bg=bg_color,fg='gold')
credits_frame.place(x=10,y=653,width=740,height=55)

# credit=Label(credits_frame,text="Developed by Ashutosh Rai",font=('times new roman',20,'bold'),bg=bg_color,fg='gold',bd=12,relief=GROOVE,pady=2).pack(fill=X)

# facebook button
fb = PhotoImage(file='social/fb.png')
fb_btn = Button(root,text='facebook',image=fb,command=social_fb,bd=3,relief=GROOVE,pady=0)
fb_btn.place(x=195,y=665)

# instagram button
insta = PhotoImage(file='social/insta.png')
insta_btn = Button(root,text='instagram',image=insta,command=social_insta,bd=3,relief=GROOVE,pady=0)
insta_btn.place(x=295,y=665)

# linkedin button
linkedin = PhotoImage(file='social/in.png')
in_btn = Button(root,text='linkedin',image=linkedin,command=social_in,bd=3,relief=GROOVE,pady=0)
in_btn.place(x=395,y=665)

# github button
github = PhotoImage(file='social/github.png')
github_btn = Button(root,text='github',image=github,command=social_git,bd=3,relief=GROOVE,pady=0)
github_btn.place(x=495,y=665)

# template code area
template_frame =Frame(root, bd=10, relief=GROOVE)
template_frame.place(x=760, y=80, width=592, height=530)

temp_title=Label(template_frame,text='Template', font=('times new roman', 16, 'bold'),bd=7,relief=GROOVE).pack(fill=X)

scroll_y=Scrollbar(template_frame,orient=VERTICAL)
textarea=Text(template_frame,yscrollcommand=scroll_y.set,)
scroll_y.pack(side=RIGHT,fill=Y)
scroll_y.config(command=textarea.yview)
textarea.pack(fill=BOTH,expand=1)

# template buttons frame
temp_frame=LabelFrame(root,bd=10,relief=GROOVE,font=('times new roman',15,'bold'),bg=bg_color,fg='gold')
temp_frame.place(x=760,y=618,width=585,height=90)

# show template button
show_temp_btn=Button(temp_frame,text="Show Rendered Template",command=show_template,width=20,bd=7,font=('arial',12,'bold'))
show_temp_btn.pack(padx=30, side=LEFT)

# render template button
render_temp_btn=Button(temp_frame,text="Render Template",command=render_temp,width=20,bd=7,font=('arial',12,'bold'))
render_temp_btn.pack(padx=(0,30),side=RIGHT)

root.mainloop()