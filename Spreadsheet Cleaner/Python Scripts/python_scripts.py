'''
The following script is used to process MPL files. The order of tasks that are actioned in this file are:

-Loading needed files
**User will have reference of how old Commodities and Programs files are in case they want to update them

-Validating if MPL template was used
-Printing --> Shape of unedited MPL file
-Removal of duplicate APN/Program combos
-Printing --> Count of duplicates removed
-Update columns of MPL file if proper format was not followed
-Printing --> New shape of MPL file
                        -Validate Commodities referenced in MPL
                            -If there is non existent Commoditiy being referenced, ask user to provide what the Commodity should be or to be left as is
                        -Update Commodity case sensitivity/spelling to match what is in Ninja
                        -Validate Programs referenced in MPL
                            -Update hero programs with O_
                            -Update non-here programs to ensure they do not have O_
                            -If there is a non existent Program being referenced, ask user to provide what the Program should be or to be left as is
                        -Update Program case sensitivity/spelling to match what is in Ninja

                        ***Add coloring to rows which were modified in some way throughtout this script in "Light Blue"
                        ***Add coloring to rows which have possibility of error when processing MPL in "Red"

-Export cleaned MPL file to user's desired name and location
'''



import tkinter as tk
from tkinter.constants import ACTIVE, DISABLED, END, NORMAL, TRUE
from tkinter.font import BOLD
from typing import Text
import pandas as pd
import numpy as np
import os
import sys
import time
from tkinter import Button, Canvas, Toplevel, filedialog, Entry
from tkinter.simpledialog import askstring
from UliPlot.XLSX import auto_adjust_xlsx_column_width
from fuzzywuzzy import process
from pandas.io.formats.style_render import Subset


# locating current directory to locate commodity and programs files needed
script_location = os.path.realpath(__file__)

# pointing at References folder directory and away from script
References_folder = script_location[:-(len(os.path.basename(sys.argv[0])))] + "References"

# listing out files located in References folder since file names may change over time
filesinfolder = os.listdir(References_folder)

logo = ""

# gathering the commodities file from files found in References folder
for file in filesinfolder:
    if "turtle" in file.lower():
        logo = References_folder+ "/" + file

# creating gui for app
app_window = tk.Tk()
app_window.title("MPL Cleaner")
photo = tk.PhotoImage(file=logo)
app_window.iconphoto(True, photo)



# specifying the app window size
app_w = 800
app_h = 600

# gathering users display size to center app window
ws = app_window.winfo_screenwidth()
hs = app_window.winfo_screenheight()

# determining coordinates for the app to open at on user display
app_x = (ws/2) - (app_w/2)
app_y = (hs/2) - (app_h/2)

# centering the app window location and providing app window size
app_window.geometry(f"{app_w}x{app_h}+{int(app_x)}+{int(app_y)}")
app_window.resizable(False,False)



# variables for files that will be needed for script. (Filepaths gathered later on)
# if script and other files needed won't be together, please add locations for files and remove search functions below for commodity/programs files
commodities_reference = "" #location reference
programs_reference = "" #location reference
commodity_file = "" # reading commodity file
programs_file = "" # reading programs file

uncleaned_mpl = "" #location reference
just_filename_uncleaned_mpl = "" #only the filename which will be modified for new clean file name default
mpl = "" #dataframe of unclean MPL
reduced = "" #dataframe after duplicates removed from MPL
cleaned_mpl = "" #dataframe of cleaned MPL
cleaned_mpl_with_changes = "" #dataframe with records of things that changed
name_of_new_cleaned_mpl =  "" #Name of file that will cleaned MPL will be exported with by default

# referencing steps taken to prevent error and track progress
step_count = 0

# determine if loaded MPL file processable by script. check_unclean_file function will be used for this
# zero means bad file loaded, one means good file. Values captured when script is ran
file_validity = ""

# index references of correct columns needed (program* , parts* , ops finance commodity* )
# values will auto populate when script ran
program_col = ""
parts_col = ""
commodity_col = ""

# confirmation to proceed with script if template isn't matching good
# values will auto populate when script ran
continue_regardless = ""

# Columns names that MPL file should have
mpl_cols = ['Program*','Site Group','Site Building','Part*',
            'Part Description','Procurement Commodity','Ops Finance Commodity*',
            'Cost','Supplier Description','Supplier Code','GSM DRI Name','GSM DRI Email',
            'Part Tier','FG PO FLAG']


# gathering the commodities file from files found in References folder
for file in filesinfolder:
    if "commodity" in file.lower():
        commodities_reference = References_folder+ "/" + file
        commodity_file = pd.read_excel(commodities_reference)
        break
    else:
        commodities_reference = "Not commodities file found"

# getting time commodities file was last updated (not the same as last time file opened)
# commodities_reference_last_update = ""
if commodities_reference == "Not commodities file found" or commodities_reference == "":
    commodities_reference_last_update = "No commodities file found"
else:
    commodities_reference_last_update = time.strftime("%m/%d/%Y %I:%M:%S %p",time.localtime(os.path.getmtime(commodities_reference)))


# gathering the Programs file from files found in References folder
for file in filesinfolder:
    if "program" in file.lower():
        programs_reference = References_folder+ "/" + file
        programs_file = pd.read_excel(programs_reference)
        break
    else:
        programs_reference = "No programs file found"

# getting time Programs file was last updated (not the same as last time file opened)
programs_reference_last_update = ""
if programs_reference == "No programs file found" or programs_reference == "":
    programs_reference_last_update = "No programs file found"
else:
    programs_reference_last_update = time.strftime("%m/%d/%Y %I:%M:%S %p",time.localtime(os.path.getmtime(programs_reference)))



# function to get user input if file columns are in required spots
def get_confirmation():
    global continue_regardless

    continue_regardless = ""
    button_selected = tk.IntVar()

    def user_yes():
        global continue_regardless
        continue_regardless = "Yes"
        button_selected.set(1)
        top_window.destroy()


    def user_no():
        global continue_regardless
        continue_regardless = "No"
        button_selected.set(1)
        top_window.destroy()

    top_x = (ws/2) - (150/2)
    top_y = (hs/2) - (250/2)
    top_window = Toplevel()
    top_window.title("Please Confirm")
    top_window.geometry(f"250x150+{int(top_x)}+{int(top_y)}")
    label16 = tk.Label(top_window, text=
    '''Could not locate columns needed
        are the following true:

    -Programs are on Column A
    -Parts are on Column D
    -Commodities are on Column G

    ''')
    label16.pack()


    topbtn = tk.Button(top_window, text="Yes", command = (lambda:user_yes()))
    topbtn.pack()
    topbtn.place(rely=.7,relx=.25)
    

    topbtn2 = tk.Button(top_window, text="No", command = (lambda:user_no()))
    topbtn2.pack()
    topbtn2.place(rely=.7,relx=.5)
    top_window.wait_variable(button_selected)






# function used to check file validity
def check_unclean_file():
    global program_col
    global parts_col
    global commodity_col
    global file_validity
    global label4
    global label4v2

    getting_columns = mpl.columns.tolist()

    if getting_columns == mpl_cols:
        label3 = tk.Label(frame, text= "\nProper MPL Template:" ,anchor='s',font=('Helvetica', 12, BOLD ))
        label3v2 = tk.Label(frame, text= "Was Used" ,anchor='s')
        label3.pack()
        label3v2.pack()
        program_col = 0
        parts_col = 3
        commodity_col = 6
        file_validity = 1
    
    elif len(getting_columns) < 7:
        file_validity = 0
        label4 = tk.Label(frame, text= "\nProper MPL Template:" ,anchor='s',font=('Helvetica', 12, BOLD ))
        label4v2 = tk.Label(frame, text= "Not Used" ,anchor='s')
        label4.pack()
        label4v2.pack()

    else:

        while (parts_col == "") or (parts_col == "") or (parts_col == ""):

            for col in getting_columns:
                if program_col == "":
                    if col.lower() == 'Program*':
                        program_col =  getting_columns.index(col)
                    elif (process.extractOne('Program*', getting_columns)[1] >= 90):
                        program_col = getting_columns.index(process.extractOne('Program*', getting_columns)[0])

                if parts_col == "":
                    if col.lower() == 'Part*':
                        parts_col =  getting_columns.index(col)
                    elif (process.extractOne('Part*', getting_columns)[1] >= 90):
                        parts_col = getting_columns.index(process.extractOne('Part*', getting_columns)[0])

                if commodity_col == "":
                    if col.lower() == 'Ops Finance Commodity*':
                        commodity_col =  getting_columns.index(col)
                    elif (process.extractOne('Ops Finance Commodity*', getting_columns)[1] >= 90):
                        commodity_col = getting_columns.index(process.extractOne('Ops Finance Commodity*', getting_columns)[0])

            if (program_col == "") or (parts_col == "") or (commodity_col == ""):
                get_confirmation()

                print("I got to this point")
                print(continue_regardless)
                print("This is after supposedly printing the value from pop up")
                if continue_regardless.lower() in ['yes','yea','y']:
                    print('Perfect, lets get started....\n')
                    # time.sleep(3) # Sleep for 3 seconds
                    program_col = 0
                    parts_col = 3
                    commodity_col = 6
                    file_validity = 1
                else:
                    print(continue_regardless)
                    print("Unable to locate necessary columns")
                    file_validity = 0
                    break
            # else:
            #     print("MPL Template File was not used\n")
            #     file_validity = 1
        label4 = tk.Label(frame, text= "\nProper MPL Template:" ,anchor='s',font=('Helvetica', 12, BOLD ))
        label4v2 = tk.Label(frame, text= "Not Used" ,anchor='s')
        label4.pack()
        label4v2.pack()





# To ask for excel file
def openfile():
    global uncleaned_mpl #we may not need this global variable, pending review.......
    global mpl
    global just_filename_uncleaned_mpl
    global step_count
    global name_of_new_cleaned_mpl

    filepath= filedialog.askopenfilename(title = "Provide MPL file to prepare", filetypes= [("Excel files", ".xlsx .xls")])
    uncleaned_mpl = filepath
    mpl = pd.read_excel(uncleaned_mpl) # Reading MPL file provided

    check_unclean_file()

    if (filepath is None) or (file_validity == 0):
        
        label5 = tk.Label(frame, text= "\nFile provided was invalid, please select a valid MPL file to work on" ,anchor='s')
        # uncleaned_mpl = "File provided was invalid, please select a valid MPL file to work on"
        # label4 = tk.Label(frame, text= f"\nFile being worked on:\n{uncleaned_mpl.rpartition('/')[-1]}" ,anchor='s')

        #making load button inactive to have user reset process
        button['state'] = DISABLED

        # activating process another MPL in case user changes mind on file selected
        button5['state']= NORMAL

        label5.pack()
        label4.after(20000, label4.destroy)
        label4v2.after(20000, label4.destroy)
        label5.after(20000, label5.destroy)


    else:
        # uncleaned_mpl = filepath
        # mpl = pd.read_excel(uncleaned_mpl) # Reading MPL file provided
        # just_filename_uncleaned_mpl = f'{os.path.splitext(uncleaned_mpl)[0]}_CLEANED{os.path.splitext(uncleaned_mpl)[1]}'
        just_filename_uncleaned_mpl = uncleaned_mpl.rpartition('/')[-1]
        name_of_new_cleaned_mpl = os.path.splitext(just_filename_uncleaned_mpl)[0]+ "_CLEANED"
        print(just_filename_uncleaned_mpl)
        print(name_of_new_cleaned_mpl)
        step_count +=1

        #making load button inactive since successful load of unclean MPL (avoids confusion)
        button['state'] = DISABLED
        
        # activating clean MPL button
        button2['state']= NORMAL

        # activating process another MPL in case user changes mind on file selected
        button5['state']= NORMAL

        # creating label reference of what file is being worked on
        label6 = tk.Label(frame, text= f"\nFile being worked on:" ,anchor='s',font=('Helvetica', 12, BOLD ))
        label6v2 = tk.Label(frame, text= just_filename_uncleaned_mpl ,anchor='s')
        label6.pack()
        label6v2.pack()



# removing duplicate rows before the clean up and providing reference of quantity removed
def remove_duplicates():
    global reduced
    # show count of rows prior to removal of APN-Program duplicates
    print('Total rows in file:', mpl.shape[0])
    label7= tk.Label(frame, text ='\n Total rows in file:', anchor='s',font=('Helvetica', 12, BOLD))
    label7v2= tk.Label(frame, text = "{:,}".format(mpl.shape[0]), anchor='s' )
    label7.pack()
    label7v2.pack()
    

    if (mpl.duplicated(subset=[(mpl.columns[program_col]), (mpl.columns[parts_col])]).sum()) > 0:
        # notifying of duplicates that will be removed
        print('---> Removing', mpl.duplicated(subset=[(mpl.columns[program_col]), (mpl.columns[parts_col])]).sum(), 'duplicates\n')
        label8= tk.Label(frame, text ='---> Removed '+ "{:,}".format(mpl.duplicated(subset=[(mpl.columns[program_col]), (mpl.columns[parts_col])]).sum())+ ' duplicates', anchor='s' )
        label8.pack()

        # remove duplicates based on several columns
        reduced = mpl.drop_duplicates(subset=[(mpl.columns[program_col]), (mpl.columns[parts_col])], keep ="last")
        print('\nTotal rows after removing duplicates:', reduced.shape[0])
        label9= tk.Label(frame, text ='\nTotal rows after removing duplicates:', anchor='s',font=('Helvetica', 12, BOLD))
        label9v2= tk.Label(frame, text = str(reduced.shape[0]), anchor='s' )
        label9.pack()
        label9v2.pack()
    else:
        print ("\nNo duplicates rows in file")
        label10= tk.Label(frame, text ="No duplicates rows in file", anchor='s' )
        label10.pack()
        reduced = mpl


# removing duplicate rows after clean up was done and providing reference of quantity removed
def remove_duplicates_after_cleanup():
    global cleaned_mpl_with_changes

    # show count of rows prior to removal of APN-Program duplicates
    print('Total rows in file after clean:', cleaned_mpl_with_changes.shape[0])
    y = cleaned_mpl_with_changes.shape[0]
    if (cleaned_mpl_with_changes.duplicated(subset=["Program*", "Part*"]).sum()) > 0:
        # notifying of duplicates that will be removed
        print('---> Removing', (cleaned_mpl_with_changes.duplicated(subset=["Program*", "Part*"]).sum()), 'duplicates\n')
        x = cleaned_mpl_with_changes.duplicated(subset=["Program*", "Part*"]).sum()
        # remove duplicates based on several columns
        cleaned_mpl_with_changes = cleaned_mpl_with_changes.drop_duplicates(subset=["Program*", "Part*"], keep ="last")
        print('\nTotal rows after removing duplicates:', cleaned_mpl_with_changes.shape[0])
        label17= tk.Label(frame, text = f"An addditional {x} duplicate/s were found after clean up and have been removed leaving you with {y-x} rows")
        label17.pack()

    else:
        print ("\nNo duplicates rows in file")



# function to return programs that were changed
def returning_changed_programs(word):
    programs_list = programs_file["Program"].unique().tolist()

    if process.extractOne(word,programs_list)[1] == 100:
        if process.extractOne(word,programs_list)[0] == word:
            return np.nan #sending blank since it matched exactly
        else:
            return word
    elif process.extractOne(word,programs_list)[1] >= 95:
        return word
    elif process.extractOne("O_"+word,programs_list)[1] >= 95:
        return word
    elif len(word[2:]) > 1:
        if process.extractOne(word[2:],programs_list)[1] >= 95:
            return word
        else:
            return np.nan
    else:
        return np.nan


# function to clean up programs
def comparing_programs(word):
    programs_list = programs_file["Program"].unique().tolist()

    if process.extractOne(word,programs_list)[1] == 100:
        return process.extractOne(word,programs_list)[0]
    elif process.extractOne(word,programs_list)[1] >= 95:
        return process.extractOne(word,programs_list)[0]
    elif process.extractOne("O_"+word,programs_list)[1] >= 95:
        return process.extractOne(word,programs_list)[0]
    elif len(word[2:]) > 1:
        if process.extractOne(word[2:],programs_list)[1] >= 95:
            return process.extractOne(word,programs_list)[0]
        else:
            return word
    else:
        return word

# function to return programs that were changed
def returning_changed_commodities(word):
    commodities_list = commodity_file["Commodity"].unique().tolist()
    sub_commodities_list = commodity_file["Sub-Commodity"].unique().tolist()

    if process.extractOne(word,commodities_list)[1] == 100:
        if process.extractOne(word,commodities_list)[0] == word:
            return np.nan #sending blank since it matched exactly
        else:
            return word
    elif process.extractOne(word,commodities_list)[1] >= 95:
        return word
    else:
        if word not in ["Uncategorized", "AppleCare Only", ""]:
            if process.extractOne(word,sub_commodities_list)[1] >= 89:
                match = process.extractOne(word,sub_commodities_list)[0]
                return word
            else:
                return np.nan
        else:
            return np.nan

# function to clean up commodities
def comparing_commodities(word):
    commodities_list = commodity_file["Commodity"].unique().tolist()
    sub_commodities_list = commodity_file["Sub-Commodity"].unique().tolist()

    if process.extractOne(word,commodities_list)[1] == 100:
        return process.extractOne(word,commodities_list)[0]
    elif process.extractOne(word,commodities_list)[1] >= 95:
        return process.extractOne(word,commodities_list)[0]
    
    else:
        if word not in ["Uncategorized", "AppleCare Only", ""]:
            if process.extractOne(word,sub_commodities_list)[1] >= 89:
                match = process.extractOne(word,sub_commodities_list)[0]
                return commodity_file['Commodity'][commodity_file['Sub-Commodity'][commodity_file['Sub-Commodity'] == match].index.tolist()[0]]
            else:
                return word
        else:
            return word



# main function to process unclean MPL file
def process_mpl():
    global cleaned_mpl_with_changes
    global cleaned_mpl

    remove_duplicates()

    data_new = []

    try:
        for index, row in reduced.iterrows():
            print(index)
            dict_data = {
                "Original Program if Changed": returning_changed_programs(row[reduced.columns.tolist()[program_col]]),### work in progress to only get values which have a change
                "Program*": comparing_programs(row[reduced.columns.tolist()[program_col]]),
                "Site Group": np.nan,
                "Site Building": np.nan,
                "Part*": row[reduced.columns.tolist()[parts_col]],
                "Part Description": np.nan,
                "Procurement Commodity": np.nan,
                "Original Commodity if Changed": returning_changed_commodities(row[reduced.columns.tolist()[commodity_col]]),### work in progress to only get values which have a change
                "Ops Finance Commodity*": comparing_commodities(row[reduced.columns.tolist()[commodity_col]]),
                "Cost": np.nan,
                "Supplier Description": np.nan,
                "Supplier Code": np.nan,
                "GSM DRI Name": np.nan,
                "GSM DRI Email": np.nan,
                "Part Tier": np.nan,
                "FG PO FLAG":np.nan
            }

            data_new.append(dict_data)

        cleaned_mpl_with_changes = pd.DataFrame(data=data_new)

        remove_duplicates_after_cleanup()

        cleaned_mpl = cleaned_mpl_with_changes.drop(columns=["Original Program if Changed", "Original Commodity if Changed"])

        #making clean MPL button deactive since clean successful (avoids confusion)
        button2['state'] = DISABLED
        
        # activating export clean MPL button
        button3['state']= NORMAL

        # activating export clean MPL with changes button
        button4['state']= NORMAL

        label14= tk.Label(frame, text ='\n Total Programs that were updated:', anchor='s',font=('Helvetica', 12, BOLD))
        label14v2= tk.Label(frame, text = "{:,}".format(cleaned_mpl_with_changes['Original Program if Changed'].count()), anchor='s' )
        label14.pack()
        label14v2.pack()

        label15= tk.Label(frame, text ='\n Total Commodities that were updated:', anchor='s',font=('Helvetica', 12, BOLD))
        label15v2= tk.Label(frame, text = "{:,}".format(cleaned_mpl_with_changes['Original Commodity if Changed'].count())+ "\n", anchor='s' )
        label15.pack()
        label15v2.pack()


        # # statement to save finalized excel file and space out columns
        # with pd.ExcelWriter(root+new_file) as writer:
        #     cleaned_mpl.to_excel(writer, sheet_name="Sheet1", index=0)
        #     auto_adjust_xlsx_column_width(cleaned_mpl, writer, sheet_name="Sheet1", index=0)#adjusting column sizing
        # print("\nCleaned MPL file successfully saved")


    except:
        print("An exception occurred")
        label11 = tk.Label(frame, text= "\nUnable to finalize MPL clean up \nFile provided may be missing values or have incorrect values \n Please make adjustments or select diffeerent file" ,anchor='s')
        label11.pack()


        #making clean MPL button deactive since clean successful (avoids confusion)
        button2['state'] = DISABLED
        
        # activating export clean MPL button
        button3['state']= DISABLED

        # activating export clean MPL with changes button
        button4['state']= DISABLED

        # activating process another MPL for user to start over
        button5['state']= NORMAL







    




#tester function to check things are working
# ERASE ONCE DONE WITH THIS
def printfile():
    print (uncleaned_mpl )
    label = tk.Label(frame, text= f"Your file has successfully exported" ,anchor='s')
    label.pack()

    #making load button inactive since successful load of unclean MPL (avoids confusion)
    button3['state'] = DISABLED




# To save excel file
def savefile():
    newfilename= f'{os.path.splitext(just_filename_uncleaned_mpl)[0]}_Cleaned{os.path.splitext(just_filename_uncleaned_mpl)[1]}'
    filepath2= filedialog.asksaveasfilename(title="Where to save cleaned MPL",defaultextension=".xlsx", initialfile= newfilename)
   
    with pd.ExcelWriter(filepath2) as writer:
        cleaned_mpl.to_excel(writer, sheet_name="Sheet1", index=0)
        auto_adjust_xlsx_column_width(cleaned_mpl, writer, sheet_name="Sheet1", index=0)#adjusting column sizing
    
    label12 = tk.Label(frame, text= "Cleaned MPL file has successfully exported" ,anchor='s',font=('Helvetica', 12, BOLD))
    label12.pack()

    # deactivating export clean MPL button since action successfully completed
    button3['state']= DISABLED

    if filepath2 is None:
        return
    else:
        print (filepath2)


# To save excel file with references to original data which changed
def savefile_w_references():
    newfilename= f'{os.path.splitext(just_filename_uncleaned_mpl)[0]}_Cleaned_w_Changes{os.path.splitext(just_filename_uncleaned_mpl)[1]}'
    filepath2= filedialog.asksaveasfilename(title="Where to save cleaned MPL",defaultextension=".xlsx", initialfile= newfilename)
   
    with pd.ExcelWriter(filepath2) as writer:
        cleaned_mpl_with_changes.to_excel(writer, sheet_name="Sheet1", index=0)
        auto_adjust_xlsx_column_width(cleaned_mpl_with_changes, writer, sheet_name="Sheet1", index=0)#adjusting column sizing

    label13 = tk.Label(frame, text= "Cleaned MPL Referencing changes has successfully exported" ,anchor='s',font=('Helvetica', 12, BOLD))
    label13.pack()

    # deactivating export clean MPL with changes button
    button4['state']= DISABLED

    if filepath2 is None:
        return
    else:
        print (filepath2)


# To reset variables
def resetAll():
    global uncleaned_mpl
    global just_filename_uncleaned_mpl
    global mpl
    global reduced
    global cleaned_mpl
    global cleaned_mpl_with_changes
    global step_count
    global file_validity
    global program_col
    global parts_col
    global commodity_col
    global continue_regardless

    uncleaned_mpl = "" #location reference
    just_filename_uncleaned_mpl = "" #only the filename which will be modified for new clean file name default
    mpl = "" #dataframe of unclean MPL
    reduced = "" #dataframe after duplicates removed from MPL
    cleaned_mpl = "" #dataframe of cleaned MPL
    cleaned_mpl_with_changes = "" #dataframe with records of things that changed

    # referencing steps taken to prevent error and track progress
    step_count = 0

    # determine if loaded MPL file processable by script. check_unclean_file function will be used for this
    # zero means bad file loaded, one means good file. Values captured when script is ran
    file_validity = ""

    # index references of correct columns needed (program* , parts* , ops finance commodity* )
    # values will auto populate when script ran
    program_col = ""
    parts_col = ""
    commodity_col = ""

    # confirmation to proceed with script if template isn't matching good
    # values will auto populate when script ran
    continue_regardless = ""
    
    button['state']=NORMAL
    button2['state']=DISABLED
    button3['state']=DISABLED
    button4['state']=DISABLED
    button5['state']=DISABLED

    for widgets in frame.winfo_children():
      widgets.destroy()

























# # Reference of Commodity file
label = tk.Label(app_window, text= "Commodity File Last Updated: " ,anchor='s',font=('Helvetica', 12, BOLD ))
label.pack()
label.place(relx=.0625, rely=.03)
label_one = tk.Label(app_window, text= commodities_reference_last_update ,anchor='s',font=('Helvetica', 12))
label_one.pack()
label_one.place(relx= .274, rely=.03)

# Reference of Commodity file
label2 = tk.Label(app_window, text= f"Programs File Last Updated: " ,anchor='s', font=('Helvetica', 12, BOLD ))
label2.pack()
label2.place(relx=.0625, rely=.075)
label_two = tk.Label(app_window, text= programs_reference_last_update ,anchor='s',font=('Helvetica', 12))
label_two.pack()
label_two.place(relx= .26, rely=.075)

# frame where main info will be housed
frame = tk.LabelFrame(app_window, text= "Progress Information",labelanchor='n',font=('Helvetica', 12, BOLD ))
# frame.place(width=700, height= 450, relx=.0625, rely=.03)
frame.place(width=700, height= 450, relx=.0625, rely=.12)

# # TextBox Creation
# inputtxt = tk.Text(frame,
#                    height = 1,
#                    width = 50)
# inputtxt.pack()
# inputtxt.place(relx=.5, rely=.5)


# creating button to get excel file to clean up
button = tk.Button(app_window, text= "Load unclean MPL file", command= (lambda: openfile()), fg="#6F0A21", font=('Helvetica', 14, BOLD ))
button.pack()
button.place(rely=.9,relx=.06)

# creating button to process MPL cleaning
button2 = tk.Button(app_window, text= "Clean MPL file", command= (lambda: process_mpl()), font=('Helvetica', 14, BOLD ))
button2.pack()
button2.place(rely=.9,relx=.29)
button2['state']=DISABLED

# creating button to save excel file once cleaned up
button3 = tk.Button(app_window,text= "Export cleaned MPL file",command= (lambda:savefile()), font= ('Helvetica', 14 ,BOLD))
button3.pack()
# button3.place(rely=.9,relx=.455)
button3.place(rely=.875,relx=.455)
button3['state']=DISABLED

# creating button to save excel file once cleaned up
button4 = tk.Button(app_window,text= "Export MPL file changes",command= (lambda:savefile_w_references()), font= ('Helvetica', 14 ,BOLD))
button4.pack()
button4.place(rely=.925,relx=.455)
button4['state']=DISABLED

# # test button to reset
# button5 = tk.Button(text= "Process Another MPL?",command= (lambda:resetAll()), fg="#152284", font= ('Helvetica', 14 ,'bold'))
# button5.pack()
# button5.place(rely=.70,relx=.36)


# button to reset
button5 = tk.Button(text= "Process Different MPL", command = (lambda:resetAll()) ,font= ('Helvetica', 14 ,'bold'))
button5.pack()
button5.place(rely=.9,relx=.705)
button5['state']=DISABLED


    


app_window.mainloop()
