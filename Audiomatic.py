###############################################
#
#
#
#   Amazing Program written by Peter Wright
#
#   using the IBM Text to speech software
#
#                   6/3/22
#
#
#
###############################################
from os.path import join, dirname
import os
from ibm_watson import TextToSpeechV1
from ibm_watson.websocket import SynthesizeCallback
from ibm_cloud_sdk_core.authenticators import IAMAuthenticator
import pandas as pd
import tkinter as tk
from tkinter import *
from tkinter import ttk
import tkinter.font as tkFont
from tkinter.filedialog import askopenfile, askopenfilename
from tkinter.filedialog import askdirectory
from pydub import AudioSegment

#class Excel_to_Audio_Converter():
# API key and URL from 'ibm-credentials.txt'

authenticator = IAMAuthenticator('KEY REMOVED FOR SECURITY PURPOSES')
service = TextToSpeechV1(authenticator=authenticator)
service.set_service_url('URL REMOVED FOR SECURITY PURPOSES')

# designates when converting is happening
global converting_in_progress
converting_in_progress = False

# text box to hold progress
global progress

# counts the number of inputs correctly assigned
global count
count = 0

global file
global folder

#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
# OPTIONS TO BE CHOSEN BELOW
# Kevin on slow speed, medium pitch is the best in my opinion
voiceOptions = [
    'Allison', # 1
    'Emily',   # 2
    'Henry',   # 3
    'Kevin',   # 4
    'Lisa',    # 5
    'Michael', # 6
    'Olivia']  # 7

pitchOptions = [
    'x-low',    #1
    'low',      #2
    'medium',   #3
    'high',     #4
    'x-high']   #5

rateOptions = [
    'x-slow',   #1
    'slow',     #2
    'medium',   #3
    'fast',     #4
    'x-fast']   #5

global sheetOptions
sheetOptions = []

#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'''
getInt()
returns integer number for volume level
Params-> None
Returns-> num: integer volume level, can be [0, 5, 10, 15, 20]
'''
def getInt(): # gets volume level integer
    lvl = int(current_vol.get()[4])
    num = (lvl*5) - 5
    return num


'''
converter(out, message, VOICE, PITCH, RATE, SAMP_RATE)
Converts string to audio using IBM code
Params-> out: string used for path to new file
         message: string converted to audio
         VOICE: string designating the voice model to use
         PITCH: string designating the pitch
         RATE: string designating the reading speed
         SAMP_RATE: integer designating the file frequency
Returns-> puts new file in desinated location
'''
def converter(out, mess, VOICE, PITCH, RATE, SAMP_RATE):
    message = stringMaker(mess, RATE, PITCH)
    filetype = 'audio/wav;rate=' + str(SAMP_RATE)
    global converting_in_progress
    with open(join(dirname(__file__), out),'wb') as audio_file: # converts message to audio file
        if not converting_in_progress:
            return
        response = service.synthesize(message, accept=filetype, voice = VOICE).get_result()
        audio_file.write(response.content)

    wav_file = AudioSegment.from_file(file=out, format="wav") # adjust volume to desired level
    new_wav_file = wav_file + getInt()
    new_wav_file.export(out_f= out, format="wav") # overwrites file that was created
    


'''
nameMaker(int)
Converts input number to correctly formatted file name
Params-> int: input number
Returns-> filename: string file name to be added to end of path
'''
def nameMaker(int):
    filename = 'INPT' # makes usable filename for Rees system
    if(int > 99):
        filename += str(int) + '.wav'
        return filename
    elif(int > 9):
        filename += '0' + str(int) + '.wav'
        return filename
    else:
        filename += '00' + str(int) + '.wav'
        return filename


'''
stringMaker(message, rate, pitch)
Takes variables and makes formatted string for the converter
Params-> message: string to be spoken
         rate: speaking speed format string
         pitch: speaking pitch format string
Returns-> str: formatted string ready for converter
'''
def stringMaker(message, rate, pitch):
    str = '<prosody pitch="' + pitch + '" rate="' + rate + '">' + message + '</prosody>' # creates string for the desire format
    return str


'''
resetSheets(excelfile)
Changes the sheet options to reflect the currently loaded excel file
Params-> excelfile: excel file to be loaded
Returns-> adjusts sheet options on GUI to reflect excel file
'''
def resetSheets(excelfile):
    dict = pd.read_excel(excelfile, sheet_name=None) # reads in the excel file in question
    global sheetOptions
    sheetOptions = [] # initially the sheet options are null
    sheets = list(dict.keys())
    for i in range(len(dict.keys())):
        j = i + 1
        if 'Input #' in dict[sheets[i]].columns and 'Exact Text to be spoken by system' in dict[sheets[i]].columns:
            sheetOptions.append('Sheet ' + str(j) + ': ' + sheets[i]) # adds sheets that have the columns needed
    sheet['values'] = sheetOptions
    current_sheet.set('Please Select a Sheet') # prompts the user to select a sheet
    root.update()



'''
main(excelfile, dir, voice, pitch, rate, samp_rate)
Uses given information to convert excel file to wav files
Params-> excelfile: excel file in same folder to be converted
         dir: string directory of the folder to output to
         voice: string designating which voice model to use
         pitch: string designating which pitch value to use
         rate: string designating which reading rate to use
         samp_rate: number designating what sampling rate to use
Returns-> Creates new folder located in 'Output' with the created wav files
'''
def main(excelfile, dir, voice, pitch, rate, samp_rate, sheet):
    data = pd.read_excel(excelfile, sheet_name = None) # reads all sheets into dataframe dictionary with keys as sheet names
    df = data[sheet[9:]] # grabs the dataframe regarding the sheet designated by user
    global progress # global progress textbox to show text in
    if 'Input #' in df.columns:
        inputs = df['Input #'].values # grabs input numbers if they exist
        numrows = len(inputs)
        if numrows == 0:
            progress.insert(tk.END, '  Error: ', 'red')
            progress.insert(tk.END, 'Input # column empty\n') # otherwise tells the user they dont
            root.update()
            return
    else:
        progress.insert(tk.END, '  Error: ', 'red')
        progress.insert(tk.END, 'No "Input #" column\n') # or tells the user the whole column doesnt
        root.update()
        return

    if 'Exact Text to be spoken by system' in df.columns:
        text = df['Exact Text to be spoken by system'].values # grabs the text at each input
        numrows2 = len(text)
        if numrows2 == 0:
            progress.insert(tk.END, '  Error: ', 'red')
            progress.insert(tk.END, 'Spoken text column empty\n') # or says the column is empty
            root.update()
            return
    else:
        progress.insert(tk.END, '  Error: ', 'red') # or says it doesnt exist
        progress.insert(tk.END, 'No "Exact Text to be spoken by system" column\n') 
        root.update()
        return
    name = os.path.basename(excelfile.name)[:-5] 
    path = dir + "/" + name # creates folder path to make folder based on excel name
    try:
        os.mkdir(path) # makes folder 
    except:
        do_nothing = True # does nothing if folder exists
    
    progress.insert(tk.END, '\n  '+name+':\n') # starts the excel progress conversion
    root.update()
    global converting_in_progress
    for i in range(numrows):
        if not converting_in_progress:
            progress.insert(tk.END, '  Conversion Halted\n\n') # this would only happen if the convert button is clicked while running IE "Stop"
            root.update()
            return
        file = path
        file += "/" + nameMaker(inputs[i]) # creates file name and location
        mess = text[i]
        converter(file, mess, voice, pitch, rate, samp_rate) # converts message and makes file
        progress.insert(tk.END, '  Input ' + str(inputs[i]) + ' Converted\n') # tells the user
        root.update()


'''
browse_click()
responds when browse button is clicked and assigns a file and a folder
Params-> none
Returns-> assigns global file and folder, adjusts count to show if successful
'''
def browse_click():
    if converting_in_progress:
        return
    global count
    count = 0
    global file
    file = askopenfile(parent=root, mode='rb', title="Choose an Excel file to convert", 
        filetype=[("Excel file", "*.xlsx")]) # asks the user to pick excel file
    if file: # if the file is chosen correctly
        count += 1
        resetSheets(file) # changes the sheet options to represent available sheets
        global folder
        folder = askdirectory(title="Choose a directory to place folder of audio files in") # asks for folder to place files
        if folder:
            count += 1 # counts if file and folder are successfull, if they are count == 2
    

'''
start_click()
responds when convert button is clicked and runs main
'''
def start_click():
    global converting_in_progress
    if (converting_in_progress):  # if converting is happening it halts it
        converting_in_progress = False
        start_text.set("Convert")
        start_btn.configure(bg=green)
        root.update()
        return
    converting_in_progress = True
    start_text.set("Stop")
    start_btn.configure(bg='red') # otherwise it starts the converting 
    root.update()

    voiceChosen = 'en-US_' + current_voice.get() + 'V3Voice'
    pitchChosen = current_pitch.get()
    readChosen = current_read.get()
    sampChosen = current_samp.get() # grabs all of the user settings for the audio file
    sheetChosen = current_sheet.get()
    if(count == 2 and sheetChosen[0:5] == 'Sheet'): # checks to make sure the user has dont everything correctly
        main(file, folder, voiceChosen, pitchChosen, readChosen, sampChosen, sheetChosen)  # starts the conversion process
    else:
        if count != 2:
            progress.insert(tk.END, "  Error: ", 'red')
            progress.insert(tk.END, "Must browse before conversion\n") # otherwise is finds the error and tells the user
            root.update()
        else:
            progress.insert(tk.END, "  Error: ", 'red')
            progress.insert(tk.END, "Please Select Sheet\n") # otherwise is finds the error and tells the user
            root.update()

    start_text.set("Convert")
    start_btn.configure(bg=green) # when finished it sets the visuals back
    converting_in_progress = False
    root.update()

#Color used
blue = '#5260e0'
white = 'white'
green = '#3ed714'

#start of window GUI
root = tk.Tk(className = " Rees Scientific Voice Tag Software")
root.resizable(False, False)

#Fonts
title = tkFont.Font(family="Times New Roman", size = 24, weight='bold')
big = tkFont.Font(family="Times New Roman", size = 15, weight='bold')
reg = tkFont.Font(family="Times New Roman", size = 12)

#Window Creation
canvas = tk.Canvas(root, width=940, height=500)
canvas.grid(columnspan=10, rowspan=10)
canvas.configure(background=blue)

#quit button
quit = tk.Button(root, text="Quit", command=root.destroy, font=reg, width=8,
    bg=green, fg=white)
quit.place(x=840, y=450)

#title 
title_label = tk.Label(text="Excel to Audio Converter", background=blue, 
    foreground=white, font=title)
title_label.place(x=300, y=40)

#voice dropdown box
current_voice = tk.StringVar()
voices = ttk.Combobox(root, textvariable=current_voice, width=12, font=reg)
voices.place(x=75, y=175)
voices['values'] = voiceOptions
voices['state'] = 'readonly'
current_voice.set('Olivia')
voices_label = tk.Label(text="Voices", background=blue, 
    foreground=white,font=big)
voices_label.place(x=103, y=125)

#pitch dropdown box
current_pitch = tk.StringVar()
pitches = ttk.Combobox(root, textvariable=current_pitch, width=12, font=reg)
pitches.place(x=200, y=175)
pitches['values']=pitchOptions
pitches['state']='readonly'
current_pitch.set('medium')
pitches_label = tk.Label(text = "Pitches", background=blue, 
    foreground=white, font=big)
pitches_label.place(x=225, y=125)

#Reading rate dropdown box
current_read = tk.StringVar()
read = ttk.Combobox(root, textvariable=current_read, width=12, font=reg)
read.place(x=325, y=175)
read['values'] = rateOptions
read['state'] = 'readonly'
current_read.set('slow')
read_label = tk.Label(text="Speed", background=blue, 
    foreground=white, font=big)
read_label.place(x=353, y=125)

#Sheet on Excel dropdown box
current_sheet = tk.StringVar()
sheet = ttk.Combobox(root, textvariable = current_sheet, width=17, font=reg)
sheet.place(x=707, y=175)
sheet['values'] = sheetOptions
sheet['state'] = 'readonly'
current_sheet.set('Browse for File')
sheet_label = tk.Label(text="Sheet", background=blue, foreground=white, font=big)
sheet_label.place(x=755, y=125)

#volume dropdown box
current_vol = tk.StringVar()
volume = ttk.Combobox(root, textvariable=current_vol, width = 12, font=reg)
volume.place(x=450, y=175)
volume['values'] = ['Lvl 1', 'Lvl 2', 'Lvl 3', 'Lvl 4', 'Lvl 5']
sheet['state'] = 'readonly'
current_vol.set('Lvl 3')
volume_label = tk.Label(text='Volume', bg=blue, fg=white, font=big)
volume_label.place(x=475, y=125)



#sampling rate text box
current_samp = tk.StringVar()
samp = ttk.Entry(root, textvariable=current_samp, width = 15, font=reg)
samp.place(x=575, y=175)
current_samp.set('8000')
samp_label = tk.Label(text = "Sampling Rate", background=blue, 
    foreground=white, font=big)
samp_label.place(x=575, y=125)

#browsing button
browse_text = tk.StringVar()
browse_btn = tk.Button(root, textvariable=browse_text, font=reg, width=8,
     fg=white, bg=green, command=lambda:browse_click())
browse_text.set("Browse")
browse_btn.place(x=25, y= 450)

#start button
start_text = tk.StringVar()
start_btn = tk.Button(root, textvariable=start_text, font=reg, width=8,
    fg = white, bg = green, command=lambda:start_click())
start_text.set("Convert")
start_btn.place(x=125, y=450)

#instructions
words = tk.StringVar()
words.set('Instructions for use:\n1) Browse for file to be converted\n2) Select excel file\n3) Select location for folder to be created\n4) Select desired settings\n5) Press convert button\n6) Watch progress box')
instructions = tk.Label(root, textvariable = words, font=reg, width = 30, height=7, justify=LEFT)
instructions.place(x=75, y=225)

#progress textbox
progress = tk.Text(root, font=reg, width=62, height=7)
progress.place(x=364, y=225)
progress.insert("1.0", "  Press Browse to Begin\n")
progress.tag_configure('red', foreground='red')

root.mainloop()
