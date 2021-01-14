#!/usr/bin/env python3
#coding:utf8

## windows only

import sys
import win32com.client
if sys.version_info[0] < 3:
    from Tkinter import *
    import tkMessageBox as messagebox
    from tkFileDialog import asksaveasfilename 
else:
    from tkinter import *
    from tkinter import messagebox
    from tkinter.filedialog import asksaveasfilename 


class SAPI_Wrapper(object):
    def __init__(self):
        self.speaker = win32com.client.Dispatch("SAPI.SpVoice")
        self.filestream = None
        self.fname = ""
        #Sself.set_voice()
    
    def get_voice_names(self):
        return [ i.GetDescription() for i in self.speaker.GetVoices()]
        
    def set_voice(self, voice = "Chinese (Simplified)"):
        self.voices = self.speaker.GetVoices()
        # set default voice
        self.speaker.Voice = self.voices[0]
        for i in self.voices:
            desc = i.GetDescription()
            if voice in desc:
                self.speaker.Voice = i
                return True
        return False
    
    def set_save(self, fname):
        self.fname = fname
        
    def generate(self, text, speak_rate = 0):
        if len(self.fname) > 0:
            self.filestream = win32com.client.Dispatch("SAPI.SpFileStream")
            self.filestream.open(self.fname,3,False)
            backup_audio = self.speaker.AudioOutputStream
            self.speaker.AudioOutputStream = self.filestream
            self.speaker.rate = speak_rate
            self.speaker.Speak(text)
            self.filestream.close()
            self.filestream = None
            self.speaker.AudioOutputStream = backup_audio
            return self.fname
        else:
            self.speaker.Speak(text)
        return ""

g_sapi = SAPI_Wrapper()
window = Tk()
window.geometry('300x200+400+200')
window.resizable(0,0)

scroll = Scrollbar()
text_input = Text(window,width='30',height='10')
scroll.pack(side = RIGHT, fill = Y)
text_input.pack(side = TOP, fill = Y)
scroll.config(command=text_input.yview)
text_input.config(yscrollcommand = scroll.set)
text_input.pack() 

def set_save_path(): 
    global g_sapi
    files = [("wav音频文件", "*.wav")] 
    f_name = asksaveasfilename(filetypes = files, defaultextension = files) 
    g_sapi.set_save(f_name)
    
def generate(text):
    global g_sapi
    f_name = g_sapi.generate(text)
    if len(f_name) > 0:
        messagebox.showinfo(title='生成成功', message='文件已生成：'+f_name)
    else:
        messagebox.showinfo(title='播放完成', message='播放完成')

btn_save = Button(window, text = u"存储为", command = set_save_path) 
btn_save.pack()
#btn_save.pack(side = TOP, pady = 20) 

btn_gen = Button(window, text =u"生成", command = lambda: generate(text_input.get("1.0",END)), width = 10).pack()
window.mainloop()