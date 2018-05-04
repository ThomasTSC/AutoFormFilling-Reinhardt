# -*- coding: utf-8 -*-
"""
Created on Mon Sep 11 13:43:03 2017

@author: chou
"""


from tkinter import messagebox

import tkinter as tkinter 


import Get_Batch_Info
import Tooling_List_Info
import Create_Folder_Files

class REINHARDT_AutoFabForm(tkinter.Frame):
    
    def __init__(self, parent):
        '''
        Constructor
        '''
        tkinter.Frame.__init__(self, parent)
        self.parent=parent
        self.REINHARDT_Interface()
    
    def REINHARDT_Interface(self):
        

        self.parent.title("REINHARDT")       
        self.parent.grid_rowconfigure(0,weight=1)
        self.parent.grid_columnconfigure(0,weight=1)
        self.parent.geometry('350x100')
        
        
        #Enter the calendar week#
        self.CW_Label = tkinter.Label(self.parent, text = ' CW # ', font = ("Arial", 12), width = 5,  height = 1).place (x = 10, y = 20)
        self.var_CW = tkinter.IntVar()
        self.CW_Entry = tkinter.Entry(self.parent, textvariable = self.var_CW, width = 10,show = None).place(x = 10, y = 50)
        
        
        #Enter the number of batch#
        self.Batch_Label = tkinter.Label(self.parent, text = ' Batch # ', font = ("Arial", 12), width = 5,  height = 1).place (x = 90, y = 20)
        self.var_Batch_Number = tkinter.IntVar()
        self.Batch_Entry = tkinter.Entry(self.parent, textvariable = self.var_Batch_Number, width = 10,show = None).place(x = 90, y = 50)


        #Choose a machine#
        self.Option_Machine = ['','VG2', 'Lesker']
        self.Option_Machine_var = tkinter.StringVar()
        self.Option_Machine_Window = tkinter.OptionMenu(self.parent, self.Option_Machine_var , *self.Option_Machine).place(x = 170, y = 45)
        


        #Create the folder, file with all information#
        
        def Create():
            
            
            #Jumping window#
            messagebox.showinfo(title = "Info", message = "Please remember the GB condition and room condition!!")
            


            self.Batch_Info = Get_Batch_Info.Get_Batch_Info(self)
            
            self.Tooling_List_Info = Tooling_List_Info.Tooling_List_Info(self)

            Create_Folder_Files.Create_Folder_Files(self) 

        
        self.Create_Button = tkinter.Button(self.parent, text = " Create ", width = 10, command = Create).place(x = 255, y = 48)
        
        
def main():
    root=tkinter.Tk()
    REINHARDT_AutoFabForm(root)
    root.mainloop()

if __name__=="__main__":
    main()
    
    
