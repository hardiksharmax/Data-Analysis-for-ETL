'''
Created on July 10, 2016
hardik
'''

import pandas as pd
import numpy as np
import os
import xlrd
import ntpath
import tempfile
import Tkinter,tkFileDialog
import Tkinter
from Tkinter import *
from tkFileDialog import askopenfilename,asksaveasfile,askdirectory
import tkFileDialog
class MyClass:
    def remove_existfile(self):
        if(os.path.exists("filename.name")):  # """to test whether the directory exists"""
                    os.chdir(os.getcwd())      #""""changing directory to current working directory"""
                    os.remove("filename.name") #"""removing the file"""   
    def etlextract(self):
        filename=''
        for i in range(0,72):
                    print "#",  #prints # 72 times
        print    
        print("\t\t\t\t\t\tWelcome to Import Data File")
        for i in range(0,72):
            
            print "#",#prints # 72 times
        print
        while True:
            inpt=raw_input("enter yes or y if you want to continue:::") #main while loop for continuing or exiting"""
            if(inpt=='yes' or inpt=='y'):   
                    
                    while True:
                        try:
                            mask = [("xls files","*.xls")]#used to specify file type in tkfiledialog.askopenfilename 1
                            file_save = " "
                            root = Tkinter.Tk()#"""saves an emty window in root"""
                            root.withdraw()#"""removes window from screen"""
                            filez = tkFileDialog.askopenfilenames(parent=root"""Makes window the logical parent of the dialog. The dialog is displayed on top of its parent window.""",title='Choose a file'"""gives title to window""",defaultextension=".xls",filetypes=mask"""specified in line 1""")
                            b = root.tk.splitlist(filez)#splits the attributes
                            print b
                            if( len(b)==0 or len([ x for x in b if x[len(x)-4:len(x)]!='.xls'])>0 ):
                                 print "error:please select a  input xls file"
                            else:
                                break       
                        except: 
                            print "error:please select a  input xls file" 
                    b = root.tk.splitlist(filez)
                    print '>>>>>>>>>', list(b)#prints b as a list
                    df=[]#empty list
                    for d in list(b):
                        if filez != "":#loops till end of file
                            str = d
                            file1 = ntpath.basename(str)#print file1#for indivisual file name from file of path
                            print 'file>>>>>>>>',d
                            
                            i=0
                            xls = pd.ExcelFile(d,logfile=open(os.devnull, 'wb'))#The file path of the null device
                            sheet_name=xls.sheet_names#assign sheet names to variable sheet_name
                            print '>>>>>>>>>>>>>>>>>', len(sheet_name)#prit sheet name
                            
                        for sheet in sheet_name:
                            sh=xls.parse(sheet,skiprows=12,skip_footer=5)#parse the excel file with skipping first 12 rows frm top,5 from bottom cause they are not in our intrested data
                            sh.loc[1:]#set location at sheet 1
                            str1 = sheet_name[i]#coping sheet name in srt1
                            str2 = str1
                            str3 = str2.split('f')#splitting in half cause we only need date part of name 
                            str4 = str3[1]#date part is on 1 index in name
                            #print str4
                            s1 =str4[0:5]#extract 2012 from date
                            s2 = str4[5:7]#extracting 06 i.e month from date
                            s3 = str4[7:9]#extarcting day number  frome date
                            str5 = s3+"/"+s2+"/"+s1 #concatinating 
                            sh['Week']=str5#appending new row in excel file week
                            sh['FileName']=file1#appending new row in excel file name
                            if not os.path.isfile("C:\\tmp") and not os.path.isdir("C:\\tmp"):#checking if file or directory prexists
                                os.mkdir('C:\\tmp')#making new directory 
                                print 'Building a file name yourself:'
                                filename = 'C:\\tmp\\temporaryfile.%s.csv' % os.getpid()3#creates a temp file
                                temp = open(filename, 'w+b')
                                try:
                                    print 'temp:', temp
                                    print 'temp.name:', temp.name #print file name
                                finally:
                                    temp.close()#closes the file
                    # Clean up the temporary file yourself
                                    os.remove(filename)#removes the temporary file
                                    print
                                    print 'TemporaryFile:'
                                    temp = tempfile.TemporaryFile()
                                try:
                                    print 'temp:', temp
                                    print 'temp.name:', temp.name
                                finally:
                        # Automatically cleans up the file
                                        temp.close() 
                            i=i+1
                            #f = open('filename.name', 'a') # Open file as append mode
                            sh.to_csv('filename.name',mode='a', header = False)#convering file to csv in append mode
                            
                    #print filename
                    return filename
            else:
                print "enter correct input"                #sh.to_csv('C:\\Users\\developer\\Desktop\\book.csv',mode='a',header=False)
        #filename.close()
    def etltranspose(self,filename):
        #dividing the column according to output needed
        g=pd.read_csv('filename.name',names=['Brand', 'Montreal Frd CONVEC', 'Montreal Frd SPEC','Torontod CONVEC','Torontod SPEC','Calgaryd CONVEC','Calgaryd SPEC','Vancouverd CONEC','Vancouverd SPEC','Week','FileName'])
        h=pd.melt(g,id_vars=['FileName','Week','Brand']).sort(['FileName','Week','Brand'])
        s=h.join(h["variable"].apply(lambda x: pd.Series(x.split('d'))))
        m=s[['FileName', 0,'Week', 'Brand',1, 'value']]
        m=m.rename(columns={'value':'GRP',0:"Market",1:"CONVEC_or_SPEC"})
        m=m[m.GRP>0]
        print m
        return m
    def etl_load(self,m): 
        root = Tkinter.Tk()
        root.withdraw()    
        #m.index = np.arange(0, len(m))
        print('Please select a place to save the compiled .csv file\n')
        raw_input('Press Enter to continue...')
        print('\n')
        file4=tkFileDialog.asksaveasfile(parent=root,mode='wb',title="Save as",defaultextension=".csv")
        if file4!=None:
            m.to_csv(file4.name,index=False) 
            return m
            


if __name__ == '__main__':
    ch="y" or "Y"
    while(ch=="y" or ch=="Y"):
        c=MyClass()
        c.remove_existfile()
        filename=c.etlextract()
        m=c.etltranspose(filename)
        m=c.etl_load(m)
        print "Your File has been saved in csv formate"
        print "Do You want to import more files , Prees y or Y and then Enter to be continue"
        print "Press any key and then enter to exit"
        if(ch=="y" or ch=="Y"):
            ch=raw_input()
    

    
    
    
    
    
    
