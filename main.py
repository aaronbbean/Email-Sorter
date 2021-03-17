import win32com.client
from tkinter import Tk, Label, Button, Listbox
import os
from tkinter import messagebox


class GUI:
    def __init__(self, master):

        self.master = master
        master.title("Email")

        self.guiSender = []
        self.guiFolder = []

        self.label = Label(master, text="What email would you like to file?")
        self.label.grid(row=0)
        self.dropDown1 = Listbox(master)
        self.dropDown1.grid(row=2)
        self.dropDown2 = Listbox(master)
        self.dropDown2.grid(row=2,column=1)
        self.list1 = 0
        self.list2 = 0
        self.getSender = Button(master, text="get sender", command=self.GetSender)
        self.getSender.grid(row=3)
        self.getFolder = Button(master, text="get folder", command=self.GetFolder)
        self.getFolder.grid(row=3,column=1)
        self.start = Button(master, text="start", command=self.GettingNameAndFile)
        self.start.grid(row=4,column=1)
        self.close_button = Button(master, text="Peace!", command=master.quit)
        self.close_button.grid(row=5)
    def makeList(self,sender):
        Email = email()
        Email.GettingStarted()
      
        g = 0

        for i in Email.SenderList:
            self.dropDown1.insert(g,i)
            g += 1




        k = 0
        for f in Email.FolderList:
            self.dropDown2.insert(k,f)

            k += 1


    def GetFolder(self):
        self.FileName = self.dropDown2.get(self.dropDown2.curselection())

    def GetSender(self):
        self.SendersName = self.dropDown1.get(self.dropDown1.curselection())

    def GettingNameAndFile(self):
        Email = email()
        x = Email.Sorter(self.FileName,self.SendersName)
        print(x)
        self.dropDown1.delete(0,x)
        self.makeList(self.SendersName)



class email:

    def __init__(self):
        self.FolderList = []
        self.SenderList = []

    def GettingStarted(self):
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

        inbox = outlook.GetDefaultFolder(6)
        for i in inbox.Folders:
            if i not in self.FolderList:
                self.FolderList.append(i)

        messages = inbox.Items

        for i in messages:
            try:
                body_content = i.sender
                if i.sender in self.SenderList:
                    pass
                else:
                    self.SenderList.append(body_content)

            except AttributeError:
                print("bad name")
        return self.SenderList
    def Sorter(self,folder,sender):

        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

        inbox = outlook.GetDefaultFolder(6)
        messages = inbox.Items
        k = 0
        for i in messages:
            try:
                checkSender = i.sender
                h = str(checkSender)
                j = str(sender)
                k += 1
                if j == h:

                    done = outlook.GetDefaultFolder(6).Folders[folder]
                    i.Move(done)
            except:
                print("well")
        return k        
    def WhoTheEmailIsFrom(self,emailContent):
        pass


    def FileEmail(self,EmailToSort):
        pass
if __name__=="__main__":
    root = Tk()


    my_gui = GUI(root)
    my_gui.makeList("none")

    root.mainloop()
