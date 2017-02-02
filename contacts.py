#!/usr/bin/python
# -*- coding: utf-8 -*-
# Contacts module. Last modification date: 02/02/2017.
# Updated from GitHub separately at each version update.

import tkinter as tk  
from tkinter import ttk 
import tkinter.messagebox as mb 
import xlwt 
from tkinter import filedialog 
import webbrowser 
from time import strftime, localtime 
from classes import CreateToolTip ################### Kostik
 
class MainTab(): 
    """Main tab of contacts""" 
    def __init__(self, root): 
 
        # Initialize tab          
        self.root=root 
        self.tabCon=tk.Frame(self.root.notebook)                                    
        self.root.notebook.add(self.tabCon, text="Контакты", image=self.root.img[21], compound="left") 
        self.tabCon.columnconfigure(0, weight=1) 
        self.tabCon.rowconfigure(1, weight=1) 
        self.tabCon.bind("<Visibility>", self.update) 
        self.entryWidth=35                                                      # width of the fields         
        self.selected=None 
         
        # Contacts frame 
        self.conFrame=tk.Frame(self.tabCon) 
        self.conFrame.grid(column=0, columnspan=1, row=0, rowspan=2, sticky="wnse") 
         
        self.statFrame=tk.Frame(self.conFrame)                                  # statistics 
        self.statFrame.pack(fill="x") 
        ttk.Button(self.statFrame, text="Экспорт непос.", image=self.root.img[13], compound="left", command=self.exportNonVisit).pack(padx=self.root.padx, side="right", 
anchor="e") 
        self.stat=tk.Label(self.statFrame) 
        self.stat.pack(side="right", anchor="e")        
         
        self.headers=["", "Участок", "Адрес", "Имя", "Не пос. до"]              # list 
        self.style = ttk.Style()
        self.style.configure("Treeview", bd=0, relief="flat")
        self.conList=ttk.Treeview(self.conFrame, padding=(0,0,20,0), columns=self.headers, selectmode="browse", show="headings", style="Treeview")
        self.conList.heading(1, image=self.root.img[3])        
        self.conList.column(0, width=1)
        self.conList.column(4, width=70)
        #self.conList.column(2, width=1)
        self.rightScrollbar = ttk.Scrollbar(self.conList, orient="vertical", command=self.conList.yview) 
        self.conList.configure(yscrollcommand=self.rightScrollbar.set) 
        self.rightScrollbar.pack(side="right", fill="y") 
        self.conList.bind("<Return>", self.openTer) 
        self.conList.bind("<Double-1>", self.openTer) 
        self.conList.bind("<<TreeviewSelect>>", self.listSelect) 
        self.conList.bind("<3>", lambda event: self.listmenu.post(event.x_root, event.y_root)) 
        self.conList.bind("<Shift-space>", self.moveCon) 
        self.conList.bind("<Delete>", self.deleteCon) 
        self.conList.pack(fill="both", expand=True, padx=self.root.padx*0, pady=self.root.pady*0)         
        self.sortCon=tk.IntVar() 
        self.sortCon.set(0)
         
        self.listbar = tk.Menu(self.conList)                                    # list context menu 
        self.listmenu = tk.Menu(self.listbar, tearoff=0) 
        self.listmenu.add_command(label="Открыть (Enter)", command=self.openTer) 
        self.listmenu.add_command(label="Перенести (Shift+Space)", command=self.moveCon) 
        self.listmenu.add_command(label="Удалить (Delete)", command=self.deleteCon) 
        self.listbar.add_cascade(label="Действия", menu=self.listmenu) 
         
        self.editFrame=tk.Frame(self.tabCon)                                    # editing 
        self.editFrame.grid(column=1, row=0, sticky="sew") 
        self.editFrame.columnconfigure(1, weight=1) 
        self.editFrame.rowconfigure(0, weight=1) 
        self.ter=tk.Label(self.editFrame) 
        self.ter.grid(column=1, row=0, padx=self.root.padx, pady=self.root.pady, sticky="w") 
        tk.Label(self.editFrame, text="Адрес").grid(column=0, row=1, padx=self.root.padx, pady=self.root.pady, sticky="w") 
        self.address=ttk.Entry(self.editFrame, width=self.entryWidth, state="disabled") 
        self.address.grid(column=1, row=1, padx=self.root.padx, pady=self.root.pady, sticky="we")
        CreateToolTip(self.address, "В адрес входит улица и номер дома")        ################ Kostik
        tk.Label(self.editFrame, text="Имя").grid(column=0, row=2, padx=self.root.padx, pady=self.root.pady, sticky="w") 
        self.name=ttk.Entry(self.editFrame, width=self.entryWidth, state="disabled")         
        self.name.grid(column=1, row=2, padx=self.root.padx, pady=self.root.pady, sticky="we") 
        tk.Label(self.editFrame, text="Не пос. до").grid(column=0, row=3, padx=self.root.padx, pady=self.root.pady, sticky="w") 
        self.nonVisit=ttk.Entry(self.editFrame, width=self.entryWidth, state="disabled")         
        self.nonVisit.grid(column=1, row=3, padx=self.root.padx, pady=self.root.pady, sticky="we")
        CreateToolTip(self.nonVisit, "Укажи тут дату до какого времени не посещать тех кто об этом попросил. (+2 года). Дату вносить в формате ГГГГ.ММ")          ################  Kostik
        self.saveButton=ttk.Button(self.editFrame, text="Сохранить изменения\nв контакте", image=self.root.img[36], compound="left", state="disabled", command=self.editCon) 
        self.saveButton.grid(column=1, columnspan=1, row=4, rowspan=2, padx=self.root.padx, pady=self.root.pady, sticky="nesw")  
         
        self.new=ttk.LabelFrame(self.tabCon, text="Новый контакт")              # new contact 
        self.new.grid(column=1,row=1, padx=self.root.padx, pady=self.root.pady, sticky="we") 
        tk.Label(self.new, text="Адрес").grid(column=0, row=0, padx=self.root.padx, pady=self.root.pady, sticky="w") 
        self.addressNew=ttk.Entry(self.new, width=self.entryWidth, state="disabled") 
        self.addressNew.grid(column=1, row=0, padx=self.root.padx, pady=self.root.pady, sticky="we") 
        CreateToolTip(self.addressNew, "В адрес входит улица и номер дома")        ################ Kostik
        tk.Label(self.new, text="Имя").grid(column=0, row=1, padx=self.root.padx, pady=self.root.pady, sticky="w") 
        self.nameNew=ttk.Entry(self.new, width=self.entryWidth, state="disabled") 
        self.nameNew.grid(column=1, row=1, padx=self.root.padx, pady=self.root.pady, sticky="we")         
        tk.Label(self.new, text="Не пос. до").grid(column=0, row=2, padx=self.root.padx, pady=self.root.pady, sticky="w") 
        self.nonVisitNew=ttk.Entry(self.new, width=self.entryWidth, state="disabled") 
        self.nonVisitNew.grid(column=1, row=2, padx=self.root.padx, pady=self.root.pady, sticky="we")
        CreateToolTip(self.nonVisitNew, "Укажи тут дату до какого времени не посещать тех кто об этом попросил. (+2 года). Дату вносить в формате ГГГГ.ММ")       ################ Kostik
        self.newButton=ttk.Button(self.new, state="disabled", image=self.root.img[21], compound="left", command=self.newSave) 
        self.newButton.grid(column=1, columnspan=1, row=3, rowspan=2, padx=self.root.padx, pady=self.root.pady, sticky="nesw") 
         
        self.chosenTer=tk.Label(self.tabCon, image=self.root.img[25], compound="left") # chosen ter on ter tab 
        self.chosenTer.grid(column=1,row=1, sticky="es") 
     
    def drawList(self): 
        self.conList.delete(*self.conList.get_children())
        self.values=tuple(self.getContent())
        for col in self.headers: self.conList.heading(col, text=col.title(), command=lambda c=col: self.sort(c))
        for item in self.values: self.conList.insert('', 'end', values=item)
        self.stat["text"]="Всего контактов: %d, не посещать: %d" % (len(self.contentFormatted), self.nonVisitNumber)       
     
    def update(self, event=None): 
        if len(self.root.list.curselection())>0: 
            self.selected=self.root.db[self.root.list.curselection()[0]]        # actual selected ter (db index)         
        if self.selected!=None: 
            self.chosenTer["text"]="Выбран участок %s" % self.selected.number 
            self.newButton["text"]="Создать контакт\nв участке %s" % self.selected.number 
            self.newButton["state"]="!disabled" 
            self.addressNew["state"]="normal" 
            self.nameNew["state"]="normal" 
            self.nonVisitNew["state"]="normal" 
        else: 
            self.chosenTer["text"]="" 
            self.newButton["text"]="Создать контакт\nв участке" 
            self.newButton["state"]="disabled" 
            self.addressNew["state"]="disabled" 
            self.nameNew["state"]="disabled" 
            self.nonVisitNew["state"]="disabled"         
        self.drawList() 
         
    def listSelect(self, event=None): 
        if len(self.conList.selection())==1: 
            self.selectedCon=self.getSelectedCon()[0] 
            self.selectedTer=self.getSelectedTer() 
            self.address["state"]="normal" 
            self.address.delete(0, "end") 
            self.address.insert(0, self.getSelectedCon()[0][0]) 
            self.name["state"]="normal" 
            self.name.delete(0, "end") 
            self.name.insert(0, self.getSelectedCon()[0][1]) 
            self.nonVisit["state"]="normal" 
            self.nonVisit.delete(0, "end") 
            self.nonVisit.insert(0, self.getSelectedCon()[0][2])            
            self.saveButton["state"]="!disabled"
            curItem = self.conList.focus()
            s=self.conList.item(curItem)["values"][0]            
            self.saveButton["text"]="Сохранить изменения\nв контакте %s" % s
        elif self.address.focus_get()=="" and self.name.focus_get()=="" and self.nonVisit.focus_get()=="": 
            self.name["state"]="disabled" 
            self.name.delete(0, "end") 
            self.address["state"]="disabled" 
            self.nonVisit["state"]="disabled" 
            self.ter["text"]="" 
            self.saveButton["text"]="Сохранить изменения\nв контакте" 
            self.saveButton["state"]="disabled"         
         
    def getSelectedCon(self): 
        """Return ter.extra object of item selected from list as [0], and all self.selected attributes as [1]""" 
        curItem = self.conList.focus()
        s=self.conList.item(curItem)["values"][0]-1
        return self.content[s][4].extra[0][self.content[s][5]], self.content[s] 
         
    def getSelectedTer(self): 
        """Return ter object of item selected from list""" 
        try:
            curItem = self.conList.focus()
            s=self.conList.item(curItem)["values"][0]-1
            return self.content[s][4] 
        except: pass
    
    def sort(self, col):
        """sort tree contents when a column header is clicked on"""
        if col=="Участок":
            self.sortCon.set(0)
            self.conList.heading(0, image="")
            self.conList.heading(1, image=self.root.img[3])
            self.conList.heading(2, image="")
            self.conList.heading(3, image="")            
            self.conList.heading(4, image="")
        elif col=="Адрес":
            self.sortCon.set(1)
            self.conList.heading(0, image="")
            self.conList.heading(1, image="")
            self.conList.heading(2, image=self.root.img[3])
            self.conList.heading(3, image="")
            self.conList.heading(4, image="")
        elif col=="Имя":
            self.sortCon.set(2)
            self.conList.heading(0, image="")
            self.conList.heading(1, image="")
            self.conList.heading(2, image="")
            self.conList.heading(3, image=self.root.img[3])
            self.conList.heading(4, image="")
        elif col=="Не пос. до":
            self.sortCon.set(3)
            self.conList.heading(0, image="")
            self.conList.heading(1, image="")
            self.conList.heading(2, image="")
            self.conList.heading(3, image="")
            self.conList.heading(4, image=self.root.img[3])            
        self.drawList()
   
    def editCon(self, event=None): 
        self.selectedCon[0]=self.address.get().strip() 
        self.selectedCon[1]=self.name.get().strip() 
        self.selectedCon[2]=self.nonVisit.get().strip() 
        self.root.log("Контакт в участке %s изменен на «%s, %s, %s»." % (self.selectedTer.number, self.selectedCon[0], self.selectedCon[1], self.selectedCon[2])) 
        self.root.save() 
        self.update() 
         
    def moveCon(self, event=None):         
        if self.selected==None: mb.showwarning("Ошибка", "Для переноса выберите один участок на вкладке участков.") 
        else:                 
            if len(self.selected.extra)==0: self.selected.extra.append([]) 
            self.selected.extra[0].append([self.getSelectedCon()[0][0], self.getSelectedCon()[0][1], self.getSelectedCon()[0][2]]) 
            self.root.log("Контакт «%s, %s, %s» перемещен из участка %s в участок %s." %\
                (self.getSelectedCon()[0][0], self.getSelectedCon()[0][1], self.getSelectedCon()[0][2], self.getSelectedTer().number, self.selected.number)) 
            self.deleteCon(move=True)                         
         
    def deleteCon(self, event=None, move=False): 
        curItem = self.conList.focus()
        try: s=self.conList.item(curItem)["values"][0]-1
        except: return        
        con=self.content[s][4].extra[0][self.content[s][5]] 
        if move==False: self.root.log("Удален контакт в участке %s (%s, %s, %s)." % (self.getSelectedTer().number, con[0], con[1], con[2])) 
        del self.content[s][4].extra[0][self.content[s][5]] 
        self.root.save() 
        self.update()
        
        curItem = self.conList.focus()
        #s=self.conList.item(curItem)["values"][0]
        self.conList.selection_set() 
         
    def newSave(self): 
        if len(self.selected.extra)==0: self.selected.extra.append([]) 
        self.selected.extra[0].append([self.addressNew.get().strip(), self.nameNew.get().strip(), self.nonVisitNew.get().strip()])        
        self.root.save() 
        self.root.log("Создан новый контакт в участке %s (%s, %s, %s)." % (self.selected.number, self.addressNew.get().strip(), self.nameNew.get().strip(), 
self.nonVisitNew.get().strip())) 
        self.update() 
         
    def getContent(self):         
        self.content=[] 
        self.contentFormatted=[] 
        self.nonVisitNumber=0 
        for ter in self.root.db: 
            if len(ter.extra)>0: 
                for e in range(len(ter.extra[0])): 
                    self.content.append([ter.number, ter.extra[0][e][0], ter.extra[0][e][1], ter.extra[0][e][2], ter, e])         
                    if ter.extra[0][e][2].strip()!="": self.nonVisitNumber+=1 
        if self.sortCon.get()==0: 
            try: self.content.sort(key=lambda x: int(x[0]))  
            except: self.content.sort(key=lambda x: x[0])           
        elif self.sortCon.get()==1: self.content.sort(key=lambda x: x[1])  
        elif self.sortCon.get()==2: self.content.sort(key=lambda x: x[2])  
        elif self.sortCon.get()==3: self.content.sort(key=lambda x: x[3], reverse=True)          
        for i in range(len(self.content)): 
            self.contentFormatted.append((i+1, "№%s–%s" % (self.content[i][0][:5], self.content[i][4].address[:25]), self.content[i][1][:25], self.content[i][2][:25], self.content[i][3])) 
        return self.contentFormatted 
 
    def openTer(self, event=None): 
        try: self.getSelectedTer().show(self.root) 
        except: pass
             
    def exportNonVisit(self):         
        wb=xlwt.Workbook() 
        ws=wb.add_sheet("Контакты не посещать") 
        row=0
        shrink=xlwt.easyxf('alignment: shrink True')         
        for i in range(len(self.content)): 
            if self.content[i][3]!="": 
                ws.write(row, 0, "№%s-%s" % (self.content[i][0], self.content[i][4].address), style=shrink) 
                ws.write(row, 1, self.content[i][1]+"\u00A0", style=shrink) 
                ws.write(row, 2, self.content[i][2]+"\u00A0", style=shrink)
                ws.write(row, 3, self.content[i][3]+"\u00A0", style=shrink)
                row+=1 
        ws.col(0).width = 4800 
        ws.col(1).width = 4800 
        ws.col(2).width = 4800
        ws.col(3).width = 1600         
        ftypes=[('Книга Excel 97-2003 (*.xls)', '.xls')]                        # save 
        filename=filedialog.asksaveasfilename(filetypes=ftypes, initialfile='Не посещать!.xls', defaultextension='.xls') 
        if filename!="": 
            try: wb.save(filename) 
            except: 
                mb.showerror("Ошибка", "Не удалось сохранить файл %s. Возможно, файл открыт или запрещен для записи." % filename) 
                print("export error") 
                self.card.root.log("Ошибка экспорта данных в файл %s." % filename) 
            else: 
                print("export successful") 
                self.root.log("Выполнен экспорт контактов в файл %s." % filename) 
                if mb.askyesno("Экспорт", "Экспорт успешно выполнен. Открыть созданный файл?")==True: webbrowser.open(filename) 
 
class TerTab(): 
    """Tab inside ter"""     
    def __init__(self, card): 
        self.card=card 
        self.tab=tk.Frame(self.card.nb) 
        self.card.nb.add(self.tab, text="Контакты", image=self.card.root.img[21], compound="left") 
        self.tab.grid_columnconfigure (1, weight=1) 
        self.tab.grid_rowconfigure (1, weight=1) 
        self.info=tk.Label(self.tab, image=None, compound="right") 
        self.info.grid(column=0, columnspan=2, row=0, sticky="e") 
        ttk.Button(self.info, text="Экспорт", image=self.card.root.img[13], compound="left", command=self.export).grid(column=1,row=0, sticky="e")       
        self.list=tk.Listbox(self.tab, relief="flat", activestyle="dotbox", font="{%s} 9" % self.card.root.listFont.get()) # list 
        self.list.grid(column=0, row=1, columnspan=2, rowspan=2, padx=self.card.root.padx*2, sticky="nesw") 
        rightScrollbar = ttk.Scrollbar(self.list, orient="vertical", command=self.list.yview) 
        self.list.configure(yscrollcommand=rightScrollbar.set) 
        rightScrollbar.pack(side="right", fill="y") 
        contacts=self.getContent() 
        self.content=tk.StringVar(value=tuple(contacts)) 
        self.list.configure(listvariable=self.content)         
        ttk.Label(self.info, text="Контактов: %d" % len(contacts)).grid(column=0,row=0, sticky="w")         
         
    def getContent(self): 
        if len(self.card.ter.extra)==0: return [] 
        else:             
            try: self.card.ter.extra.sort(key=lambda x: int(x[0]))  
            except: self.card.ter.extra.sort(key=lambda x: x[0])  
            output=[] 
            for i in range(len(self.card.ter.extra[0])): 
                if self.card.ter.extra[0][i][2]!="": nonVisit="(не пос.до %s)" % self.card.ter.extra[0][i][2] 
                else: nonVisit="" 
                output.append("%3d) %-14s %s" % (i+1, self.card.ter.extra[0][i][0][:14], self.card.ter.extra[0][i][1]+nonVisit)) 
            return output 
             
    def export(self): 
        wb=xlwt.Workbook() 
        ws=wb.add_sheet("Контакты") 
        pagesTotal=1 
        date=strftime("%d.%m", localtime()) + "." + str(int(strftime("%Y", localtime()))-2000) 
        remark =    xlwt.easyxf('alignment: shrink True;' 'font: height 200;' 'font: bold False;' 'alignment: horizontal center') 
        header1=    xlwt.easyxf('alignment: shrink True;' 'font: height 250;' 'font: bold True;'  'alignment: horizontal center;' 'borders: top medium, left medium, bottom medium, right medium') 
        header2=    xlwt.easyxf('alignment: shrink True;' 'font: height 200;' 'font: bold True;'  'alignment: horizontal center') 
        contactTop=    xlwt.easyxf('alignment: shrink True;' 'font: height 250;' 'borders: top thin') 
        contactAll=    xlwt.easyxf('alignment: shrink True;' 'font: height 250;' 'borders: top thin, left thin, bottom thin') 
        contactEmpty=    xlwt.easyxf('alignment: shrink True;' 'font: height 250') 
        #ws.write_merge(0,0, 0,1, "Не используй этот лист для записей! Перед сдачей участка вычеркни", style=remark) 
        #ws.write_merge(1,1, 0,1, "переехавших и аккуратно допиши новых на другой стороне.", style=remark)
        ws.write_merge(0,0, 0,1, "Участок №%s - %s,  (%d)" % (self.card.ter.number, self.card.ter.address, len(self.card.ter.extra[0])), style=header1) 
        ws.write_merge(21,21, 0,1, "Последний обработал: %s %s" % (self.card.ter.getPublisherFinished(), self.card.ter.getDateLastSubmit()), style=header2) 
        ws.col(0).width = 4500 
        ws.col(1).width = 6500 
         
        if len(self.card.ter.extra)!=0:                                         # writing contacts, if exist 
            row=1 
            col=0 
            if len(self.card.ter.extra[0])>20: pagesTotal=2 
            address=""             
            try: self.card.ter.extra[0].sort(key=lambda x: int(x[0]))  
            except: self.card.ter.extra[0].sort(key=lambda x: x[0])  
            for e in self.card.ter.extra[0]: 
                if address!=e[0]+"\u00A0": 
                    address=e[0]+"\u00A0" 
                    ws.write(row, col, address,style=contactTop)                     
                else: ws.write(row, col, "–",style=contactEmpty) 
                if e[2]!="": nonVisit="(не пос-ть)"
                else: nonVisit="\u00A0" 
                ws.write(row, col+1, e[1]+nonVisit ,style=contactAll) 
                row+=1 
                if row>=20: 
                    col+=2                     
                    ws.col(col).width = 4500 
                    ws.col(col+1).width = 6500 
                    row=1         
        ws.write_merge(22,22, 0,1, "(%s) Вернуть с участком! Стр. 1/%d" % (date, pagesTotal), style=remark) 
        if pagesTotal==2: 
            ws.write_merge(0,0, 2,3, "Участок №%s - %s,  (%d)" % (self.card.ter.number, self.card.ter.address, len(self.card.ter.extra[0])), style=header1) 
            ws.write_merge(22,22, 2,3, "(%s) Стр. 2/2" % date, style=remark) 
         
        ftypes=[('Книга Excel 97-2003 (*.xls)', '.xls')]                        # save 
        filename=filedialog.asksaveasfilename(filetypes=ftypes, initialfile='Контакты участка %s.xls' % self.card.ter.number, defaultextension='.xls') 
        if filename!="": 
            try: wb.save(filename) 
            except: 
                mb.showerror("Ошибка", "Не удалось сохранить файл %s. Возможно, файл открыт или запрещен для записи." % filename) 
                print("export error") 
                self.card.root.log("Ошибка экспорта данных в файл %s." % filename) 
            else: 
                print("export successful") 
                self.card.root.log("Выполнен экспорт контактов участка %s в файл %s." % (self.card.ter.number, filename)) 
                if mb.askyesno("Экспорт", "Экспорт успешно выполнен. Открыть созданный файл?")==True: webbrowser.open(filename) 
