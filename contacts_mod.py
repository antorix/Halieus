#!/usr/bin/python
# -*- coding: utf-8 -*-
# Updated separately from GitHub on each version update.
import xlwt
import os
import webbrowser
from tkinter import ttk
from tkinter import filedialog 
import tkinter.messagebox as mb 
from time import strftime, localtime
from classes import CreateToolTip

noteField =     "Не пос. до"
contactStats =  "Всего контактов: %d, не посещать: %d"

def init(self):
    tip1="В адрес входит улица и номер дома"
    tip2="Укажи тут дату, до какого времени не посещать тех, кто об этом попросил (+2 года). Дату вносить в формате ГГГГ.ММ"
    CreateToolTip(self.addressNew, tip1)
    CreateToolTip(self.address, tip1)
    CreateToolTip(self.noteNew, tip2)
    CreateToolTip(self.note, tip2)    
    self.exportButton.bind("<1>", lambda x: exportNonVisit(self))
    ttk.Radiobutton(self.buttonFrame, text="Только\nне пос.", variable=self.exportType, value=1).grid(column=2,row=2, padx=self.root.padx*5, pady=self.root.pady, sticky="nw")
    ttk.Label(self.editFrame, text="Не пос.").grid(column=0, row=3, padx=self.root.padx, pady=self.root.pady, sticky="w") 
    ttk.Label(self.new, text="Не пос.").grid(column=0, row=2, padx=self.root.padx, pady=self.root.pady, sticky="w") 

def terInit(self):
    self.tabContacts.list.grid_remove()
    self.tabContacts.list.grid(column=0, row=1, columnspan=3, rowspan=2, padx=self.root.padx, sticky="nesw")
    self.tabContacts.tab.columnconfigure(0, weight=1)
    ttk.Button(self.tabContacts.tab, text="Экспорт в A6", image=self.root.img[13], compound="left", command=lambda: exportTab(self)).grid(column=1,row=0, padx=self.root.padx, pady=self.root.pady, sticky="e")            
    self.contacts=getContentMod(self)
    self.values=tuple(self.contacts)
    for col in self.tabContacts.headers: self.tabContacts.list.heading(col, text=col.title())
    for item in self.values: self.tabContacts.list.insert('', 'end', values=item)
    
def getContentMod(self):
    if len(self.ter.extra)==0: return [] 
    else:             
        self.ter.extra[0].sort(key=lambda x: x[0])  
        output=[] 
        for i in range(len(self.ter.extra[0])): 
            if self.ter.extra[0][i][2]!="": nonVisit=" (не пос. до %s)" % self.ter.extra[0][i][2].strip()
            else: nonVisit="" 
            output.append([i+1, self.ter.extra[0][i][0], self.ter.extra[0][i][1]+nonVisit]) 
        return output
    
def exportNonVisit(self, event=None):
    wb=xlwt.Workbook() 
    ws=wb.add_sheet("Контакты не посещать") 
    row=0
    shrink=xlwt.easyxf('alignment: shrink True')
    self.content.sort(key=lambda x: x[3])        
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
    ftypes=[('Книга Excel 97-2003 (*.xls)', '.xls')]
    filename=filedialog.asksaveasfilename(filetypes=ftypes, initialfile='Не посещать!.xls', defaultextension='.xls') 
    if filename!="": 
        try: wb.save(filename) 
        except: 
            mb.showerror("Ошибка", "Не удалось сохранить файл %s. Возможно, файл открыт или запрещен для записи." % filename) 
            self.card.root.log("Ошибка экспорта данных в файл %s." % filename) 
        else: 
            self.root.log("Выполнен экспорт контактов в файл %s." % filename) 
            if mb.askyesno("Экспорт", "Экспорт успешно выполнен. Открыть созданный файл?")==True: webbrowser.open(filename)
            
def exportTab(self):
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
    if len(self.ter.extra)!=0: contactsNumber=len(self.ter.extra[0])
    else: contactsNumber=0    
    ws.write_merge(0,0, 0,1, "Участок №%s - %s (%d)" % (self.ter.number, self.ter.address, contactsNumber), style=header1) 
    ws.write_merge(21,21, 0,1, "Последний обработал: %s %s" % (self.ter.getPublisherFinished(), self.ter.getDateLastSubmit()), style=header2) 
    ws.col(0).width = 4500 
    ws.col(1).width = 6500      
    row=1 
    col=0 
    if len(self.ter.extra[0])>20: pagesTotal=2 
    address=""             
    try: self.ter.extra[0].sort(key=lambda x: int(x[0]))  
    except: self.ter.extra[0].sort(key=lambda x: x[0])  
    for e in self.ter.extra[0]: 
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
        ws.write_merge(0,0, 2,3, "Участок №%s - %s,  (%d)" % (self.ter.number, self.ter.address, len(self.ter.extra[0])), style=header1) 
        ws.write_merge(22,22, 2,3, "(%s) Стр. 2/2" % date, style=remark)     
    ftypes=[('Книга Excel 97-2003 (*.xls)', '.xls')]                            # save 
    filename=filedialog.asksaveasfilename(filetypes=ftypes, initialfile='Контакты участка %s.xls' % self.ter.number, defaultextension='.xls') 
    if filename!="": 
        try: wb.save(filename) 
        except: 
            mb.showerror("Ошибка", "Не удалось сохранить файл %s. Возможно, файл открыт или запрещен для записи." % filename) 
            print("export error") 
            self.root.log("Ошибка экспорта данных в файл %s." % filename) 
        else: 
            print("export successful") 
            self.root.log("Выполнен экспорт контактов участка %s в файл %s." % (self.ter.number, filename)) 
            if mb.askyesno("Экспорт", "Экспорт успешно выполнен. Открыть созданный файл?")==True: webbrowser.open(filename) 

def convertNumber(filename):
    relpath=os.path.relpath(filename, '.xls')
    return relpath[3:relpath.index(".xls")]
