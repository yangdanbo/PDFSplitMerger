#https://blog.csdn.net/kmesky/article/details/102695520
#coding=utf8

import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from tkinter import messagebox #弹窗库
from tkinter.messagebox import askyesno, askquestion
from PyPDF2 import PdfFileReader, PdfFileWriter
import os

MERGE_NONE = 0
MERGE_FILES = 1
MERGE_FOLDER = 2
SPLIT_NONE = 3
SPLIT_FILES = 4

def get_screen_size(window):
    return window.winfo_screenwidth(), window.winfo_screenheight()
 
def get_window_size(window):
    return window.winfo_reqwidth(), window.winfo_reqheight()
 
def center_window(root,  width,  height):
    screenwidth = root.winfo_screenwidth()
    screenheight = root.winfo_screenheight()
    size = '%dx%d+%d+%d' % (width, height,  (screenwidth - width)/2,  (screenheight - height)/2)
    root.geometry(size)
    # root.resizable(0,0) 

# Create instance
win = tk.Tk()   

# Add a title       
win.title("PDF Split Merger")  
win.iconbitmap("pdf.ico")

#设定600*480居中,不可调整大小
center_window(win,  706, 400)

# Exit GUI cleanly
def _quit():
    answer = askyesno(title='确认',
                          message='您确认要退出吗?')
    if answer:
        win.quit()
        win.destroy()
        #exit() 
    
def setFileAndPage(fname,  pdf,  ps):
    pdf.set(fname)
    input = PdfFileReader(open(fname, "rb"))
    # 获得源PDF文件中页面总数
    pageCount = input.getNumPages()
    ps.set("1-{}".format(pageCount))
    
def openPdfFile1():
    fname = filedialog.askopenfilename(title='打开Pdf文件', filetypes=[('Pdf file', '*.pdf'), ('All Files', '*')])
    setFileAndPage(fname, pdf1, ps1)

def openPdfFile2():
    fname = filedialog.askopenfilename(title='打开Pdf文件', filetypes=[('Pdf file', '*.pdf'), ('All Files', '*')])
    setFileAndPage(fname, pdf2, ps2)
   
def openPdfFile3():
    fname = filedialog.askopenfilename(title='打开Pdf文件', filetypes=[('Pdf file', '*.pdf'), ('All Files', '*')])
    setFileAndPage(fname, pdf3, ps3)
 
def openPdfFile4():
    fname = filedialog.askopenfilename(title='打开Pdf文件', filetypes=[('Pdf file', '*.pdf'), ('All Files', '*')])
    setFileAndPage(fname, pdf4, ps4)
   
def openPdfFile5():
    fname = filedialog.askopenfilename(title='打开Pdf文件', filetypes=[('Pdf file', '*.pdf'), ('All Files', '*')])
    setFileAndPage(fname, pdf5, ps5)

def openPdfFile6():
    fname = filedialog.askopenfilename(title='打开Pdf文件', filetypes=[('Pdf file', '*.pdf'), ('All Files', '*')])
    setFileAndPage(fname, pdf6, ps6)

def openPdfFile7():
    fname = filedialog.askopenfilename(title='打开Pdf文件', filetypes=[('Pdf file', '*.pdf'), ('All Files', '*')])
    setFileAndPage(fname, pdf7, ps7)

def openPdfFile8():
    fname = filedialog.askopenfilename(title='打开Pdf文件', filetypes=[('Pdf file', '*.pdf'), ('All Files', '*')])
    setFileAndPage(fname, pdf8, ps8)
    
def openPdfFolder():
    folderName = filedialog.askdirectory(title="选择pdf文件夹")
    print(folderName)
    foundPdf = False
    for path, dirnames, filenames in os.walk(folderName):
        for filename in filenames:
            if filename.lower().endswith(".pdf"):
                foundPdf = True
                break
            
    if foundPdf:
        folder.set(folderName)
    else:
         messagebox.showerror(folderName,'未找到pdf文件！！！')

#判断当前页码是否在选择的页码范围之内
#pageScope  1-5,10-20,15-30,3,5,-
def inPageScope(curPage, pageScope):
    if pageScope == "":
        return False
    else:
        scopes = pageScope.split(",")
        for scope in scopes:
            if scope == "-":
                return True
            elif '-' not in scope:
                return curPage+1 == scope
            else:
                limits = scope.split("-")
                if (limits[0] == "" or curPage+1 >= int( limits[0]) ) and \
                   (limits[1] == "" or curPage+1 <= int( limits[1]) ):
                    return True
                else:
                    continue
                
        return False
                
def mergeExtractedPdf(pdfFiles, outfile):
    try:
        print(outfile)
        output = PdfFileWriter()
        outputPages = 0
        for pdf_file,ps in pdfFiles:
            print(pdf_file)
            print(ps)
            print("路径：%s"%pdf_file)
            if pdf_file == "" or ps == "":
                continue

            try:
                # 读取源PDF文件
                input = PdfFileReader(open(pdf_file, "rb"), strict = False)

                # 获得源PDF文件中页面总数
                pageCount = input.getNumPages()
                print("页数：%d"%pageCount)

                # 分别将page添加到输出output中
                for iPage in range(pageCount):
                    if inPageScope(iPage, ps):
                        output.addPage(input.getPage(iPage))
                        outputPages += 1
            except:
                continue
        
        print("合并后的总页数:%d."%outputPages)
        # 写入到目标PDF文件
        outputStream = open(outfile, "wb")
        output.write(outputStream)
        outputStream.close()
        messagebox.showinfo("合并后的总页数:%d."%outputPages, "PDF文件合并完成！")
    except Exception as e:
        messagebox.showerror("合并pdf出错", str(e))
        
# 使用os模块的walk函数，搜索出指定目录下的全部PDF文件
# 获取同一目录下的所有PDF文件的绝对路径
def getPdfFiles(filedir):
    file_list = [os.path.join(root, filespath) \
                 for root, dirs, files in os.walk(filedir) \
                 for filespath in files \
                 if str(filespath).endswith('pdf')
                 ]
    return file_list if file_list else []

# 合并同一目录下的所有PDF文件
def mergePDFInFolder(filepath, outfile):
    output = PdfFileWriter()
    outputPages = 0
    pdfFiles= getPdfFiles(filepath)

    if pdfFiles:
        for pdf_file in pdfFiles:
            print("路径：%s"%pdf_file)

            # 读取源PDF文件
            input = PdfFileReader(open(pdf_file, "rb"))

            # 获得源PDF文件中页面总数
            pageCount = input.getNumPages()
            outputPages += pageCount
            print("页数：%d"%pageCount)

            # 分别将page添加到输出output中
            for iPage in range(pageCount):
                output.addPage(input.getPage(iPage))

        print("合并后的总页数:%d."%outputPages)
        # 写入到目标PDF文件
        #outputStream = open(os.path.join(filepath, outfile), "wb")
        outputStream = open(outfile, "wb")
        output.write(outputStream)
        outputStream.close()
        messagebox.showinfo("提示", "PDF文件合并完成！")
    else:
        messagebox.showinfo("提示", "没有可以合并的PDF文件！")

#检查是否选择要分割的pdf文件
def checkSplitPdfSelect():
    if pdf1.get() == "" and pdf2.get() == "" and pdf3.get() == "" :
        return SPLIT_NONE
    else:
        return SPLIT_FILES

def inSplitPageScope(curPage, pageScope):
    if pageScope == "":
        return False
    elif pageScope == "-":
        return True
    elif '-' not in pageScope:
        return curPage + 1 == int(pageScope)
    else:
        limits = pageScope.split("-")
        if (limits[0] == "" or curPage+1 >= int( limits[0]) ) and \
           (limits[1] == "" or curPage+1 <= int( limits[1]) ):
            return True
        else:
            return False
    
def splitPdf(pdf_file, ps):
     try:
        # 读取源PDF文件
        input = PdfFileReader(open(pdf_file, "rb"), strict = False)

        # 获得源PDF文件中页面总数
        pageCount = input.getNumPages()
        print("页数：%d"%pageCount)
        pdf_file_name = pdf_file.split('.')[0]
        file_idx = 0
        for ps_range in ps.split(","):
            print("PS_range：%s"%ps_range)
            output = PdfFileWriter()
            outputPages = 0
            file_idx += 1
            outfile = pdf_file_name + "_" + str(file_idx)+ ".pdf"
            # 分别将page添加到输出output中
            for iPage in range(pageCount):
                if inSplitPageScope(iPage, ps_range):
                    output.addPage(input.getPage(iPage))
                    outputPages += 1
                                
            print("分割后的总页数:%d."%outputPages)
            # 写入到目标PDF文件
            outputStream = open(outfile, "wb")
            output.write(outputStream)
            outputStream.close()
            
        messagebox.showinfo("分割后的文件数:%d."%file_idx, "PDF文件:" + pdf_file + " 分割完成！")
     except Exception as e:
        messagebox.showerror("分割pdf出错", str(e))
        
def splitPdfFiles():
    split = checkSplitPdfSelect()
    if split == SPLIT_NONE:
         messagebox.showerror('温馨提示','请选择要分割的pdf文件！！！')
    else:
        splitFiles = [(pdf1.get(),ps1.get()),(pdf2.get(),ps2.get()),(pdf3.get(),ps3.get())]
        for pdf_file,ps in splitFiles:
            if pdf_file == "" or ps == "":
                continue
            splitPdf(pdf_file,ps)
        
#检查是否选择要合并的pdf文件
def checkMergePdfSelect():
    if pdf4.get() == "" and pdf5.get() == "" and pdf6.get() == ""  and pdf7.get() == "" and pdf8.get() == "":
        return MERGE_NONE
    else:
        return MERGE_FILES
        
def mergePdfFiles():
    merge = checkMergePdfSelect()
    if merge == MERGE_NONE:
         messagebox.showerror('温馨提示','请选择要合并的pdf文件！！！')
    else:
        mergeFile = filedialog.asksaveasfilename(title="合并后的文件路径名？", filetypes=[("PDF", ".pdf")])
        if mergeFile:
            if mergeFile.lower().endswith(".pdf") == False:
                mergeFile += ".pdf"
            pdfFiles = [(pdf4.get(),ps4.get()),(pdf5.get(),ps5.get()),(pdf6.get(),ps6.get()),(pdf7.get(),ps7.get()),(pdf8.get(),ps8.get())]
            mergeExtractedPdf(pdfFiles, mergeFile)

def mergePdfFilesInFolder():
    if folder.get() == "":
        messagebox.showerror('温馨提示','请选择要合并的pdf文件所在的文件夹！！！')
    else:
        mergeFile = filedialog.asksaveasfilename(title="合并后的文件路径名？", filetypes=[("PDF", ".pdf")])
        if mergeFile:
            if mergeFile.lower().endswith(".pdf") == False:
                mergeFile += ".pdf"
            mergePDFInFolder(folder.get(), mergeFile)
         
# 选择文件分割
splitFiles = ttk.LabelFrame(win, text="分割PDF文件，页码范围：1-3,5-8,3(单页),-(全部单页)",width=640)
splitFiles.grid(column=0, row=0, padx=4, pady=4)

#可以选择三个pdf文件分割
#pdfFile1
curRow = 0
lblF1 =  ttk.Label(splitFiles,  text="文件1：")
lblF1.grid(column=0, row=curRow, sticky='W', pady=4)

pdf1 = tk.StringVar()
pdfFile1 = ttk.Entry(splitFiles,  textvariable=pdf1, width=48)
pdfFile1.grid(column=1, row=curRow, sticky='W', padx=4, pady=4)

btn1 = ttk.Button(splitFiles,  text='...', width=5, command=openPdfFile1)
btn1.grid(column=2, row=curRow, sticky='W', pady=4)

label1 =  ttk.Label(splitFiles,  text="页码范围：")
label1.grid(column=3, row=curRow, sticky='W', pady=4)

ps1 = tk.StringVar()
startPage1 = ttk.Entry(splitFiles,  textvariable=ps1,  width=18)
startPage1.grid(column=4, row=curRow, sticky='W', pady=4)

btnSplit = ttk.Button(splitFiles,  text='分割', width=5, command=splitPdfFiles)
btnSplit.grid(column=5, row=curRow, padx=4, pady=4)

#pdfFile2
curRow += 1
lblF2 =  ttk.Label(splitFiles,  text="文件2：")
lblF2.grid(column=0, row=curRow, sticky='W', pady=4)

pdf2 = tk.StringVar()
pdfFile2 = ttk.Entry(splitFiles,  textvariable=pdf2, width=48)
pdfFile2.grid(column=1, row=curRow, sticky='W', padx=4, pady=4)

btn2 = ttk.Button(splitFiles,  text='...', width=5, command=openPdfFile2)
btn2.grid(column=2, row=curRow, sticky='W', pady=4)

label2 =  ttk.Label(splitFiles,  text="页码范围：")
label2.grid(column=3, row=curRow, sticky='W', pady=4)

ps2 = tk.StringVar()
startPage2 = ttk.Entry(splitFiles,  textvariable=ps2, width=18)
startPage2.grid(column=4, row=curRow, sticky='W', pady=4)

#pdfFile3
curRow += 1
lblF3 =  ttk.Label(splitFiles,  text="文件3：")
lblF3.grid(column=0, row=curRow , sticky='W', pady=4)

pdf3 = tk.StringVar()
pdfFile3= ttk.Entry(splitFiles,  textvariable=pdf3, width=48)
pdfFile3.grid(column=1, row=curRow , sticky='W', padx=4, pady=4)

btn3 = ttk.Button(splitFiles,  text='...', width=5, command=openPdfFile3)
btn3.grid(column=2, row=curRow , sticky='W')

label3 =  ttk.Label(splitFiles,  text="页码范围：")
label3.grid(column=3, row=curRow , sticky='W', pady=4)

ps3 = tk.StringVar()
startPage3 = ttk.Entry(splitFiles,  textvariable=ps3,  width=18)
startPage3.grid(column=4, row=curRow , sticky='W', pady=4)

# 选择文件合并
mergeFiles = ttk.LabelFrame(win, text="合并PDF文件，页码范围：1-3,5-8,3(单页),-(全部页码)",width=640)
mergeFiles.grid(column=0, row=3, padx=4, pady=4)

#pdfFile4
curRow += 1
lblF4 =  ttk.Label(mergeFiles,  text="文件1：")
lblF4.grid(column=0, row=curRow , sticky='W', pady=4)

pdf4 = tk.StringVar()
pdfFile4= ttk.Entry(mergeFiles,  textvariable=pdf4, width=48)
pdfFile4.grid(column=1, row=curRow , sticky='W', padx=4, pady=4)

btn4 = ttk.Button(mergeFiles,  text='...', width=5, command=openPdfFile4)
btn4.grid(column=2, row=curRow , sticky='W')

label4 =  ttk.Label(mergeFiles,  text="页码范围：")
label4.grid(column=3, row=curRow , sticky='W', pady=4)

ps4 = tk.StringVar()
startPage4 = ttk.Entry(mergeFiles,  textvariable=ps4,  width=18)
startPage4.grid(column=4, row=curRow , sticky='W', pady=4)

btnMerge = ttk.Button(mergeFiles,  text='合并', width=5, command=mergePdfFiles)
btnMerge .grid(column=5, row=curRow, padx=4, pady=4)

#pdfFile5
curRow += 1
lblF5 =  ttk.Label(mergeFiles,  text="文件2：")
lblF5.grid(column=0, row=curRow , sticky='W', pady=4)

pdf5 = tk.StringVar()
pdfFile5= ttk.Entry(mergeFiles,  textvariable=pdf5,  width=48)
pdfFile5.grid(column=1, row=curRow , sticky='W', padx=4, pady=4)

btn5 = ttk.Button(mergeFiles,  text='...', width=5, command=openPdfFile5)
btn5.grid(column=2, row=curRow , sticky='W')

label5 =  ttk.Label(mergeFiles,  text="页码范围：")
label5.grid(column=3, row=curRow , sticky='W', pady=4)

ps5 = tk.StringVar()
startPage5 = ttk.Entry(mergeFiles,  textvariable=ps5,  width=18)
startPage5.grid(column=4,  row=curRow , sticky='W', pady=4)

#pdfFile6
curRow += 1
lblF6 =  ttk.Label(mergeFiles,  text="文件3：")
lblF6.grid(column=0, row=curRow , sticky='W', pady=4)

pdf6 = tk.StringVar()
pdfFile6= ttk.Entry(mergeFiles,  textvariable=pdf6,  width=48)
pdfFile6.grid(column=1, row=curRow , sticky='W', padx=4, pady=4)

btn6 = ttk.Button(mergeFiles,  text='...', width=5, command=openPdfFile6)
btn6.grid(column=2, row=curRow , sticky='W')

label6 =  ttk.Label(mergeFiles,  text="页码范围：")
label6.grid(column=3, row=curRow , sticky='W', pady=4)

ps6 = tk.StringVar()
startPage6 = ttk.Entry(mergeFiles,  textvariable=ps6,  width=18)
startPage6.grid(column=4, row=curRow , sticky='W', pady=4)

#pdfFile7
curRow += 1
lblF7 =  ttk.Label(mergeFiles,  text="文件4：")
lblF7.grid(column=0, row=curRow , sticky='W', pady=4)

pdf7 = tk.StringVar()
pdfFile7= ttk.Entry(mergeFiles,  textvariable=pdf7,  width=48)
pdfFile7.grid(column=1, row=curRow , sticky='W', padx=4, pady=4)

btn7 = ttk.Button(mergeFiles,  text='...', width=5, command=openPdfFile7)
btn7.grid(column=2, row=curRow , sticky='W')

label7 =  ttk.Label(mergeFiles,  text="页码范围：")
label7.grid(column=3, row=curRow , sticky='W', pady=4)

ps7 = tk.StringVar()
startPage7 = ttk.Entry(mergeFiles,  textvariable=ps7,  width=18)
startPage7.grid(column=4, row=curRow , sticky='W', pady=4)

#pdfFile8
curRow += 1
lblF8 =  ttk.Label(mergeFiles,  text="文件5：")
lblF8.grid(column=0, row=curRow , sticky='W', pady=4)

pdf8 = tk.StringVar()
pdfFile8= ttk.Entry(mergeFiles,  textvariable=pdf8,  width=48)
pdfFile8.grid(column=1, row=curRow , sticky='W', padx=4, pady=4)

btn8 = ttk.Button(mergeFiles,  text='...', width=5, command=openPdfFile8)
btn8.grid(column=2, row=curRow , sticky='W')

label8 =  ttk.Label(mergeFiles,  text="页码范围：")
label8.grid(column=3, row=curRow , sticky='W', pady=4)

ps8 = tk.StringVar()
startPage8 = ttk.Entry(mergeFiles,  textvariable=ps8,  width=18)
startPage8.grid(column=4, row=curRow , sticky='W', pady=4)

# 选择文件夹，合并其中所有pdf
curRow += 1
folderFiles = ttk.LabelFrame(win, text="合并文件夹中所有PDF文件",width=640)
folderFiles.grid(column=0, row=9, padx=4, pady=4)

lblF11 =  ttk.Label(folderFiles,  text="文件夹：")
lblF11.grid(column=0, row=curRow, sticky='W', pady=4)

folder = tk.StringVar()
pdfFolder = ttk.Entry(folderFiles,  textvariable=folder, width=66)
pdfFolder.grid(column=1, row=curRow, sticky='W', padx=4, pady=4)

btn11 = ttk.Button(folderFiles,  text='...', width=5, command=openPdfFolder)
btn11.grid(column=2, row=curRow, sticky='W')

btnFolderMerge = ttk.Button(folderFiles,  text='合并', width=5, command=mergePdfFilesInFolder)
btnFolderMerge.grid(column=4, row=curRow, padx=4, pady=4)

btnExit = ttk.Button(folderFiles,  text='退出', width=5, command=_quit)
btnExit.grid(column=5, row=curRow, padx=10, pady=4)

#======================
# Start GUI
#======================
win.mainloop()
