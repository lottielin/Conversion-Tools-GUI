# 1. pdf -> img
# 2. pdf -> word
# 3. img --- text extraction --> word


from tkinter import *
from tkinter.filedialog import askdirectory, askopenfilename
import fitz

from pdfminer.pdfparser import PDFParser
from pdfminer.pdfdocument import PDFDocument
from pdfminer.pdfpage import PDFPage
from pdfminer.pdfpage import PDFTextExtractionNotAllowed
from pdfminer.pdfinterp import PDFResourceManager
from pdfminer.pdfinterp import PDFPageInterpreter
from pdfminer.layout import *
from pdfminer.converter import PDFPageAggregator

from paddleocr import PaddleOCR


window = Tk()
window.title('Conversion Toolset')
window.geometry('520x500')
window.resizable(width=False, height=False)


# pdf -> img
def selectPDFPath():
    in_path_pdf = askopenfilename(title='Select pdf', filetypes=[('PDF', '*.pdf')])
    path_pdf.set(in_path_pdf)

def selectImgPath():
    in_path_img = askdirectory(title='Select Img directory')
    path_img.set(in_path_img + '/')

def pdf2img(file_path, save_path):
    doc = fitz.open(file_path)
    assert doc.pageCount != 0
    for pNum in range(doc.pageCount):
        page = doc[pNum]
        rotate = int(0)
        trans = fitz.Matrix(2, 2).preRotate(rotate)
        pm = page.getPixmap(matrix=trans, alpha=False)
        pm.writePNG(save_path + '%s.png' % pNum)
    print('Successfully converted PDF to IMG')
    
def exportImg():
    pdf2img(path_pdf.get(), path_img.get())


label_2img = Label(window, text='PDF --> Img', fg='blue')
label_2img.place(x=220, y=20)
path_pdf = StringVar()
entry_pdf = Entry(window, width=30, textvariable=path_pdf)
entry_pdf.place(x=20, y=40)
button_pdf = Button(window, text='Select pdf', command=selectPDFPath, width=18, height=1)
button_pdf.place(x=300, y=38)

path_img = StringVar()
entry_img = Entry(window, width=30, textvariable=path_img)
entry_img.place(x=20, y=70)
button_img = Button(window, text='Select img saving path', command=selectImgPath, width=18, height=1)
button_img.place(x=300, y=68)
button_2img = Button(window, text='Export PDF as PNG', command=exportImg, width=15, height=1)
button_2img.place(x=150, y=100)


# pdf -> word
def selectPDFPath2():
    in_path_pdf = askopenfilename(title='Select pdf', filetypes=[('PDF', '*.pdf')])
    path_pdf2.set(in_path_pdf)

def selectWordPath():
    in_word_path = askdirectory(title='Select Word directory')
    path_word.set(in_word_path + '/')

def exportWord():
    """
    https://pdfminer-docs.readthedocs.io/programming.html
    PDFParser fetch data from file
    PDFDocument stores data
    PDFPageInterpreter process page content
    PDFResourceManager stores shared resources such as fonts or images
    """
    fp = open(path_pdf.get(), 'rb')
    parser = PDFParser(fp)   
    doc = PDFDocument(parser)

    if not doc.is_extractable:
        raise PDFTextExtractionNotAllowed
    else:
        rsrcmgr = PDFResourceManager()
        laparams = LAParams()
        device = PDFPageAggregator(rsrcmgr, laparams=laparams)
        interpreter = PDFPageInterpreter(rsrcmgr, device)
 
        for page in PDFPage.create_pages(doc):
            interpreter.process_page(page)
            layout = device.get_result()
            for element in layout:
                if isinstance(element, LTTextBoxHorizontal):
                    with open(path_wordocr.get()+'result_word.doc', 'a', encoding='utf-8') as f: 
                        results = element.get_text()
                        f.write(results)
                        f.write('\n')
    print('Successfully converted PDF to Word')


label_2word = Label(window, text='PDF --> Word', fg='blue')
label_2word.place(x=210, y=150)


path_pdf2 = StringVar()
entry_pdf2 = Entry(window, width=30, textvariable=path_pdf2)
entry_pdf2.place(x=20, y=170)
button_pdf2 = Button(window, text='Select pdf', command=selectPDFPath2, width=18, height=1)
button_pdf2.place(x=300, y=168)


path_word = StringVar()
entry_word = Entry(window, width=30, textvariable=path_word)
entry_word.place(x=20, y=200)
button_word = Button(window, text='Select Word saving directory', command=selectWordPath, width=18, height=1)
button_word.place(x=300, y=198)
button_2word = Button(window, text='Export PDF as Word', command=exportWord, width=15, height=1)
button_2word.place(x=150, y=230)



# text extraction from image
def selectImgOcr():
    in_path_imgocr = askopenfilename(title='Select image flie', filetypes=[('img', ('*.jpg','*.png')),('PNG', '*.png')])
    path_imgocr.set(in_path_imgocr)

def selectWordOcr():
    in_path_wordocr = askdirectory(title='Select directory')
    path_wordocr.set(in_path_wordocr + '/')

def ocr():
    """ ocr  = PaddleOCR(use_angle_cls=True, lang="ch", use_gpu=False,
                    rec_model_dir='./inference/ch_ppocr_server_v2.0_rec_infer/',
                    cls_model_dir='./inference/ch_ppocr_mobile_v2.0_cls_infer/',
                    det_model_dir='./inference/ch_ppocr_server_v2.0_det_infer/') """
    ocr  = PaddleOCR(use_angle_cls=True, lang="en", use_gpu=False)
    result = ocr.ocr(path_imgocr.get(), cls=True)
    for line in result:
        with open(path_wordocr.get() + "result_text.doc", 'a', encoding='utf-8') as f:
            f.write(line[1][0])
            f.write('\n')
    print('Successfully extracted text from img')

label_ocr = Label(window, text='Text extraction from image', fg='blue')
label_ocr.place(x=170, y=278)
path_imgocr = StringVar()
entry_imgoccr = Entry(window, width=30, textvariable=path_imgocr)
entry_imgoccr.place(x=20, y=300)
button_imgocr = Button(window, text='Select image to process', command=selectImgOcr, width=18, height=1)
button_imgocr.place(x=300, y=298)

path_wordocr = StringVar()
entry_wordoccr = Entry(window, width=30, textvariable=path_wordocr)
entry_wordoccr.place(x=20, y=330)
button_imgocr = Button(window, text='Select Word saving directory', command=selectWordOcr, width=18, height=1)
button_imgocr.place(x=300, y=328)

button_img2word = Button(window, text='Extract text from img', command=ocr, width=15, height=1)
button_img2word.place(x=150, y=360)

window.mainloop()