# mini-project on DIP
# convert image into 'pdf' or 'doc' file

# importing necessary libraries
import PIL
from PIL import Image
import pytesseract
import easygui  # to open the filebox
import sys
import matplotlib.pyplot as plt
import os
import tkinter as tk
from tkinter import *
from fpdf import FPDF
import docx
from docx.shared import Pt, RGBColor

top = tk.Tk()
top.geometry('600x500')
top.title('Convert image into PDF or DOC file')
top.configure(background='white')
label = Label(top, background='#CDCDCD', font=('calibri', 20, 'bold'))


class PDF(FPDF):
    pass


pdf = PDF(orientation='P', unit='mm', format='A4')
pdf.add_page()
pdf.set_font("Arial", size=15)

# Create an instance of a word document
doc = docx.Document()

def upload():
    '''
    used to find the path of the image and forward it to convert_to_pdf() function.
    :return:
    '''
    ImagePath = easygui.fileopenbox()
    # print("upload() function completed")
    read_image(ImagePath)


def read_image(path):
    print(path)
    # img = PIL.Image.open(path)
    try:
        img = PIL.Image.open(path)
    except:
        print('Error in reading file')
        msg = "Error in reading file"
        tk.messagebox.showinfo(title=None, message=msg)
    #     sys.exit()
    try:
        img = img.convert('RGBA')
        pix = img.load()
    except:
        print('Error in converting the file')
        msg = "Error in converting the file"
        tk.messagebox.showinfo(title=None, message=msg)
        sys.exit()

    for y in range(img.size[1]):
        for x in range(img.size[0]):
            if pix[x, y][0] > 102 or pix[x, y][1] > 102 or pix[x, y][2] > 102:
                pix[x, y] = (0, 0, 0, 255)
            else:
                pix[x, y] = (255, 255, 255, 255)
    fig, axes = plt.subplots(1, 1, figsize=(5, 5), subplot_kw={'xticks': [], 'yticks': []},
                             gridspec_kw=dict(hspace=0.1, wspace=0.1))
    axes.imshow(img, cmap='gray')
    # img.show()
    img.save('temp.png')
    find_text(path)
    return


def find_text(path):
    '''
    used to extract the text from the temporary png image created in the previous process
    :param path:
    :return:
    '''
    pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
    text = pytesseract.image_to_string(PIL.Image.open('temp.png'))
    msg = text
    print(msg)
    # tk.messagebox.showinfo(title=None, message=msg)
    save1 = Button(top, text="Save image in 'pdf' format", command=lambda: save_as_pdf(msg, path), padx=30, pady=5)
    save1.configure(background='#364156', foreground='white', font=('calibri', 10, 'bold'))
    save1.pack(side=TOP, pady=50)

    save2 = Button(top, text="Save image in 'doc' format", command=lambda: save_as_doc(msg, path), padx=30, pady=5)
    save2.configure(background='#364156', foreground='white', font=('calibri', 10, 'bold'))
    save2.pack(side=TOP, pady=50)


def save_as_pdf(msg, path):
    '''
    used to save the text in the pdf format.
    :param msg:
    :param path:
    :return:
    '''
    # print(path)

    newName = "Pdf_file"
    path1 = os.path.dirname(path)
    msg1 = msg.encode('latin-1', 'replace').decode('latin-1')
    # extension = os.path.splitext(path)[1]
    path = os.path.join(path1, newName + '.pdf')
    fileName = os.path.join(newName,'.pdf')
    pdf.cell(20, 5, txt=msg1, ln=2, align='P')
    try:
        # pdf.output(fileName) #, 'F'
        pdf.output("Pdf_file.pdf", 'F')
        msg = "'PDF' File created successfully"
        tk.messagebox.showinfo(title="Confirmation Message", message=msg)
        # sys.exit()
    except:
        msg = "Error in creating pdf file"
        tk.messagebox.showinfo(title="Error Message", message=msg)
        sys.exit()


def save_as_doc(msg, path):
    newName = "Doc_file"
    path1 = os.path.dirname(path)
    path = os.path.join(path1, newName + '.doc')
    # doc.add_heading(msg, 3)
    para = doc.add_paragraph().add_run(msg)
    para.font.size = Pt(12)
    para.font.color.rgb = RGBColor(0,0,0)

    try:
        doc.save(path)
        msg = "Successfully created 'doc' file"
        tk.messagebox.showinfo(title="Confirmation Message", message=msg)
        # sys.exit()
    except:
        msg = "Error in creating 'doc' file"
        tk.messagebox.showinfo(title="Error Message", message=msg)
        sys.exit()
    return


upload = Button(top, text="Upload the image", command=upload, padx=10, pady=5)
upload.configure(background='#364156', foreground='white', font=('calibri', 10, 'bold'))
upload.pack(side=TOP, pady=50)

# Button for closing
exit_button = Button(top, text="Exit", command=top.quit)
exit_button.configure(background='#364156', foreground='white', font=('calibri', 10, 'bold'))
exit_button.pack(pady=20)

# main function
if __name__ == '__main__':
    top.mainloop()
    exit(0)

