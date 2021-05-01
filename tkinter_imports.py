from tkinter import Tk, Label, Entry, Button, LabelFrame, IntVar, Checkbutton

janela = Tk()
janela.title('Words replacer')
janela.geometry('450x145')
janela.resizable(width=False, height=False)

frame_1 = LabelFrame()
frame_1.pack()
janela.configure(bg='lightsteelblue2')
frame_1.configure(bg='slategray1')

textinho_0= Label(frame_1, text='Word to find', bg='slategray1')
textinho_0.grid(column=1, row=0)
textinho = Label(frame_1, text='Replace', bg='slategray1')
textinho.grid(column=2, row=0)
textinho1 = Label(frame_1, text='paragraph', bg='slategray1')
textinho1.grid(column=0, row=1)
textinho1 = Label(frame_1, text='tables', bg='slategray1')
textinho1.grid(column=0, row=2)
textinho2 = Label(frame_1, text='file name: ', bg='slategray1')
textinho2.grid(column=0, row=3)

word_entry_paragraph= Entry(frame_1)
word_entry_paragraph.grid(column=1,row=1)

word_entry_tables= Entry(frame_1)
word_entry_tables.grid(column=1,row=2)

replace_paragraph = Entry(frame_1)
replace_paragraph.grid(column=2, row=1)

replace_table = Entry(frame_1)
replace_table.grid(column=2, row=2)

path_entry = Entry(frame_1)
path_entry.grid(column=1, row=3)
word= ''
var = IntVar()
checkbox = Checkbutton(frame_1, text=f'Clean this {word} word', variable=var, bg='slategray1')
checkbox.grid(column=2, row=7)

from docx_imports import feriados_retirar, paragraph_replace

botao_data = Button(frame_1, width='25', height='1', text='Do it!', bg='slategray3',
                    command=lambda: [feriados_retirar(), paragraph_replace()])
botao_data.grid(column=1, row=7)
janela.mainloop()
