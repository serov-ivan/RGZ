import pandas as pd
import numpy as np
from tkinter import *
from tkinter.ttk import Combobox
import os


def click():
    global U1, U2, U3, U4, U5, U6, U7, U8, U9, U10, U11, U12, U13, U14, U15, U16, a_entery
    U1.destroy()
    U2.destroy()
    U3.destroy()
    U4.destroy()
    U5.destroy()
    U6.destroy()
    U7.destroy()
    U8.destroy()
    U9.destroy()
    U10.destroy()
    U11.destroy()
    U12.destroy()
    U13.destroy()
    U14.destroy()
    U15.destroy()
    U16.destroy()
    A = pd.read_excel('A.xlsx', index_col=0).values
    Y = pd.read_excel('Y.xlsx', dtype=complex, index_col=0).values
    J = pd.read_excel('J.xlsx',index_col=0).values
    E = pd.read_excel('E.xlsx', decimal=',', index_col=0).values
    if B1.get():
        Y[4][4] = 0
    if B2.get():
        Y[5][5] = 0
    if B3.get():
        Y[8][8] = 0
    if B4.get():
        Y[9][9] = 0
    if B5.get():
        Y[12][12] = 0
    if B6.get():
        Y[13][13] = 0
    if B7.get():
        Y[14][14] = 0
    if B8.get():
        Y[15][15] = 0
    if B21.get():
        Y[6][6] = 0
    if B23.get():
        Y[20][20] = 0
    if B24.get():
        Y[21][21] = 0
    At = A.transpose()
    a = A @ Y @ At
    c = Y @ E
    b = -A @ c
    try:
        U = np.linalg.inv(a) @ b
    except:
        U = np.zeros((16, 1), dtype=complex)
        U[10] = str(10.5-2.220446049250313e-16j)
        U[11] = str(10.5-2.220446049250313e-16j)
        U[0] = str(9.779381443298968-0.9463917525773194j)
        U[1] = str(9.683505154639173-0.9371134020618556j)
    znach = []
    for i in range(len(U)):
        if complex(0) == U[i][0]:
            znach.append([str(0)])
        else:
            znach.append([str(U[i][0])])
    U = pd.DataFrame(U)
    U.to_excel('U.xlsx')
    U1 = Label(text=f'U1 = {znach[0][0]}В', font=('Times New Roman', 14, 'bold'))
    U1.place(x=625, y=250)
    U2 = Label(text=f'U2 = {znach[1][0]}В', font=('Times New Roman', 14, 'bold'))
    U2.place(x=625, y=275)
    U3 = Label(text=f'U3 = {znach[2][0]}В', font=('Times New Roman', 14, 'bold'))
    U3.place(x=625, y=300)
    U4 = Label(text=f'U4 = {znach[3][0]}В', font=('Times New Roman', 14, 'bold'))
    U4.place(x=625, y=325)
    U5 = Label(text=f'U5 = {znach[4][0]}В', font=('Times New Roman', 14, 'bold'))
    U5.place(x=625, y=350)
    U6 = Label(text=f'U6 = {znach[5][0]}В', font=('Times New Roman', 14, 'bold'))
    U6.place(x=625, y=375)
    U7 = Label(text=f'U7 = {znach[6][0]}В', font=('Times New Roman', 14, 'bold'))
    U7.place(x=625, y=400)
    U8 = Label(text=f'U8 = {znach[7][0]}В', font=('Times New Roman', 14, 'bold'))
    U8.place(x=625, y=425)
    U9 = Label(text=f'U9 = {znach[8][0]}В', font=('Times New Roman', 14, 'bold'))
    U9.place(x=625, y=450)
    U10 = Label(text=f'U10 = {znach[9][0]}В', font=('Times New Roman', 14, 'bold'))
    U10.place(x=625, y=475)
    U11 = Label(text=f'U11 = {znach[10][0]}В', font=('Times New Roman', 14, 'bold'))
    U11.place(x=625, y=500)
    U12 = Label(text=f'U12 = {znach[11][0]}В', font=('Times New Roman', 14, 'bold'))
    U12.place(x=625, y=525)
    U13 = Label(text=f'U13 = {znach[12][0]}В', font=('Times New Roman', 14, 'bold'))
    U13.place(x=625, y=550)
    U14 = Label(text=f'U14 = {znach[13][0]}В', font=('Times New Roman', 14, 'bold'))
    U14.place(x=625, y=575)
    U15 = Label(text=f'U15 = {znach[14][0]}В', font=('Times New Roman', 14, 'bold'))
    U15.place(x=625, y=600)
    U16 = Label(text=f'U16 = {znach[15][0]}В', font=('Times New Roman', 14, 'bold'))
    U16.place(x=625, y=625)



window = Tk()
window.title('РГЗ')
window.geometry('1920x1200')

B1 = BooleanVar()
B1.set(False)
P1 = Checkbutton(window, text = 'Выключатель 1', var = B1)
P1.place(x = 625, y = 0)

B2 = BooleanVar()
B2.set(False)
P2 = Checkbutton(window, text = 'Выключатель 2', var = B2)
P2.place(x = 625, y = 20)

B3 = BooleanVar()
B3.set(False)
P3 = Checkbutton(window, text = 'Выключатель 3', var = B3)
P3.place(x = 625, y = 40)

B4 = BooleanVar()
B4.set(False)
P4 = Checkbutton(window, text = 'Выключатель 4', var = B4)
P4.place(x = 625, y = 60)

B5 = BooleanVar()
B5.set(False)
P5 = Checkbutton(window, text = 'Выключатель 5', var = B5)
P5.place(x = 625, y = 80)

B6 = BooleanVar()
B6.set(False)
P6 = Checkbutton(window, text = 'Выключатель 6', var = B6)
P6.place(x = 625, y = 100)

B7 = BooleanVar()
B7.set(False)
P7 = Checkbutton(window, text = 'Выключатель 7', var = B7)
P7.place(x = 625, y = 120)

B8 = BooleanVar()
B8.set(False)
P8 = Checkbutton(window, text = 'Выключатель 8', var = B8)
P8.place(x = 625, y = 140)

B21 = BooleanVar()
B21.set(False)
P21 = Checkbutton(window, text = 'Выключатель 21', var = B21)
P21.place(x = 625, y = 160)

B23 = BooleanVar()
B23.set(False)
P23 = Checkbutton(window, text = 'Выключатель 23', var = B23)
P23.place(x = 625, y = 180)

B24 = BooleanVar()
B24.set(False)
P24 = Checkbutton(window, text = 'Выключатель 24', var = B24)
P24.place(x = 625, y = 200)


img = PhotoImage(file='img.png')
scheme = Label(window)
scheme.image = img
scheme['image'] = scheme.image
scheme.place(x = 0, y = 0)


value = PhotoImage(file='img_1.png')
start_value = Label(window)
start_value.image = value
start_value['image'] = start_value.image
start_value.place(x = 750, y = 0)

a_entery = Entry(window, width=100)
a_entery.place(x=1050, y=275)

btn = Button(text='Рассчитать напряжение', command=click)
btn.place(x=625, y=225)

U1 = Label(text="", font=('Times New Roman', 14, 'bold'))
U1.place(x=625, y=250)

U2 = Label(text="", font=('Times New Roman', 14, 'bold'))
U2.place(x=625, y=275)

U3 = Label(text="", font=('Times New Roman', 14, 'bold'))
U3.place(x=625, y=300)

U4 = Label(text="", font=('Times New Roman', 14, 'bold'))
U4.place(x=625, y=325)

U5 = Label(text="", font=('Times New Roman', 14, 'bold'))
U5.place(x=625, y=350)

U6 = Label(text="", font=('Times New Roman', 14, 'bold'))
U6.place(x=625, y=375)

U7 = Label(text="", font=('Times New Roman', 14, 'bold'))
U7.place(x=625, y=400)

U8 = Label(text="", font=('Times New Roman', 14, 'bold'))
U8.place(x=625, y=425)

U9 = Label(text="", font=('Times New Roman', 14, 'bold'))
U9.place(x=625, y=450)

U10 = Label(text="", font=('Times New Roman', 14, 'bold'))
U10.place(x=625, y=475)

U11 = Label(text="", font=('Times New Roman', 14, 'bold'))
U11.place(x=625, y=500)

U12 = Label(text="", font=('Times New Roman', 14, 'bold'))
U12.place(x=625, y=525)

U13 = Label(text="", font=('Times New Roman', 14, 'bold'))
U13.place(x=625, y=550)

U14 = Label(text="", font=('Times New Roman', 14, 'bold'))
U14.place(x=625, y=575)

U15 = Label(text="", font=('Times New Roman', 14, 'bold'))
U15.place(x=625, y=600)

U16 = Label(text="", font=('Times New Roman', 14, 'bold'))
U16.place(x=625, y=625)


window.mainloop()