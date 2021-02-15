import tkinter as tk
import pandas as pd


from tkinter import *

##############!!!!!!!!!!!!THE FILE PATH MUST BE DECLARED BELOW!!!!!!!!!!!!!!##################
######The program will not function if the file path is not declared
pdpath = r"C:\ "


ft = ".xlsx"

UsefulData = ['Last name', 'First name', 'Country', 'Company name', 'Job title', 'Direct Phone Number', 'Email address', 'Person City']



root = tk.Tk()
tbg = "SeaGreen3"
bf = 36
ef = 75
clrf = 30
buttony = .9
buttonx = .4
clearx = .025
cleary = .9
clrw = .15
clrh =.1



canvas = tk.Canvas(root, height=400, width=550, bg="#263D42")
canvas.pack()
frame = tk.Frame(root, bg="white")
frame.place(relwidth=0.8, relheight=0.8, relx=0.1, rely=0.1)
bcki = tk.PhotoImage(file="backround.png")
bckl = tk.Label(canvas, image=bcki)
bckl.place(relwidth=1, relheight=1)
tlabel = Label(root, text="Bell Excell",font='Helvetica 18 bold', bg=tbg)
tlabel.place(relx=0, rely=0, relwidth=.3, relheight=.1)
def openex():
    #of = load_workbook('BulkFile.xlsx')
    #open(of.sheetnames)
    open(pdpath)
openexb = Button(root, text="Open Excel", command=openex, bg=tbg)
openexb.place(relx=.85, rely=0, relwidth=.15, relheight=.08)
#"turquoise"





def menu(event):
    if getcmmd.get() == "Run 1 Zoom File":
        def run_zoom():
            f1 = e1.get()
            ef1 = pd.read_excel(f1 + ft)
            v1 = ef1[UsefulData]
            dataframes = [v1]
            cf = pd.concat(dataframes)
            cf.to_excel(pdpath)

            def clear():
                b1.destroy()
                e1.destroy()
                cb.destroy()
            clear()
        def clearcmd():
            b1.destroy()
            e1.destroy()
            cb.destroy()
        cb = tk.Button(root, text='Clear', font=clrf, command=clearcmd, bg=tbg)
        cb.place(relx=clearx, rely=cleary, relwidth=clrw, relheight=clrh)
        b1 = tk.Button(root, text='Enter', font=bf, command=run_zoom, bg=tbg)
        b1.place(relx=buttonx, rely=buttony, relwidth=.2, relheight=.1)

        e1 = tk.Entry(frame, bg=tbg, font=ef)
        e1.grid(row=0, column=0)
    if getcmmd.get() == "Run 3 Zoom Files":
        def run_zoom3():
            f1 = e1.get()
            f2 = e2.get()
            f3 = e3.get()
            ef1 = pd.read_excel(f1 + ft)
            ef2 = pd.read_excel(f2 + ft)
            ef3 = pd.read_excel(f3 + ft)
            v1 = ef1[UsefulData]
            v2 = ef2[UsefulData]
            v3 = ef3[UsefulData]
            dataframes = [v1, v2, v3]

            cf = pd.concat(dataframes)

            cf.to_excel(pdpath)
            def clear3():
                b3.destroy()
                e1.destroy()
                e2.destroy()
                e3.destroy()
                cb3.destroy()
            clear3()
        def clearcmd3():
            b3.destroy()
            e1.destroy()
            e2.destroy()
            e3.destroy()
            cb3.destroy()
        cb3 = tk.Button(root, text='Clear', font=clrf, command=clearcmd3, bg=tbg)
        cb3.place(relx=clearx, rely=cleary, relwidth=clrw, relheight=clrh)
        b3 = tk.Button(root, text='Enter', font=bf, command=run_zoom3, bg=tbg)
        b3.place(relx=buttonx, rely=buttony, relwidth=.2, relheight=.1)

        e1 = tk.Entry(frame, bg=tbg, font=ef)
        e1.grid(row=0, column=0)
        e2 = tk.Entry(frame, bg=tbg, font=ef)
        e2.grid(row=2, column=0)
        e3 = tk.Entry(frame, bg=tbg, font=ef)
        e3.grid(row=4, column=0)
    if getcmmd.get() == "Run 5 Zoom Files":
        def run_zoom5():
            f1 = e1.get()
            f2 = e2.get()
            f3 = e3.get()
            f4 = e4.get()
            f5 = e5.get()

            ef1 = pd.read_excel(f1 + ft)
            ef2 = pd.read_excel(f2 + ft)
            ef3 = pd.read_excel(f3 + ft)
            ef4 = pd.read_excel(f4 + ft)
            ef5 = pd.read_excel(f5 + ft)
            v1 = ef1[UsefulData]
            v2 = ef2[UsefulData]
            v3 = ef3[UsefulData]
            v4 = ef4[UsefulData]
            v5 = ef5[UsefulData]
            dataframes = [v1, v2, v3, v4, v5]

            cf = pd.concat(dataframes)

            cf.to_excel(pdpath)
            def clear5():
                b5.destroy()
                e1.destroy()
                e2.destroy()
                e3.destroy()
                e4.destroy()
                e5.destroy()
                cb5.destroy()
            clear5()
        def clearcmd5():
            b5.destroy()
            e1.destroy()
            e2.destroy()
            e3.destroy()
            e4.destroy()
            e5.destroy()
            cb5.destroy()
        cb5 = tk.Button(root, text='Clear', font=clrf, command=clearcmd5, bg=tbg)
        cb5.place(relx=clearx, rely=cleary, relwidth=clrw, relheight=clrh)
        b5 = tk.Button(root, text='Enter', font=bf, command=run_zoom5, bg=tbg)
        b5.place(relx=buttonx, rely=buttony, relwidth=.2, relheight=.1)

        e1 = tk.Entry(frame, bg=tbg, font=ef)
        e1.grid(row=0, column=0)
        e2 = tk.Entry(frame, bg=tbg, font=ef)
        e2.grid(row=2, column=0)
        e3 = tk.Entry(frame, bg=tbg, font=ef)
        e3.grid(row=4, column=0)
        e4 = tk.Entry(frame, bg=tbg, font=ef)
        e4.grid(row=6, column=0)
        e5 = tk.Entry(frame, bg=tbg, font=ef)
        e5.grid(row=8, column=0)
    if getcmmd.get() == "Run 10 Zoom Files":
        def run_zoom10():
            f1 = e1.get()
            f2 = e2.get()
            f3 = e3.get()
            f4 = e4.get()
            f5 = e5.get()
            f6 = e6.get()
            f7 = e7.get()
            f8 = e8.get()
            f9 = e9.get()
            f10 = e10.get()
            ef1 = pd.read_excel(f1 + ft)
            ef2 = pd.read_excel(f2 + ft)
            ef3 = pd.read_excel(f3 + ft)
            ef4 = pd.read_excel(f4 + ft)
            ef5 = pd.read_excel(f5 + ft)
            ef6 = pd.read_excel(f6 + ft)
            ef7 = pd.read_excel(f7 + ft)
            ef8 = pd.read_excel(f8 + ft)
            ef9 = pd.read_excel(f9 + ft)
            ef10 = pd.read_excel(f10 + ft)
            v1 = ef1[UsefulData]
            v2 = ef2[UsefulData]
            v3 = ef3[UsefulData]
            v4 = ef4[UsefulData]
            v5 = ef5[UsefulData]
            v6 = ef6[UsefulData]
            v7 = ef7[UsefulData]
            v8 = ef8[UsefulData]
            v9 = ef9[UsefulData]
            v10 = ef10[UsefulData]
            dataframes = [v1, v2, v3, v4, v5, v6, v7, v8, v9, v10]

            cf = pd.concat(dataframes)

            cf.to_excel(pdpath)
            def clear10():
                b10.destroy()
                e1.destroy()
                e2.destroy()
                e3.destroy()
                e4.destroy()
                e5.destroy()
                e6.destroy()
                e7.destroy()
                e8.destroy()
                e9.destroy()
                e10.destroy()
                cb10.destroy()
            clear10()
        def clearcmd10():
            b10.destroy()
            e1.destroy()
            e2.destroy()
            e3.destroy()
            e4.destroy()
            e5.destroy()
            e5.destroy()
            e6.destroy()
            e7.destroy()
            e8.destroy()
            e9.destroy()
            e10.destroy()
            cb10.destroy()
        cb10 = tk.Button(root, text='Clear', font=clrf, command=clearcmd10, bg=tbg)
        cb10.place(relx=clearx, rely=cleary, relwidth=clrw, relheight=clrh)

        b10 = tk.Button(root, text='Enter', font=bf, command=run_zoom10, bg=tbg)
        b10.place(relx=buttonx, rely=buttony, relwidth=.2, relheight=.1)

        e1 = tk.Entry(frame, bg=tbg, font=ef)
        e1.grid(row=0, column=0)
        e2 = tk.Entry(frame, bg=tbg, font=ef)
        e2.grid(row=1, column=0)
        e3 = tk.Entry(frame, bg=tbg, font=ef)
        e3.grid(row=2, column=0)
        e4 = tk.Entry(frame, bg=tbg, font=ef)
        e4.grid(row=3, column=0)
        e5 = tk.Entry(frame, bg=tbg, font=ef)
        e5.grid(row=4, column=0)
        e6 = tk.Entry(frame, bg=tbg, font=ef)
        e6.grid(row=0, column=1)
        e7 = tk.Entry(frame, bg=tbg, font=ef)
        e7.grid(row=1, column=1)
        e8 = tk.Entry(frame, bg=tbg, font=ef)
        e8.grid(row=2, column=1)
        e9 = tk.Entry(frame, bg=tbg, font=ef)
        e9.grid(row=3, column=1)
        e10 = tk.Entry(frame, bg=tbg, font=ef)
        e10.grid(row=4, column=1)
    if getcmmd.get() == "Run 15 Zoom Files":
        def run_zoom15():
            f1 = e1.get()
            f2 = e2.get()
            f3 = e3.get()
            f4 = e4.get()
            f5 = e5.get()
            f6 = e6.get()
            f7 = e7.get()
            f8 = e8.get()
            f9 = e9.get()
            f10 = e10.get()
            f11 = e11.get()
            f12 = e12.get()
            f13 = e13.get()
            f14 = e14.get()
            f15 = e15.get()

            ef1 = pd.read_excel(f1 + ft)
            ef2 = pd.read_excel(f2 + ft)
            ef3 = pd.read_excel(f3 + ft)
            ef4 = pd.read_excel(f4 + ft)
            ef5 = pd.read_excel(f5 + ft)
            ef6 = pd.read_excel(f6 + ft)
            ef7 = pd.read_excel(f7 + ft)
            ef8 = pd.read_excel(f8 + ft)
            ef9 = pd.read_excel(f9 + ft)
            ef10 = pd.read_excel(f10 + ft)
            ef11 = pd.read_excel(f11 + ft)
            ef12 = pd.read_excel(f12 + ft)
            ef13 = pd.read_excel(f13 + ft)
            ef14 = pd.read_excel(f14 + ft)
            ef15 = pd.read_excel(f15 + ft)
            v1 = ef1[UsefulData]
            v2 = ef2[UsefulData]
            v3 = ef3[UsefulData]
            v4 = ef4[UsefulData]
            v5 = ef5[UsefulData]
            v6 = ef6[UsefulData]
            v7 = ef7[UsefulData]
            v8 = ef8[UsefulData]
            v9 = ef9[UsefulData]
            v10 = ef10[UsefulData]
            v11 = ef11[UsefulData]
            v12 = ef12[UsefulData]
            v13 = ef13[UsefulData]
            v14 = ef14[UsefulData]
            v15 = ef15[UsefulData]
            dataframes = [v1, v2, v3, v4, v5, v6, v7, v8, v9, v10, v11, v12, v13, v14, v15]
            cf = pd.concat(dataframes)


            cf.to_excel(pdpath)
            def clear15():
                b15.destroy()
                e1.destroy()
                e2.destroy()
                e3.destroy()
                e4.destroy()
                e5.destroy()
                e6.destroy()
                e7.destroy()
                e8.destroy()
                e9.destroy()
                e10.destroy()
                e11.destroy()
                e12.destroy()
                e13.destroy()
                e14.destroy()
                e15.destroy()
                cb15.destroy()
            clear15()
        def clearcmd15():
            b15.destroy()
            e1.destroy()
            e2.destroy()
            e3.destroy()
            e4.destroy()
            e5.destroy()
            e6.destroy()
            e7.destroy()
            e8.destroy()
            e9.destroy()
            e10.destroy()
            e11.destroy()
            e12.destroy()
            e13.destroy()
            e14.destroy()
            e15.destroy()
            cb15.destroy()
        cb15 = tk.Button(root, text='Clear', font=clrf, command=clearcmd15, bg=tbg)
        cb15.place(relx=clearx, rely=cleary, relwidth=clrw, relheight=clrh)

        b15 = tk.Button(root, text='Enter', font=bf, command=run_zoom15, bg=tbg)
        b15.place(relx=buttonx, rely=buttony, relwidth=.2, relheight=.1)

        e1 = tk.Entry(frame, bg=tbg, font=ef)
        e1.grid(row=0, column=0)
        e2 = tk.Entry(frame, bg=tbg, font=ef)
        e2.grid(row=1, column=0)
        e3 = tk.Entry(frame, bg=tbg, font=ef)
        e3.grid(row=2, column=0)
        e4 = tk.Entry(frame, bg=tbg, font=ef)
        e4.grid(row=3, column=0)
        e5 = tk.Entry(frame, bg=tbg, font=ef)
        e5.grid(row=4, column=0)
        e6 = tk.Entry(frame, bg=tbg, font=ef)
        e6.grid(row=0, column=1)
        e7 = tk.Entry(frame, bg=tbg, font=ef)
        e7.grid(row=1, column=1)
        e8 = tk.Entry(frame, bg=tbg, font=ef)
        e8.grid(row=2, column=1)
        e9 = tk.Entry(frame, bg=tbg, font=ef)
        e9.grid(row=3, column=1)
        e10 = tk.Entry(frame, bg=tbg, font=ef)
        e10.grid(row=4, column=1)
        e11 = tk.Entry(frame, bg=tbg, font=ef)
        e11.grid(row=5, column=0)
        e12 = tk.Entry(frame, bg=tbg, font=ef)
        e12.grid(row=6, column=0)
        e13 = tk.Entry(frame, bg=tbg, font=ef)
        e13.grid(row=7, column=0)
        e14 = tk.Entry(frame, bg=tbg, font=ef)
        e14.grid(row=8, column=0)
        e15 = tk.Entry(frame, bg=tbg, font=ef)
        e15.grid(row=9, column=0)



options = ["Run 1 Zoom File",
           "Run 3 Zoom Files",
           "Run 5 Zoom Files",
            "Run 10 Zoom Files",
            "Run 15 Zoom Files",
           ]

getcmmd = StringVar()
getcmmd.set(options[0])
drop = OptionMenu(root, getcmmd, *options, command=menu)
drop.config(bg=tbg)
drop.place(relx=.4, rely=0, relwidth=.3, relheight=.08)








root.mainloop()