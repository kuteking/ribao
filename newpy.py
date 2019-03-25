import tkinter

def mywin():
    global c
    root=tkinter.Tk()
    root.title('kankan')
    root.geometry('200x300')
    str_list=['aaa','bbb','ccc']


    for n in str_list:
        c.append(tkinter.IntVar())
        b=tkinter.Checkbutton(root,text=n,variable=c[-1])
        b.pack(anchor='w')
    for nn in range(3):
        print(c[nn].get())


    tkinter.Button(root,text='anniu',command=anniu).pack()
    tkinter.Button(root,text='qiut',command=root.quit).pack()


    root.mainloop()
def anniu():
    global c
    for i in range(3):
        print(c[i].get())
if __name__=="__main__":
    c=[]
    mywin()