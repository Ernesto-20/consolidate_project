import tkinter

def test():
    root = tkinter.Tk('Consolidacion')
    root.title('Consolidacion')
    root.geometry('350x120')
    root.configure(bg='white')
    # frameCnt = 30
    frameCnt = 28
    frames = [tkinter.PhotoImage(file='assets/gif/loading_2.gif', format='gif -index %i' % (i)) for i in range(frameCnt)]

    def update(ind):
        frame = frames[ind]
        ind += 1
        if ind == frameCnt:
            ind = 0
        label_2.configure(image=frame)
        root.after(30, update, ind)

    label_1 = tkinter.Label(root, bg='white')
    label_1.config(font=("Calibri", 16))
    label_1.configure(text='El sistema esta procesando')
    label_1.pack()
    label_2 = tkinter.Label(root, bg='white')
    # label.place(x=0, y=50)
    label_2.pack()
    label_3 = tkinter.Label(root, bg='white')
    label_3.config(font=("Calibri", 12))
    label_3.configure(text='Por favor, espere...')
    label_3.pack()

    root.after(0, update, 0)
    # root.mainloop()

if __name__ == '__main__':
    test()