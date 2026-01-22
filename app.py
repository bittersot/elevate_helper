import tkinter 
import body
import app_icon

def button_click() -> None:
    n_copy = n_copy_entry.get()
    path = path_entry.get()

    try:
        body.main(n_copy, path)
    except:
        error_state_val.set('NOT OK!')

def print_report() -> None:
    path = path_entry.get()

    try:
        body.print_repot(path)
    except:
        error_state_val.set('NOT OK!')

icon = app_icon.icon_64()

root = tkinter.Tk()
root.title('Elevate Helper')
root.iconphoto(True, tkinter.PhotoImage(data=icon))

frame_1 = tkinter.LabelFrame(root)
frame_1.pack(padx=10, pady=10, fill='both', expand=True)

frame_1.rowconfigure(0, weight=1)
frame_1.rowconfigure(1, weight=1)
frame_1.rowconfigure(2, weight=1)
frame_1.rowconfigure(3, weight=1)
frame_1.columnconfigure(0, weight=1)
#frame_1.columnconfigure(1, weight=1)

n_copy_label = tkinter.Label(frame_1, text='Number of copies to make:', width=30, anchor='w')
n_copy_label.grid(column=0, row=0, padx=10, pady=5, sticky='news')
n_copy_entry = tkinter.Entry(frame_1, width=35, borderwidth=2)
n_copy_entry.grid(column=0, row=1, padx=10, pady=5, sticky='news')

path_label = tkinter.Label(frame_1, text='Path to the Elevate file:', width=30, anchor='w')
path_label.grid(column=0, row=2, padx=10, pady=5, sticky='news')
path_entry = tkinter.Entry(frame_1, width=35, borderwidth=2)
path_entry.grid(column=0, row=3, padx=10, pady=5, sticky='news')

frame_2 = tkinter.LabelFrame(root)
frame_2.pack(padx=10, pady=10, fill='both', expand=True)

frame_2.rowconfigure(0, weight=1)
frame_2.rowconfigure(1, weight=1)
frame_2.columnconfigure(0, weight=1)
frame_2.columnconfigure(1, weight=1)

run_button = tkinter.Button(frame_2, text='Run', command=button_click, padx=37, pady=5)
run_button.grid(column=0, row=0, padx=(10, 5), pady=5, sticky='news', columnspan=1)

exit_button = tkinter.Button(frame_2, text='Exit', command=root.destroy, padx=37, pady=5)
exit_button.grid(column=1, row=0, padx=(5, 10), pady=5, sticky='news', columnspan=1)

#frame_3 = tkinter.LabelFrame(root)
#frame_3.pack(padx=10, pady=10, fill='both', expand=True)

report_button = tkinter.Button(frame_2, text='Print report', command=print_report, padx=5, pady=5)
report_button.grid(column=0, row=1, columnspan=2, padx=10, pady=5, ipadx=70, sticky='news')

frame_4 = tkinter.LabelFrame(root)
frame_4.pack(padx=10, pady=10, fill='both', expand=True)

frame_4.rowconfigure(0, weight=1)
frame_4.columnconfigure(0, weight=1)
frame_4.columnconfigure(1, weight=1)

error_state_val = tkinter.StringVar()
error_state_val.set('OK!')

error_label = tkinter.Label(frame_4, text='Checkup:')
error_label.grid(column=0, row=0, padx=10, pady=5, sticky='w')

error_state = tkinter.Label(frame_4, textvariable=error_state_val)
error_state.grid(column=1, row=0, padx=10, pady=5, sticky='e')

root.mainloop()
