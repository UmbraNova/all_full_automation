import tkinter as tk


window = tk.Tk()


greeting = tk.Label(text="ITEM",
                    fg="white",
                    bg="black",
                    width=50,
                    height=10).pack(fill=tk.X)

execute_btn = tk.Button(text="Teste pentru aplicatie renaming",
                        bg="red",
                        fg="white",
                        width=45,
                        height=5,
                        ).pack(fill=tk.X)


border_effects = {
    "flat": tk.FLAT,
    "sunken": tk.SUNKEN,
    "raised": tk.RAISED,
    "groove": tk.GROOVE,
    "ridge": tk.RIDGE,
}

for relief_name, relief in border_effects.items():
    frame = tk.Frame(master=window, relief=relief, borderwidth=5)
    frame.pack(side=tk.LEFT)
    label = tk.Button(master=frame, text=relief_name)
    label.pack()

user_path = tk.Entry(fg="blue", bg="black", width=60).pack()


window.mainloop()