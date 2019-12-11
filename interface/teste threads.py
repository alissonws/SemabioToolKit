import tkinter as tk
from tkinter import ttk
from threading import Thread
import math

window = tk.Tk()
window.geometry("300x150")

lbl = tk.Label(window, text="Press Start")
lbl.place(width=280, height=60, x=10, y=10)

pb = ttk.Progressbar(window)
pb.place(width=280, height=25, x=10, y=80)

def calculate():
    """ 1 / i^2 =  PI^2 / 6 """
    s = 0.0
    for i in range(1, 10000001):
        s += (1 / i**2)
        if i % 1000000 == 0:
            value = math.sqrt(s * 6)
            lbl.config(text=value) #???
            pb.step(10) #???

def start():
    lbl.config(text="Press Start")
    #calculate() #irresponsive GUI this way, obviously
    t = Thread(target=calculate)
    t.start()

btn = tk.Button(window, text="Start", command=start)
btn.place(width=280, height=25, x=10, y=115)

window.mainloop()