import tkinter as tk
from classes.Taubot import TauBot


class MainGUI(tk.Frame):
    def __init__(self, parent, *args, **kwargs):
        tk.Frame.__init__(self, parent, *args, **kwargs)
        self.parent = parent
        self.parent.geometry('275x970')
        self.parent.title('TAUBOT')
        self.parent['bg'] = '#7d7d7d'
        self.tau_bot = TauBot(self)
        self.tau_bot.pack()


if __name__ == '__main__':
    root = tk.Tk()
    MainGUI(root).pack()
    root.mainloop()
