# This is a sample Python script.

# Press Maj+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.

import tkinter as tk
import pandas as pd
from pandastable import Table
df = pd.DataFrame({
    'A': [1,2,3,4,5,6],
    'B': [1,2,3,4,5,6],
    'C': [1,2,3,4,5,6],
})
root = tk.Tk()
table_frame = tk.Frame(root)
table_frame.pack()
pt = Table(table_frame, dataframe=df) # it can't be `root`, it has to be `frame`
pt.show()
mask_1 = pt.model.df['A'] < 5
pt.setColorByMask('A', mask_1, 'red')
mask_2 = pt.model.df['A'] >= 5
pt.setColorByMask('A', mask_2, 'green')
root.mainloop()