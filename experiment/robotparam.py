"""
Chat gpt

Create a python script that will dynamically generate a tk/tcl form from a data dictionnary describing the data entries, that will then create aa JSON and send it to over UDP the IP and port mentionned in the data dictionnary. The form shouldbe organized using the "line" key/value pair to place all elements on the same line in the form. entries = {
    "IP":   {"line":1,"type": "string", "default": "192.168.1.167"},
    "port": {"line":1,"type": "int", "default": 5000},
    "save": {"line":1,"type": "string", "default": "yes"},
    "KAp":  {"line":2,"type": "float", "default": 0.1},
    "KAi":  {"line":2,"type": "float", "default": 0.1},
    "KAd":  {"line":2,"type": "float", "default": 0.1},
    "KAa":  {"line":2,"type": "float", "default": 0.1},
    "KSp":  {"line":3,"type": "float", "default": 0.1},
    "KSi":  {"line":3,"type": "float", "default": 0.1},
    "KSd":  {"line":3,"type": "float", "default": 0.1},
    "KSa":  {"line":3,"type": "float", "default": 0.1}
}
"""



import tkinter as tk
from tkinter import ttk
import json
import socket

# Data dictionary describing the data fields
entries = {
    "IP":   {"line":1,"type": "string", "default": "192.168.1.167"},
    "port": {"line":1,"type": "int", "default": 5000},
    "save": {"line":1,"type": "string", "default": "yes"},
    "KAp":  {"line":2,"type": "float", "default": 0.1},
    "KAi":  {"line":2,"type": "float", "default": 0.1},
    "KAd":  {"line":2,"type": "float", "default": 0.1},
    "KAa":  {"line":2,"type": "float", "default": 0.1},
    "KSp":  {"line":3,"type": "float", "default": 0.1},
    "KSi":  {"line":3,"type": "float", "default": 0.1},
    "KSd":  {"line":3,"type": "float", "default": 0.1},
    "KSa":  {"line":3,"type": "float", "default": 0.1},
    "calibration":  {"line":4,"type": "float", "default": 0.1}

}

# Function to send data over UDP
def send_data(data):
    ip = data["IP"]
    port = data["port"]
    message = json.dumps(data).encode('utf-8')
    
    #sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    #sock.sendto(message, (ip, port))
    #sock.close()

# Function to create the form and handle submission
def create_form(entries):
    def submit():
        data = {}
        for key, entry in form_entries.items():
            value = entry.get()
            if entries[key]["type"] == "int":
                value = int(value)
            elif entries[key]["type"] == "float":
                value = float(value)
            data[key] = value
        
        send_data(data)
        print("Data sent:", data)
    
    root = tk.Tk()
    root.title("Dynamic Form")
    
    form_entries = {}
    current_line = 0
    line_frame = None
    
    for key, entry in entries.items():
        if entry["line"] != current_line:
            current_line = entry["line"]
            line_frame = ttk.Frame(root)
            line_frame.pack(fill='x', padx=5, pady=5)
        
        label = ttk.Label(line_frame, text=key)
        label.pack(side='left', padx=5, pady=5)
        
        entry_var = tk.StringVar(value=str(entry["default"]))
        entry_widget = ttk.Entry(line_frame, textvariable=entry_var)
        entry_widget.pack(side='left', fill='x', expand=True, padx=5, pady=5)
        
        form_entries[key] = entry_var
    
    submit_button = ttk.Button(root, text="Submit", command=submit)
    submit_button.pack(pady=10)
    
    root.mainloop()

# Create the form with the given entries
create_form(entries)