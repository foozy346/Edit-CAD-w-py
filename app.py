# this Python
#  script can be Converted into a full .NET Plugin
# #  So your C# plugin would expose AutoCAD commands, e.g.:

# [CommandMethod("RUNPY")]
# public void RunPython() {
#     PythonEngine.Initialize();
#     PythonEngine.Exec("print('Hello from Python!')");
# }

# ✅ Bonus Tip – Silent User Experience
# You can even:
# Auto-detect if AutoCAD is running, and launch it if not.
# Auto-close your script after completion.
# Show a simple GUI or no window at all (--noconsole in PyInstaller).


from pyautocad import Autocad, APoint
import re
import json
import pandas as pd
import tkinter as tk
from tkinter import Tk, filedialog, ttk
import win32com.client
import pythoncom
import time
import threading
from datetime import datetime

class Terminals:
    def __init__(self, MLHandle: str, L_S: str, FDH_Name: str, input_value: int, Terminaloutput: str, output: str, fusbm: int, units: str):
        self.Id = MLHandle
        self.le_st= L_S
        self.FDH = FDH_Name
        self.In = input_value
        self.Out = output
        self.fusbm = fusbm
        self.Tout = Terminaloutput
        self.units = units

    def __str__(self):
        return f"Terminals(id={self.Id}- input={self.FDH},{self.In}- output={self.Out}- fusbm=1x{self.fusbm}- TERMINAL OUTPUT={self.Tout}- units={self.units})"


def parse_Terminal(handle, text):
    # Remove surrounding braces if present
    text = re.sub(r'{', '', text)
    text = re.sub(r'}', '', text)


    # Remove font and color formatting like \f...; and \Cxx;
    text = re.sub(r'\\f[^;]*;', '', text)   # remove font blocks
    text = re.sub(r'\\C\d+;', '', text)     # remove color codes
    text = re.sub(r'\\.{1,4};', '', text)     # remove color codes

    # Split by \P (paragraphs/newlines)
    parts = re.split(r'\\P', text)

    data = Terminals(handle, "", "", 0, "", "", 0, "")
    identifier_set = False  # Track first non-empty line to store as Identifier

    for part in parts:
        part = part.strip()

        # Skip empty after formatting removal
        if not part:
            continue

        # AutoCAD uses weird prefixes like \P or residual POUT
        part = part.replace("\\P", "").strip()

        # If no key-value pattern but it's early content, treat as Identifier
        if not identifier_set and (":" not in part and "=" not in part and "/" in part):
            data.le_st = part
            identifier_set = True
            continue

        # Normalize POUT: to OUT: etc.
        part = re.sub(r'^P([A-Z]+):', r'\1:', part)
        # Match KEY:VALUE or KEY=VALUE
        if "IN:" in part:
            data.FDH= part.split(":")[1].split(",")[0]
            data.In= part.split(":")[1].split(",")[1]
        elif "TERMINAL OUTPUT:" in part:
            data.Tout= part.split(":")[1]
        elif "OUT:" in part:
            data.Out= part.split(":")[1]
        elif "FUSBM" in part:
            data.fusbm= part.split("=")[0].split("X")[1]
        else:
            data.units += f"{part}-"

    return data
    

def parse_mleader_text(text):
    # Remove surrounding braces if present
    text = re.sub(r'{', '', text)
    text = re.sub(r'}', '', text)


    # Remove font and color formatting like \f...; and \Cxx;
    text = re.sub(r'\\f[^;]*;', '', text)   # remove font blocks
    text = re.sub(r'\\C\d+;', '', text)     # remove color codes
    text = re.sub(r'\\.{1,4};', '', text)     # remove color codes

    # Split by \P (paragraphs/newlines)
    parts = re.split(r'\\P', text)

    data = {}
    identifier_set = False  # Track first non-empty line to store as Identifier

    for part in parts:
        part = part.strip()

        # Skip empty after formatting removal
        if not part:
            continue

        # AutoCAD uses weird prefixes like \P or residual POUT
        part = part.replace("\\P", "").strip()

        # If no key-value pattern but it's early content, treat as Identifier
        if not identifier_set and (":" not in part and "=" not in part and "/" in part):
            data["Id"] = part
            identifier_set = True
            continue

        # Normalize POUT: to OUT: etc.
        part = re.sub(r'^P([A-Z]+):', r'\1:', part)

        # Match KEY:VALUE or KEY=VALUE
        match = re.match(r'([^:=]+)[:=](.*)', part)
        if match:
            
            key = match.group(1).strip()
            value = match.group(2).strip()

            # Try convert numeric values
            if value.isdigit():
                value = int(value)

            data[key] = value

    return data


def get_mleader_texts_from_layer(acad: Autocad, layer_name="_units"):
    
    
    print(f"Connected to AutoCAD: {acad.doc.Name}")
    
    Units = []
    Terminal_list = []

    for entity in acad.iter_objects("AcDbMLeader"):  # Iterate all Multileaders
        try:
            if entity.Layer.lower() == layer_name.lower():
                # Attempt to extract MLeader text (may vary depending on structure)
                if hasattr(entity, "TextString") and entity.TextString:
                    text = entity.TextString
                else:
                    text = "<No direct text>"
                if '@' in text:
                    ter =parse_Terminal(entity.Handle, text)
                    Terminal_list.append(ter)
                else:
                    Units.append(parse_mleader_text(text))
                # print(f"Found MLeader on {layer_name}: {text}")
        except Exception as e:
            print(f"Error reading entity: {e}")

    if not Units and not Terminal_list:
        print(f"No MLeaders found on layer '{layer_name}'.")
    return Units, Terminal_list


def read_js():
    file_path = filedialog.askopenfilename(
        title="Select Excel File",
        filetypes=[("Excel Files", "*.xlsx *.xls *.xlsm")],
        initialdir= acad.doc.Path
    )

    if not file_path:
        print("No file selected.")
    else:
        print(f"Selected file: {file_path}")
        js = pd.read_excel(file_path)
        return js


def read_odn():
    file_path = filedialog.askopenfilename(
        title="Select Excel File",
        filetypes=[("Excel Files", "*.xlsx *.xls *.xlsm")],
        initialdir= acad.doc.Path
    )

    if not file_path:
        print("No file selected.")
    else:
        print(f"Selected file: {file_path}")
        odn_sheets= pd.ExcelFile(file_path).sheet_names
        odn_to_read = [i for i in odn_sheets if "Splitter" in i]
        odn = pd.read_excel(file_path, sheet_name=odn_to_read)
        return odn


def check_js(js: pd.DataFrame, ter: dict):
    js=read_js()
    ter_id_li=list(ter.keys())
    lindex=0
    for row in js.itertuples(index=False):
        lindex+=1
        #Pandas(Index=53, LEAD=1097, _2='@211', FIBER=8, PORTS=4, _5=3, _6=25, _7=28, _8=nan, ZF='ZF317,25-28', SPLTTER='2B6', SPARE='1-7', _12='HH', _13=20.0)
        # lead=0, st=1, fiberin=2, fusbm=3, hh=4, zf=8, splitter=9
        try:
            t=ter[str(int(row[0]))+'/'+row[1]]
            print(lindex, " - ", t.le_st)
            if int(t.In) != int(row[2]):
                print(f"Mismatch for Terminal {row[0]}/{row[1]}: Expected In={t.In}, Found In={row[2]}")
            if int(t.fusbm) != int(row[3]):
                print(f"Mismatch for Terminal {row[0]}/{row[1]}: Expected FUSBM={t.fusbm}, Found FUSBM={row[3]}")
            if t.Out != row[8]:
                print(type(row[8]), type(t.Out))
                print(f"Mismatch for Terminal {row[0]}/{row[1]}: Expected Out={t.Out}, Found Out={row[8]}")

        except KeyError:
            print(f"Terminal ID {row[0]}/{row[1]} not found in terminal data.", ter[ter_id_li[lindex-1]].le_st, " ?")
            x= input("press y to fix, n to exit: ")
            if x == "y":
                zoom_to_handle(ter[ter_id_li[lindex-1]].Id)
                # mleader.TextString = "Updated"
            else:   
                return
            continue
    return


def zoom_to_handle(handle):

    # Example: get object by handle
    dacad = win32com.client.Dispatch("AutoCAD.Application")
    doc = dacad.ActiveDocument
    obj = doc.HandleToObject(handle)

    # Prepare placeholders for output parameters
    ## min_point = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, [0,0,0])
    ## max_point = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, [0,0,0])
    ## print(min_point, max_point)
    min_point, max_point = obj.GetBoundingBox()

    # Zoom to window
    doc.SendCommand(f'_ZOOM W {min_point[0]},{min_point[1]} {max_point[0]},{max_point[1]} ')
    cx = (min_point[0] + max_point[0]) / 2
    cy = (min_point[1] + max_point[1]) / 2
    width = max_point[0] - min_point[0]
    height = max_point[1] - min_point[1]
    radius = max(width, height) / 2 
    circle = acad.model.AddCircle(APoint(cx, cy), radius)
    circle.Color = 1  
    circle.Layer = "NONPRINTABLE" 
    circle.LineWeight = 50 
    circle.Update()   


    # Optional: make sure selection is visible
    doc.Regen(1)


def rearrange_terminals(ter: list):
    ter.sort(key=lambda t: int(t.In), reverse=True)
    ter_map = {t.le_st: t for t in ter}
    return ter_map
        
# get data from cad file
def Connect2CAD():
    acad = Autocad()
    units, ter_li = get_mleader_texts_from_layer(acad, "_units")
    ter_map=rearrange_terminals(ter_li)
    update_status(f"Connected to AutoCAD:{acad.doc.Name} with {len(ter_li)} Terminals.")
    log("\n--- Summary of Extracted Terminals ---")

def run_check_job():
    threading.Thread(target=check_js, daemon=True).start()
    
def run_check_odn():
    threading.Thread(target=read_odn, daemon=True).start()
    
def run_connect():
    threading.Thread(target=Connect2CAD, daemon=True).start()

def update_status(text):
    status_var.set(text)
    status_bar.update_idletasks()

def log(message):
    log_text.config(state="normal")
    timestamp = datetime.now().strftime("%H:%M:%S")
    log_text.insert("end", f"[{timestamp}] {message}\n")
    log_text.see("end")
    log_text.config(state="disabled")

if __name__ == "__main__":
    try:
        root = tk.Tk()
        root.title("AutoCAD Automation Tools")
        root.geometry("600x400")
        root.configure(bg="#f9f9f9")

        # Buttons
        btn_frame = ttk.Frame(root)
        btn_frame.pack(pady=15)

        btn_job = ttk.Button(btn_frame, text="Check Job Start", command=run_check_job)
        btn_job.grid(row=1, column=0, padx=10)

        btn_odn = ttk.Button(btn_frame, text="Check ODN", command=run_check_odn)
        btn_odn.grid(row=1, column=1, padx=10)

        btn_cad = ttk.Button(btn_frame, text="Check ODN", command=run_connect)
        btn_cad.grid(row=0, column=0, padx=10, pady=10)

        # Progress Bar
        progress = ttk.Progressbar(root, orient="horizontal", mode="determinate")
        progress.pack(fill="x", padx=20, pady=(10, 5))

        # Log Output
        log_frame = ttk.LabelFrame(root, text="Log Output")
        log_frame.pack(fill="both", expand=True, padx=20, pady=(5, 10))

        log_text = tk.Text(log_frame, wrap="word", height=10, state="disabled", bg="#fcfcfc")
        log_text.pack(fill="both", expand=True, padx=5, pady=5)

        # Status Bar
        status_var = tk.StringVar(value="Ready")
        status_bar = ttk.Label(root, textvariable=status_var, anchor="w", relief="sunken")
        status_bar.pack(side="bottom", fill="x")

        # Start App
        root.mainloop()

        
    except Exception as e:
        print(f"open a cad file first or error: {e}")
