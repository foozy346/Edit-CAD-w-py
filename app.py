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
from tkinter import Tk, filedialog, ttk, simpledialog
import win32com.client
import pythoncom
import time
import threading
from datetime import datetime

class Terminals:
    def __init__(self, MLHandle: str, L_S: str, FDH_Name: str, input_value: int, Terminaloutput: str, output: str, fusbm: int, units: str, content: str):
        self.Id = MLHandle
        self.le_st= L_S
        self.FDH = FDH_Name
        self.In = input_value
        self.Out = output
        self.fusbm = fusbm
        self.Tout = Terminaloutput
        self.units = units
        self.contents = content

    def __str__(self):
        return f"Terminals(id={self.Id}- input={self.FDH},{self.In}- output={self.Out}- fusbm=1x{self.fusbm}- TERMINAL OUTPUT={self.Tout}- units={self.units})"


def parse_Terminal(handle, text):
    # Remove surrounding braces if present
    og_con=text
    text = re.sub(r'{', '', text)
    text = re.sub(r'}', '', text)


    # Remove font and color formatting like \f...; and \Cxx;
    text = re.sub(r'\\f[^;]*;', '', text)   # remove font blocks
    text = re.sub(r'\\C\d+;', '', text)     # remove color codes
    text = re.sub(r'\\.{1,4};', '', text)     # remove color codes

    # Split by \P (paragraphs/newlines)
    parts = re.split(r'\\P', text)

    data = Terminals(handle, "", "", 0, "", "", 0, "", og_con)
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
    
    log("--- EXTRACTING TERMINALS ---")
    update_status(f"Connected to AutoCAD: {acad.doc.Name}")

    # print(f"Connected to AutoCAD: {acad.doc.Name}")
    
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
                    # log(f"{str(ter.contents)}")
                    log(f"{str(ter.le_st)} - IN={str(ter.FDH)},{str(ter.In)} - ZF={str(ter.Out)}")
                else:
                    Units.append(parse_mleader_text(text))
                # print(f"Found MLeader on {layer_name}: {text}")
        except Exception as e:
            log(f"Error reading entity: {e}")

    if not Units and not Terminal_list:
        log(f"No MLeaders found on layer '{layer_name}'.")
    log("--- DONE ---")
    return Units, Terminal_list


def read_js():
    try:
        file_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel Files", "*.xlsx *.xls *.xlsm")],
            initialdir= acad.doc.Path
        )
        if not file_path:
            log("No file selected.")
        else:
            update_status(f"Selected file: {file_path}")
            js = pd.read_excel(file_path)
            return js
    except Exception as e:
        log(f"Error opening file dialog: {e}")

def read_odn():
    """  ODN ROW INDEX
    0         1                 2            3         4                         5                            6                      7                                 
    Sequence  Span Feet         Route/Lead   Terminal Terminal Type             Drop\nPorts                 Branch?\n(Y/N)          Terminal\nOutput                  ONT Level          RUS Unit     Vendor Part #                         Terminal Count\n(ZF)    Cable Count   Input Pwr   Thru Calc    Out Calc                 0                0         NaN
    2         190               1097         @125          Pole                       4                            NaN                -17.282852                   -17.6383           USBM1X4 PLO-U-104-SCA-SCA                                    ZF314,1-4       WNDFL,72  -10.042852       0.035  -10.077852          0.032126             0.07         NaN
    """
     
    try:
        file_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel Files", "*.xlsx *.xls *.xlsm")],
            initialdir= acad.doc.Path
        )
        if not file_path:
            log("No file selected.")
        else:
            odn_d={}
            update_status(f"Selected file: {file_path}")
            odn_sheets= pd.ExcelFile(file_path).sheet_names
            odn_to_read = [i for i in odn_sheets if "Splitter" in i]
            odn = pd.read_excel(file_path, sheet_name=odn_to_read)
            for _, df in odn.items():
                    for row in df.itertuples(index=False):
                        if row[5] in [4, 8]: 
                            odn_d[f"{row[2]}/{row[3]}"] = [row[7], row[5]]
                            print(f"{row[2]}/{row[3]} - TOUT= {row[7]:.2f} - fusbm= 1x{row[5]}") 
            
            return odn_d

    except Exception as e:
        log(f"Error opening file dialog: {e}")

    

def input_ODN():
    odn=read_odn()
    global acad, glob_units, glob_terminals
    ter_id_li=list(glob_terminals.keys())
    for l_s in ter_id_li:
        try:
            t=glob_terminals[l_s]
            t.Tout=f"{odn[l_s][0]:.2f}"
            print(t.Tout)
            if int(t.fusbm) != int(odn[l_s][1]):
                log(f"Mismatch for Terminal {l_s}: Expected FUSBM=1x{t.fusbm}, ODN FUSBM=1x{odn[l_s][1]}")
            else:
                log(f"{str(t.le_st)} - IN={t.FDH},{t.In} - Tout={t.Tout} - fusbm=1x{t.fusbm}")
            t.contents = re.sub("TERMINAL OUTPUT:([A-Za-z0-9\.\-\+]+)dBm", f"TERMINAL OUTPUT:{t.Tout}dBm", t.contents)
            zoom_replace(glob_terminals[l_s].Id, t.contents) 
        except KeyError:
            log(f"Terminal ID {l_s} not found in ODN EXCEL.")
            continue
    log("---- Done ODN Updated ----")
    return

def check_js():
    global acad, glob_units, glob_terminals
    js=read_js()
    ter_id_li=list(glob_terminals.keys())
    lindex=0
    for row in js.itertuples(index=False):
        lindex+=1
        #Pandas(Index=53, LEAD=1097, _2='@211', FIBER=8, PORTS=4, _5=3, _6=25, _7=28, _8=nan, ZF='ZF317,25-28', SPLTTER='2B6', SPARE='1-7', _12='HH', _13=20.0)
        # lead=0, st=1, fiberin=2, fusbm=3, hh=4, zf=8, splitter=9
        try:
            t=glob_terminals[str(int(row[0]))+'/'+row[1]]
            log(f"{str(t.le_st)} - IN={t.FDH},{t.In} - ZF={t.Out}")
            if int(t.In) != int(row[2]):
                log(f"Mismatch for Terminal {row[0]}/{row[1]}: Expected In={t.In}, Found In={row[2]}")
            if int(t.fusbm) != int(row[3]):
                log(f"Mismatch for Terminal {row[0]}/{row[1]}: Expected FUSBM={t.fusbm}, Found FUSBM={row[3]}")
            if t.Out != row[8]:
                log(type(row[8]), type(t.Out))
                log(f"Mismatch for Terminal {row[0]}/{row[1]}: Expected Out={t.Out}, Found Out={row[8]}")

        except KeyError:
            log(f"Terminal ID {row[0]}/{row[1]} not found in terminal data. {glob_terminals[ter_id_li[lindex-1]].le_st} ?")
            zoom_to_handle(glob_terminals[ter_id_li[lindex-1]].Id)
            # mleader.TextString = "Updated"
            continue
    log("---- Done ----")
    return

def input_zf():
    global acad, glob_units, glob_terminals
    js=read_js()
    fdh=get_user_input("Enter FDH Name to process:", "FDH Name")
    ter_id_li=list(glob_terminals.keys())
    lindex=0
    for row in js.itertuples(index=False):
        lindex+=1
        #Pandas(Index=53, LEAD=1097, _2='@211', FIBER=8, PORTS=4, _5=3, _6=25, _7=28, _8=nan, ZF='ZF317,25-28', SPLTTER='2B6', SPARE='1-7', _12='HH', _13=20.0)
        # lead=0, st=1, fiberin=2, fusbm=3, hh=4, zf=8, splitter=9
        try:
            t=glob_terminals[str(int(row[0]))+'/'+row[1]]
            t.contents = re.sub("N:([A-Z0-9]+,\d+|[A-Z]+,[A-Z0-9]+)", f"N:{fdh},{row[2]}", t.contents)
            t.contents = re.sub("OUT:([A-Z0-9]+,[A-Z0-9\-]+)", f"OUT:{row[8]}".upper(), t.contents)
            zoom_replace(glob_terminals[ter_id_li[lindex-1]].Id, t.contents)  
            log(f"{str(t.le_st)} - IN={fdh},{row[2]} - ZF={row[8]}")
            

        except KeyError:
            log(f"Terminal ID {row[0]}/{row[1]} not found in terminal data. {glob_terminals[ter_id_li[lindex-1]].le_st} ?")
            zoom_to_handle(glob_terminals[ter_id_li[lindex-1]].Id)
            continue
    log("---- Done ZF Updated ----")
    return


def zoom_replace(handle, new_text):
    try:   
        dacad = win32com.client.Dispatch("AutoCAD.Application")
        doc = dacad.ActiveDocument
        obj = doc.HandleToObject(handle)
        obj.TextString = new_text
        min_point, max_point = obj.GetBoundingBox()
        doc.SendCommand(f'_ZOOM W {min_point[0]},{min_point[1]} {max_point[0]},{max_point[1]} ')
        doc.Regen(1)
    except Exception:
        log(f"Error updating MLeader: {obj.TextString}")


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
    try:
        global acad, glob_units, glob_terminals
        acad = Autocad()
        glob_units, ter_li = get_mleader_texts_from_layer(acad, "_units")
        glob_terminals=rearrange_terminals(ter_li)
        update_status(f"Connected to AutoCAD: {acad.doc.Name} with {len(ter_li)} Terminals.")
        log(f"--- {len(ter_li)} Terminals Found ---")
    except Exception as e:
        log(f"Error connecting to AutoCAD: {e}")

def update_status(text):
    status_var.set(text)
    status_bar.update_idletasks()

def log(message):
    log_text.config(state="normal")
    timestamp = datetime.now().strftime("%H:%M:%S")
    log_text.insert("end", f"[{timestamp}] {message}\n")
    log_text.see("end")
    log_text.config(state="disabled")
    log_frame.update_idletasks()

def get_user_input(prompt="Enter value", title="Input Required"):
    """Show a simple pop-up dialog and return the user input as a string."""
    value = simpledialog.askstring(title, prompt)
    return value.upper()


if __name__ == "__main__":
    try:
        #GLOBAL VARIABLES
        glob_units = []
        glob_terminals = {}
        
        root = tk.Tk()
        root.title("AutoCAD Automation Tools")
        root.geometry("600x400")
        root.configure(bg="#f9f9f9")

        # Buttons
        btn_frame = ttk.Frame(root)
        btn_frame.pack(pady=15)

        btn_zf = ttk.Button(btn_frame, text="INPUT ZF", command=input_zf)
        btn_zf.grid(row=1, column=0, padx=10)

        btn_job = ttk.Button(btn_frame, text="check Job Start", command=check_js)
        btn_job.grid(row=1, column=1, padx=10)

        btn_odn = ttk.Button(btn_frame, text="INPUT ODN", command=input_ODN)
        btn_odn.grid(row=1, column=2, padx=10)

        btn_cad = ttk.Button(btn_frame, text="Connect TO CAD", command=Connect2CAD)
        btn_cad.grid(row=0, column=0, padx=10, pady=10)
        root.grid_rowconfigure(0, weight=1)
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

