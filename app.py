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
import tempfile, base64


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
    miss=0
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
            miss +=1
            log(f"Terminal ID {l_s} not found in ODN EXCEL.")
            continue
    log(f"---- Done ODN Updated for {ter_id_li.len()-miss} out of {ter_id_li.len()} ----")
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

def load_icon_from_memory():
    """Save the base64 icon to a temp file and return its path."""
    icon_temp = tempfile.NamedTemporaryFile(delete=False, suffix=".ico")
    icon_temp.write(base64.b64decode(ICON_DATA))
    icon_temp.close()
    return icon_temp.name


if __name__ == "__main__":
    try:
        #GLOBAL VARIABLES
        glob_units = []
        glob_terminals = {}
        ICON_DATA = b"""AAABAAEAAAAAAAEAIACXHAAAFgAAAIlQTkcNChoKAAAADUlIRFIAAAEAAAABAAgGAAAAXHKoZgAAHF5JREFUeNrt3XucXGV9x/HPc+bM7oaQBAMCGm4FvLE7y0Wh1YbsbkCpVZTy0vJSqlSLtrblopXepIoKWISiraJYRVGhiKhVUCio7G6gVJEUyIZbRRC5igaRhCQ7c+Y8/eN3Znay5LK3OZfZ79tXXruLyc65PM/vPOe5/B4QEREREREREREREREREREREREREREREREREREREREREREREREREREREREREREREREREREREREREREREREREREREREREREREREREREREREREREREREREREREREREREREREREREREREREREREREREREREREREREREWkvp0sgRVAd6cfhgHha/84nX7sGx3QRFQAkz+o/rBDUSkQ99W1XZ+8CHCUgSMpvCSgnXx1QByKgCtTAx5OLufPgHZQVFBQA5lptpJLZZxetQPtbXkmtuh7n3MSj2ipo6B27AHsCLwT2Sb7uDuwCLAJ6gBDoSr4vJ+U5AsaB9cCvgceAnwP3Aw8Aj3vHBudbaoCfv8EgVJVti6VAL/aUarcIWJMU+EJoBMmouiFp1hMklb0PeIV3HAYcmFT6xUD3HHzsZmAdcL/zrAZuBlYDjwBxbaSSNA2gPLh23hRUtQDaU7hfDVyVFFzf5vv3NPB6YHWen2LVVf0EcYxPilwYx0RBsA+wHDgGOALYF1iQ1q0CHgJuAq5Jvv66cVE986NVoBZAe5SSgtyVwmctIJ2WxoyMj1YIPBD7RuXvAY6IguBNScU/ILleaStjrYwDgROxVtRVwDfrzv8s8I7qSAXX4YFAAaA9PO198k/+rFyqjVRaj24BcBTwTmAlsCRHh9oFvCL58+7Au8uBLzvrM7DzcI7ywJqOK6i5fXJIcVVHKvgbeu0HTwkYBP4DuBL4o5xV/skOAD4IXAuchvXngPeZdvAqAEgh1JJmc9QVAOyL4wLgm8BxwE4FOpWXABdirwUrSfrLaiMVNg0fogAg0spfe2DrEzIA3gh8GzidxlO0mPVjJfB14J+wIUi6NnZTG+5XABABiEb7iXZqdt4vBj4EfBk4pENOcVfgLOBS4GXxwo3gOuOVQAFAZlf5Vx1EPDFxby/gYuBM8v2ePxMOa9VciXVmAhQ+CCgA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25biDs9fwg0RERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERER2aH/B7kPQH5nsuwQAAAAAElFTkSuQmCC"""

        root = tk.Tk()
        root.title("Job Starter")
        root.geometry("600x400")
        root.configure(bg="#f9f9f9")
        icon_path = load_icon_from_memory()
        root.iconbitmap(icon_path)

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
        input("press any key to exit...")

