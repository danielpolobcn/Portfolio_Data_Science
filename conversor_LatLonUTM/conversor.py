#### Importacions
from pyproj import Proj
import re
import numpy as np
import pandas as pd
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog as fd

# Funció per convertir les comes en punts en les entrades de text
def change_commas (text):
    return text.replace(",", ".")

# Funció per transformar UTM i zona UTM a lat/lon (ETRS89, elipse WGS80)
def utm_to_latlon(utm_x, utm_y, utmZone):
    if utm_x == "NaN" and "utm_y" == "NaN" and "utmZone" == "NaN":
        return ("NaN", "NaN")
    else:
        utm_proj = Proj(proj='utm', zone=utmZone, ellps='WGS84')
        lon, lat = utm_proj(utm_x, utm_y, inverse=True)
        return (lat, lon)

# Funció per transformar lat/lon a UTM i zona UTM (ETRS89, elipse WGS80)
def latlon_to_utm(lat, lon):
    if lat == "NaN" and lon == "NaN":
        return ("NaN", "NaN", "NaN")
    else:
        utmZone = int(((lon + 180)/6) + 1)
        transformer = Proj(proj='utm', zone=utmZone, ellps='GRS80')
        utm_x, utm_y = transformer(lon, lat)
        return (utm_x, utm_y, utmZone)

# Funció per als botons "Copy"
def copy_button (id):
    root.clipboard_clear()
    if id == "utmx":
        root.clipboard_append(showUTMx["text"])
    if id == "utmy":
        root.clipboard_append(showUTMy["text"])
    if id == "utmfull":
        root.clipboard_append(showUTMfull["text"])
    if id == "lat":
        root.clipboard_append(showLat["text"])
    if id == "lon":
        root.clipboard_append(showLon["text"])

# Funció per seleccionar i carregar l'arxiu Excel
def select_file():
    filetypes = (('Excel files', '*.xlsx'), ('All files', '*.*'))
    filename = fd.askopenfilename(title='Obrir arxiu', initialdir='/', filetypes=filetypes)
    split = re.search("[^/]+$", filename)
    global name
    name = re.split("\.", split.group())[0]
    global fileDir
    fileDir = filename[0:split.span()[0]]
    global df
    df = pd.read_excel(filename)
    global columns
    columnsDF = df.columns.values
    columns = [str(item) for item in columnsDF]
    columns.sort()
    columns = tuple(columns)
    text_columns()
    combo1["values"] = columns
    combo2["values"] = columns
    combo3["values"] = columns

# Funció per etiquetar la selecció dels camps de l'Excel
def text_columns ():
    output = selectedOutput.get()
    if output == "utm":
        chooseColumn1["text"] = "Latitud:"
        chooseColumn2["text"] = "Longitud:"
        chooseColumn3.grid_forget()
        combo3.grid_forget()
    if output == "latlon":
        chooseColumn1["text"] = "UTM X:"
        chooseColumn2["text"] = "UTM Y:"
        chooseColumn3["text"] = "Zona UTM:"
        chooseColumn3.grid(row=3, column=0, sticky="w", padx=5, pady=5)
        combo3.grid(row=3, column=1, sticky="w", padx=5, pady=5)
    frameChooseColumns.grid(row=2, column=0, sticky="w")

# Funció per adequar el format de les columnes del Dataframe
def clean_columns (output, *args):
    def cleaner (item):
        if item == "NaN":
            return "NaN"
        else:
            replaced = str(item).replace(",", ".")
            return float(replaced)
    def make_int (item):
        if item == "NaN":
            return "NaN"
        else:
            return int(item)
    df[args[0]] = df[args[0]].apply(cleaner)
    df[args[1]] = df[args[1]].apply(cleaner)
    if output == "latlon":
        df[args[2]] = df[args[2]].apply(cleaner)
        df[args[2]] = df[args[2]].apply(make_int)    

# Funció per seleccionar i aplicar la conversió al Excel
def convert_coords ():
    def extract_coords (item, column):
        if item[0] == "NaN":
            return np.nan
        if column == "lat":
            return str(round(item[0], 7)).replace(".", ",")
        if column == "lon":
            return str(round(item[1], 7)).replace(".", ",")
        if column == "utm_x":
            return str(int(item[0]))
        if column == "utm_y":
            return str(int(item[1]))
        if column == "utm_zona":
            return str(int(item[2]))
    def deliver_coords ():
        outputName = fileDir + name + "_conversio.xlsx"
        df.to_excel(outputName, index=False)
        statusText["state"] = "normal"
        statusText.delete("1.0", "end")
        statusText.insert("1.0", "Procés finalitzat\nArxiu desat com: "+name+"_conversio.xlsx")
        statusText["state"] = "disabled"
        frameStatus.grid(row=3, column=0, sticky="w")
    def deliver_error ():
        statusText["state"] = "normal"
        statusText.delete("1.0", "end")
        statusText.insert("1.0", "ERROR, les columnes indicades contenen valors incorrectes")
        statusText["state"] = "disabled"
        frameStatus.grid(row=3, column=0, sticky="w")
    frameStatus.grid(row=3, column=0, sticky="w")
    output = selectedOutput.get()
    global df
    if output == "latlon":
        df["x_"] = df[combo1_option.get()]        
        df["x_"] = df["x_"].replace(np.nan, "NaN")
        df["y_"] = df[combo2_option.get()]
        df["y_"] = df["y_"].replace(np.nan, "NaN")
        utmZone = combo3_option.get()
        if utmZone in columns:
            df["z_"] = df[utmZone]
            df["z_"] = df["z_"].replace(np.nan, "NaN")
        else:
            df["z_"] = utmZone
        try:
            clean_columns(output, "x_", "y_", "z_")
            df["conversio_"] = df.apply(lambda x: utm_to_latlon(x.x_, x.y_, x.z_), axis=1)
            df["lat_"] = df["conversio_"].apply(extract_coords, args=("lat",))
            df["lon_"] = df["conversio_"].apply(extract_coords, args=("lon",))   
            df = df.drop(['x_', 'y_', 'z_', 'conversio_'], axis=1)
            deliver_coords()
        except:
            deliver_error()
    if output == "utm":
        df["lat_"] = df[combo1_option.get()]
        df["lat_"] = df["lat_"].replace(np.nan, "NaN")
        df["lon_"] = df[combo2_option.get()]
        df["lon_"] = df["lon_"].replace(np.nan, "NaN")
        try:
            clean_columns(output, "lat_", "lon_")
            df["conversio_"] = df.apply(lambda x: latlon_to_utm(x.lat_, x.lon_), axis=1)
            df["utm_x_"] = df["conversio_"].apply(extract_coords, args=("utm_x",))
            df["utm_y_"] = df["conversio_"].apply(extract_coords, args=("utm_y",))
            df["utm_zona_"] = df["conversio_"].apply(extract_coords, args=("utm_zona",))
            df = df.drop(['lat_', 'lon_', 'conversio_'], axis=1)
            deliver_coords()
        except:
            deliver_error()

#### TKinter GUI

root = tk.Tk()

root.title('Conversor de coordenades UTM - LatLon')
window_width = 600
window_height = 500
copyButton = tk.PhotoImage(file='./assets/copyButton.png')

# Presentem l'aplicació centrada en la pantalla
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
center_x = int(screen_width/2 - window_width / 2)
center_y = int(screen_height/2 - window_height / 2)
root.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')
# Presentem l'aplicació per sobre de les altres
root.attributes("-topmost", 1)

# Notebook
notebook = ttk.Notebook(root)

#Input Frame
frameInput = ttk.Frame(notebook)
frameInput.pack(padx=3, pady=3, anchor="nw")

#LatLon Frame
frameLatLon = ttk.Frame(frameInput, borderwidth=10, relief="groove")
frameLatLon.pack(side="left", anchor="n")
#Lat Label + Entry
latLabelLatLon= tk.Label(frameLatLon, text="Latitud:")
latLabelLatLon.grid(row=0, column=0, sticky="w", padx=10, pady=5)
latEntry = tk.Entry(frameLatLon)
latEntry.grid(row=0, column=1, padx=10, pady=5)
#Lon Label + Entry
lonLabelLatLon= tk.Label(frameLatLon, text="Longitud:")
lonLabelLatLon.grid(row=1, column=0, sticky="w", padx=10, pady=5)
lonEntry = tk.Entry(frameLatLon)
lonEntry.grid(row=1, column=1, padx=10, pady=5)
#Botó Conversió LatLon a UTM 
def conversio_LatLon():
    lat = change_commas(latEntry.get())
    lon = change_commas(lonEntry.get())
    try:
        lat = float(lat)
        lon = float(lon)
        result = latlon_to_utm(lat, lon)
        showUTMx["text"] = str(int(result[0]))
        showUTMy["text"] = str(int(result[1]))
        showUTMfull["text"] = str(result[2])
    except:
        showUTMx["text"] = "Error en input!"
        showUTMy["text"] = "Error en input!"    
        showUTMfull["text"] = "Error en input!"
botoConversioLatLon = tk.Button(frameLatLon, text="Conversió",\
    font=("TkDefaultFont", 8, "bold"), command=conversio_LatLon)
botoConversioLatLon.grid(row=2, column=1, padx=10, pady=10)
#UTM X Label + Text
utmxLabelLatLon= tk.Label(frameLatLon, text="UTM X:")
utmxLabelLatLon.grid(row=3, column=0, sticky="w", padx=10, pady=5)
showUTMx = tk.Label(frameLatLon)
showUTMx.grid(row=3, column=1, sticky="e", padx=10, pady=5)
copyUTMx = tk.Button(frameLatLon, image=copyButton, command=lambda: copy_button("utmx"))
copyUTMx.grid(row=3, column=2, sticky="e", padx=3, pady=2)
#UTM Y Label
utmyLabelLatLon= tk.Label(frameLatLon, text="UTM Y:")
utmyLabelLatLon.grid(row=4, column=0, sticky="w", padx=10, pady=5)
showUTMy = tk.Label(frameLatLon)
showUTMy.grid(row=4, column=1, sticky="e", padx=10, pady=5)
copyUTMy = tk.Button(frameLatLon, image=copyButton, command=lambda: copy_button("utmy"))
copyUTMy.grid(row=4, column=2, sticky="e", padx=3, pady=2)
#UTM Full Label
utmfullLabelLatLon= tk.Label(frameLatLon, text="Zona UTM:")
utmfullLabelLatLon.grid(row=5, column=0, sticky="w", padx=10, pady=5)
showUTMfull = tk.Label(frameLatLon)
showUTMfull.grid(row=5, column=1, sticky="e", padx=10, pady=5)
copyUTMfull = tk.Button(frameLatLon, image=copyButton, command=lambda: copy_button("utmfull"))
copyUTMfull.grid(row=5, column=2, sticky="e", padx=3, pady=2)

#UTM Frame
frameUTM = ttk.Frame(frameInput, borderwidth=10, relief="groove")
frameUTM.pack(side="left", anchor="n")
#UTM X Label + Entry
utmxLabelUTM= tk.Label(frameUTM, text="UTM X:")
utmxLabelUTM.grid(row=0, column=0, sticky="w", padx=10, pady=5)
utmxEntry = tk.Entry(frameUTM)
utmxEntry.grid(row=0, column=1, padx=10, pady=5)
#UTM Y Label + Entry
utmyLabelUTM= tk.Label(frameUTM, text="UTM Y:")
utmyLabelUTM.grid(row=1, column=0, sticky="w", padx=10, pady=5)
utmyEntry = tk.Entry(frameUTM)
utmyEntry.grid(row=1, column=1, padx=10, pady=5)
#UTM Full Label + Entry
utmfullLabelUTM= tk.Label(frameUTM, text="Zona UTM:")
utmfullLabelUTM.grid(row=2, column=0, sticky="w", padx=10, pady=5)
utmfullEntry = tk.Entry(frameUTM)
utmfullEntry.grid(row=2, column=1, padx=10, pady=5)
#Botó Conversió UTM a LatLon 
def conversio_UTM():
    utmx = change_commas(utmxEntry.get())
    utmy = change_commas(utmyEntry.get())
    utmfull = utmfullEntry.get()
    try:
        utmx = float(utmx)
        utmy = float(utmy)
        utmfull = int(utmfull)
        result = utm_to_latlon(utmx, utmy, utmfull)
        showLat["text"] = str(round(result[0], 6))
        showLon["text"] = str(round(result[1], 6))
    except:
        showLat["text"] = "Error en input!"
        showLon["text"] = "Error en input!"
botoConversioUTM = tk.Button(frameUTM, text="Conversió",\
    font=("TkDefaultFont", 8, "bold"), command=conversio_UTM)
botoConversioUTM.grid(row=3, column=1, padx=10, pady=10)
#Lat Label + Text
latLabelUTM= tk.Label(frameUTM, text="Latitud:")
latLabelUTM.grid(row=4, column=0, sticky="w", padx=10, pady=5)
showLat = tk.Label(frameUTM)
showLat.grid(row=4, column=1, sticky="e", padx=10, pady=5)
copyLat = tk.Button(frameUTM, image=copyButton, command=lambda: copy_button("lat"))
copyLat.grid(row=4, column=2, sticky="e", padx=3, pady=2)
#Lon Label + Text
lonLabelUTM= tk.Label(frameUTM, text="Longitud:")
lonLabelUTM.grid(row=5, column=0, sticky="w", padx=10, pady=5)
showLon = tk.Label(frameUTM)
showLon.grid(row=5, column=1, sticky="e", padx=10, pady=5)
copyLon = tk.Button(frameUTM, image=copyButton, command=lambda: copy_button("lon"))
copyLon.grid(row=5, column=2, sticky="e", padx=3, pady=2)

# Excel Frame
frameExcel = ttk.Frame(notebook)
frameExcel.pack(padx=3, pady=3, anchor="nw")


# Choose Output Frame
frameChooseOutput = ttk.Frame(frameExcel, borderwidth=5, relief="flat")
frameChooseOutput.grid(row=0, column=0, sticky="w")
outputLabel = tk.Label(frameChooseOutput, text="Sistema de coordinades de sortida:")
outputLabel.grid(row=0, column=0, sticky="w", padx=5, pady=5)
# Choose Output Radio Button
selectedOutput = tk.StringVar()
selectUTM = ttk.Radiobutton(frameChooseOutput, text="UTM", value="utm", variable=selectedOutput)
selectUTM.grid(row=1, column=0, sticky="w", padx=5, pady=5)
selectLatLon = ttk.Radiobutton(frameChooseOutput, text="Latitud / Longitud", value="latlon", variable=selectedOutput)
selectLatLon.grid(row=1, column=1, sticky="w", padx=5, pady=5)
# GetFile Frame
frameGetFile = ttk.Frame(frameExcel, borderwidth=5, relief="flat")
frameGetFile.grid(row=1, column=0, sticky="w")
# Choose File Button
botoGetFile = tk.Button(frameGetFile, text="Seleccionar arxiu",\
    font=("TkDefaultFont", 8, "bold"), command=select_file)
botoGetFile.pack(side="left", anchor="nw", padx=10, pady=10)

# Choose Columns Frame
frameChooseColumns = ttk.Frame(frameExcel, borderwidth=5, relief="flat")
frameChooseColumns.grid(row=2, column=0, sticky="w")
# Common Label
labelColumns = tk.Label(frameChooseColumns, text="Indica quines columnes contenen els camps necessaris:")
labelColumns.grid(row=0, column=0, sticky="w", padx=5, pady=5)
# Column 1
chooseColumn1 = tk.Label(frameChooseColumns, text=" ")
chooseColumn1.grid(row=1, column=0, sticky="w", padx=5, pady=5)
combo1_option = tk.StringVar()
combo1 = ttk.Combobox(frameChooseColumns, textvariable=combo1_option, state="readonly")
combo1.grid(row=1, column=1, sticky="w", padx=5, pady=5)
# Column 2
chooseColumn2 = tk.Label(frameChooseColumns, text= " ")
chooseColumn2.grid(row=2, column=0, sticky="w", padx=5, pady=5)
combo2_option = tk.StringVar()
combo2 = ttk.Combobox(frameChooseColumns, textvariable=combo2_option, state="readonly")
combo2.grid(row=2, column=1, sticky="w", padx=5, pady=5)
# Column 3
chooseColumn3 = tk.Label(frameChooseColumns, text=" ")
chooseColumn3.grid(row=3, column=0, sticky="w", padx=5, pady=5)
combo3_option = tk.StringVar()
combo3 = ttk.Combobox(frameChooseColumns, textvariable=combo3_option)
combo3.grid(row=3, column=1, sticky="w", padx=5, pady=5)
# Convert Button
botoConversio = tk.Button(frameChooseColumns, text="Conversió",\
    font=("TkDefaultFont", 8, "bold"), command=convert_coords)
botoConversio.grid(row=4, column=0, sticky="w", padx=5, pady=5)
frameChooseColumns.grid_remove()

# Status Frame
frameStatus = ttk.Frame(frameExcel, borderwidth=5, relief="flat")
frameStatus.grid(row=3, column=0, sticky="w")
# Status Label
statusLabel = tk.Label(frameStatus, text="Status:")
statusLabel.grid(row=0, column=0, sticky="w", padx=5, pady=3)
# Status Text
statusText = tk.Text(frameStatus, height=3, width=70)
statusText.grid(row=1, column=0, sticky="w", padx=5, pady=3)
statusText.insert("1.0", "Processant...")
statusText["state"] = "disabled"
frameStatus.grid_remove()

# Afegim pestanyes al Notebook
notebook.add(frameInput, text=" Coordenades ")
notebook.add(frameExcel, text=" Arxiu Excel ")
notebook.pack(anchor="nw", expand=True, padx=3, pady=3)

root.mainloop()

