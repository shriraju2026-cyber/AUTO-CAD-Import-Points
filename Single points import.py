import os
import time
import tkinter as tk
from tkinter import filedialog
import pythoncom
import win32com.client

def complete_cogo_import_and_close():
  print("AutoCAD / Civil 3D check chesthundi...")
  acad = None
  
  try:
    acad = win32com.client.GetActiveObject("AutoCAD.Application")
    print("AutoCAD already open lo undi!")
  except Exception:
    print("AutoCAD open lo ledu. Kothaga open chesthundi, dayachesi wait cheyyandi...")
    try:
      acad = win32com.client.Dispatch("AutoCAD.Application")
      acad.Visible = True
      time.sleep(5)
    except Exception as e:
      print("\n--- ERROR ---")
      print("AutoCAD open cheyyadam lo error:", e)
      input("Window close cheyyadaniki Enter nokkandi...")
      return

  root = tk.Tk()
  root.withdraw()
  root.attributes("-topmost", True)
  
  print("\nStep 1: Dayachesi mee AutoCAD Drawing (.dwg) file select chesukondi...")
  dwg_filename = filedialog.askopenfilename(
      title="Step 1: Select Target Drawing (.dwg) File",
      filetypes=[("AutoCAD Drawing", "*.dwg"), ("All Files", "*.*")],
  )
  
  if not dwg_filename:
    print("Drawing file select cheyyaledu. Process cancelled.")
    input("Window close cheyyadaniki Enter nokkandi...")
    return

  print("\nStep 2: Ippudu mee PNEZD Points (.txt / .csv) file select chesukondi...")
  txt_filename = filedialog.askopenfilename(
      title="Step 2: Select PNEZD Points (.txt / .csv) File",
      filetypes=[("Text Files", "*.txt"), ("CSV Files", "*.csv"), ("All Files", "*.*")],
  )
  
  if not txt_filename:
    print("Points file select cheyyaledu. Process cancelled.")
    input("Window close cheyyadaniki Enter nokkandi...")
    return

  try:
    dwg_path_clean = os.path.abspath(dwg_filename)
    doc = None
    
    for open_doc in acad.Documents:
      if os.path.abspath(open_doc.FullName).lower() == dwg_path_clean.lower():
        doc = open_doc
        doc.Activate()
        print(f"Already open unna '{open_doc.Name}' active aindi.")
        break
    
    if doc is None:
      print(f"'{os.path.basename(dwg_filename)}' open avthundi...")
      doc = acad.Documents.Open(dwg_path_clean)
      time.sleep(3)
  except Exception as e:
    print("\n--- ERROR ---")
    print("Drawing open cheyyadam lo error:", e)
    input("Window close cheyyadaniki Enter nokkandi...")
    return

  civil_doc = None
  cogo_points = None
  for ver in ["13.8", "13.7", "13.6", "13.5", "13.4", "13.3", "13.2", "13.0"]:
    try:
      civil_app = acad.GetInterfaceObject(f"AeccXUiLand.AeccApplication.{ver}")
      civil_doc = civil_app.ActiveDocument
      cogo_points = civil_doc.Points
      break
    except Exception:
      continue

  if cogo_points is None:
    print("\n--- ERROR ---")
    print("Civil 3D API ki connect avvaledu!")
    input("Window close cheyyadaniki Enter nokkandi...")
    return

  target_layer = "PNTS_asbuilt"
  try:
    doc.Layers.Add(target_layer)
  except Exception:
    pass

  print(f"\n'{os.path.basename(txt_filename)}' nundi points import avthunnayi...")
  count = 0
  with open(txt_filename, "r") as file:
    for line in file:
      line = line.strip()
      if not line:
        continue
      parts = [p.strip() for p in line.split(",")]
      if len(parts) >= 4:
        try:
          pt_num = int(parts[0])
          northing = float(parts[1])
          easting = float(parts[2])
          elevation = float(parts[3])
          desc = parts[4] if len(parts) >= 5 else "POINT"

          location = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, [easting, northing, elevation])
          new_pt = cogo_points.Add(location)
          
          try: new_pt.Number = pt_num
          except Exception: pass
            
          new_pt.RawDescription = desc
          new_pt.Layer = target_layer

          try:
            new_pt.Style = "Standard"
            new_pt.LabelStyle = "Standard"
          except Exception: pass

          count += 1
        except ValueError:
          continue

  print(f"Mottham {count} points import ayyayi.")

  print("\nScale 1:1, Layer OFF mariyu Save avthundi...")
  
  try:
    doc.SetVariable("CANNOSCALE", "1:1")
  except Exception:
    pass

  try:
    if doc.GetVariable("CLAYER").lower() == target_layer.lower():
      doc.SetVariable("CLAYER", "0")
    layer = doc.Layers.Item(target_layer)
    layer.LayerOn = False
  except Exception:
    pass

  # Save the drawing
  doc.SendCommand("_AeccUpdateAllPointGroups\n_REGENALL\n_ZOOM _E\n_QSAVE\n")
  
  # Save avvadaniki 3 seconds aaguthundi
  print("Save avthundi, wait 3 seconds...")
  time.sleep(3)
  
  # ONLY Drawing matrame close cheyyadam (AutoCAD open lone untundi)
  try:
      doc.Close(False)
      print(f"\nSuccess! '{os.path.basename(dwg_filename)}' save ayyi close aipoyindi.")
  except Exception as e:
      print("\nDrawing close cheyadam lo error:", e)

  print("\nAnni panulu poorthi ayyayi! Window close cheyyadaniki Enter nokkandi...")
  input()

if __name__ == "__main__":
  complete_cogo_import_and_close()