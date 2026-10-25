import os
import time
import tkinter as tk
from tkinter import filedialog
import pythoncom
import win32com.client

def batch_process_with_popup():
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
    
    print("\nStep 1: Mee .dwg files unna Folder ni select chesukondi...")
    folder_path = filedialog.askdirectory(title="Select Folder with DWG files")
    
    if not folder_path:
        print("Folder select cheyyaledu. Process cancelled.")
        input("Window close cheyyadaniki Enter nokkandi...")
        return

    dwg_files = [f for f in os.listdir(folder_path) if f.lower().endswith('.dwg')]
    
    if not dwg_files:
        print("Aa folder lo elanti .dwg files levu!")
        input("Window close cheyyadaniki Enter nokkandi...")
        return

    print(f"\nMottham {len(dwg_files)} DWG files kanipinchayi. Batch processing modalavthundi...\n")
    
    success_count = 0
    target_layer = "PNTS_asbuilt"

    for dwg_file in dwg_files:
        base_name = os.path.splitext(dwg_file)[0]
        txt_file = base_name + ".txt"
        csv_file = base_name + ".csv"
        
        points_full_path = None
        
        # 1. Exact match kosam vethukuthundi
        if os.path.exists(os.path.join(folder_path, txt_file)):
            points_full_path = os.path.join(folder_path, txt_file)
        elif os.path.exists(os.path.join(folder_path, csv_file)):
            points_full_path = os.path.join(folder_path, csv_file)
        else:
            # 2. Match dorakakapothe Manual Popup aduguthundi
            print(f"\n[!] '{dwg_file}' ki exact points file dorakaledu.")
            points_full_path = filedialog.askopenfilename(
                title=f"Select Points file for: {dwg_file}",
                initialdir=folder_path,
                filetypes=[("Text/CSV Files", "*.txt *.csv"), ("All Files", "*.*")]
            )
            
            if not points_full_path:
                print(f"Skipped: '{dwg_file}' - Manual ga file select cheyyaledu.")
                continue

        dwg_full_path = os.path.join(folder_path, dwg_file)
        print(f"\n---> Process chesthundi: {dwg_file} & {os.path.basename(points_full_path)}")
        
        doc = None
        try:
            for open_doc in acad.Documents:
                if os.path.abspath(open_doc.FullName).lower() == os.path.abspath(dwg_full_path).lower():
                    doc = open_doc
                    doc.Activate()
                    break
            
            if doc is None:
                doc = acad.Documents.Open(dwg_full_path)
                time.sleep(3)
        except Exception as e:
            print(f"Error: '{dwg_file}' open cheyyaleka poyindi: {e}")
            continue

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
            print("Error: Civil 3D API ki connect avvaledu!")
            doc.Close(False)
            continue

        try:
            doc.Layers.Add(target_layer)
        except Exception:
            pass

        points_added = 0
        with open(points_full_path, "r") as file:
            for line in file:
                line = line.strip()
                if not line: continue
                parts = [p.strip() for p in line.split(",")]
                if len(parts) >= 4:
                    try:
                        pt_num = int(parts[0])
                        northing, easting, elevation = float(parts[1]), float(parts[2]), float(parts[3])
                        desc = parts[4] if len(parts) >= 5 else "POINT"

                        location = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, [easting, northing, elevation])
                        new_pt = cogo_points.Add(location)
                        
                        try: new_pt.Number = pt_num
                        except: pass
                        
                        new_pt.RawDescription = desc
                        new_pt.Layer = target_layer

                        try:
                            new_pt.Style = "Standard"
                            new_pt.LabelStyle = "Standard"
                        except: pass

                        points_added += 1
                    except ValueError:
                        continue
        
        print(f"{points_added} points import ayyayi.")

        try: doc.SetVariable("CANNOSCALE", "1:1")
        except: pass

        try:
            if doc.GetVariable("CLAYER").lower() == target_layer.lower():
                doc.SetVariable("CLAYER", "0")
            doc.Layers.Item(target_layer).LayerOn = False
        except: pass

        doc.SendCommand("_AeccUpdateAllPointGroups\n_REGENALL\n_ZOOM _E\n_QSAVE\n")
        time.sleep(3)
        
        try:
            doc.Close(False)
            print(f"Success! '{dwg_file}' save ayyi close aipoyindi.")
            success_count += 1
        except Exception as e:
            print(f"Close error: {e}")

    print(f"\n=======================================")
    print(f"Batch Processing Complete!")
    print(f"Mottham {success_count} files successfully update ayyayi.")
    print(f"=======================================")
    input("Window close cheyyadaniki Enter nokkandi...")

if __name__ == "__main__":
    batch_process_with_popup()