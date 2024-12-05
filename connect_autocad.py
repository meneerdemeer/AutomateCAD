import win32com.client
import sys

def connect_to_autocad():
    try:
        # Get the AutoCAD application
        acad = win32com.client.Dispatch("AutoCAD.Application")
        
        # Get the active document (drawing)
        doc = acad.ActiveDocument
        
        print("Successfully connected to AutoCAD!")
        print(f"AutoCAD Version: {acad.Version}")
        print(f"Active Document: {doc.Name}")
        
        return acad, doc
    
    except Exception as e:
        print("Error connecting to AutoCAD:")
        print(str(e))
        print("\nMake sure AutoCAD is running before executing this script.")
        sys.exit(1)

if __name__ == "__main__":
    acad, doc = connect_to_autocad() 