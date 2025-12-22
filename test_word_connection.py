# -*- coding: utf-8 -*-
import win32com.client
import pythoncom

try:
    # Initialize COM
    pythoncom.CoInitialize()
    
    print("Intentando conectar con Word...")
    word = win32com.client.GetActiveObject("Word.Application")
    print("Word encontrado!")
    print(f"  Version: {word.Version}")
    print(f"  Visible: {word.Visible}")
    print(f"  Documentos abiertos: {word.Documents.Count}")
    
    # Try to list all documents
    if word.Documents.Count > 0:
        print("\nDocumentos encontrados:")
        for i, doc in enumerate(word.Documents, 1):
            print(f"  {i}. {doc.Name}")
            print(f"     Ruta: {doc.FullName}")
    else:
        print("  No hay documentos abiertos")
        print("\nIntentando informacion adicional:")
        print(f"  Windows Count: {word.Windows.Count}")
        
        # Check if there are any windows
        if word.Windows.Count > 0:
            print("\nVentanas de Word abiertas:")
            for i, win in enumerate(word.Windows, 1):
                print(f"  {i}. {win.Caption}")
    
    pythoncom.CoUninitialize()
    
except Exception as e:
    print(f"Error: {e}")
    import traceback
    traceback.print_exc()