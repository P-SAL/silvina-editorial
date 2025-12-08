import win32com.client

print("Intentando conectar con Word...")

try:
    word = win32com.client.GetActiveObject("Word.Application")
    print(f"✓ Word encontrado!")
    print(f"  Versión: {word.Version}")
    print(f"  Documentos abiertos: {word.Documents.Count}")
    
    if word.Documents.Count > 0:
        doc = word.ActiveDocument
        print(f"  Documento activo: {doc.Name}")
        print(f"  Caracteres: {doc.Characters.Count}")
    else:
        print("  No hay documentos abiertos")
        
except Exception as e:
    print(f"✗ Error: {e}")
    print(f"  Tipo de error: {type(e).__name__}")