import win32com.client
import pythoncom
import time

pythoncom.CoInitialize()

print("Opening Word...")
word = win32com.client.Dispatch("Word.Application")
word.Visible = True

filepath = r"C:\Users\usuario\Desktop\Calderon LA GUERRA DE UCRANIA EN CLAVE OPERACIONAL.docx"
print(f"Opening document: {filepath}")
doc = word.Documents.Open(filepath)

time.sleep(2)
print("Getting text...")
text = doc.Content.Text
print(f"✅ Success! Got {len(text)} characters")

doc.Close(SaveChanges=False)
word.Quit()
print("✅ Closed")