import win32com.client as win32

# Set the file paths for the Word document and PowerPoint presentation
word_path = "C:/Users/Naman/Desktop/assignment.docx"
ppt_path = "C:/Users/Naman/Desktop/assignment.pptx"

# Set up the Word and PowerPoint applications
word = win32.gencache.EnsureDispatch('Word.Application')
ppt = win32.gencache.EnsureDispatch('PowerPoint.Application')

# Try to open the Word document
try:
    doc = word.Documents.Open(word_path)
except Exception as e:
    print("Error opening Word document:", e)
    word.Quit()
    ppt.Quit()
    exit()

# Try to copy the contents of the Word document
try:
    doc.Content.Copy()
except Exception as e:
    print("Error copying content from Word document:", e)
    doc.Close()
    word.Quit()
    ppt.Quit()
    exit()

# Try to create a new PowerPoint presentation
try:
    pres = ppt.Presentations.Add()
except Exception as e:
    print("Error creating PowerPoint presentation:", e)
    doc.Close()
    word.Quit()
    ppt.Quit()
    exit()

# Try to paste the contents into a new slide in the PowerPoint presentation
try:
    slide = pres.Slides.Add(1, 11)
    slide.Shapes.PasteSpecial(DataType=2)
except Exception as e:
    print("Error pasting content into PowerPoint slide:", e)
    doc.Close()
    pres.Close()
    word.Quit()
    ppt.Quit()
    exit()

# Try to save the PowerPoint presentation
try:
    #pres.SaveAs(ppt_path + '\\output.pptx')

    pres.SaveAs(ppt_path)
except Exception as e:
    print("Error saving PowerPoint presentation:", e)

# Clean up
doc.Close()
pres.Close()
word.Quit()
ppt.Quit()

