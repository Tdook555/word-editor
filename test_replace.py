from docx import Document
import shutil

shutil.copy("uploads/doc_test.docx", "uploads/test_output2.docx")
doc = Document("uploads/test_output2.docx")

for para in doc.paragraphs:
    for run in para.runs:
        print(f"before: {repr(run.text)}")
        if "2568" in run.text:
            run.text = run.text.replace("2568", "2570")
            print(f"after: {repr(run.text)}")

doc.save("uploads/test_output2.docx")
print("saved")