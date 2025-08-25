import pandas as pd
from docx import Document
import os

df = pd.read_excel("C:\\vacation.xlsx") # paste Your data source, could be excel or whatever
template_path = "C:\\vacation.docx" # add Your document in word (template)
output_folder = "C:\\holiday_requests" #create a folder and paste the source

os.makedirs(output_folder, exist_ok=True)

def fill_template(row, template_path, output_docx_path): # function to fill the templete
    doc = Document(template_path)
    for p in doc.paragraphs:
        p.text = p.text.replace("Name", row["Name"])
        p.text = p.text.replace("Surname", row["Surname"])
        p.text = p.text.replace("Position", row["Position"])
        p.text = p.text.replace("Line_manager", row["Line_manager"])
        p.text = p.text.replace("Start_date", str(row["Start_date"].date()))
        p.text = p.text.replace("End_date", str(row["End_date"].date()))
        p.text = p.text.replace("Total_vacation_days", str(row["Total_vacation_days"]))
        p.text = p.text.replace("Type", row["Type"])
    doc.save(output_docx_path)

for index, row in df.iterrows(): # check every part of data
    name = f"{row['Name']}_{row['Surname']}"
    docx_path = os.path.join(output_folder, f"{name}_holiday_request.docx")
    fill_template(row, template_path, docx_path)

print("Great, we did it!")
