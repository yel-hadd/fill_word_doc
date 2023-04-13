import pandas as pd
import docx
from replace_word import replace_word_in_docx as rp

def generate_docs(excel_path, word_path):
    df = pd.read_excel(excel_path)
    
    columns = list(df.columns)

    for index, row in df.iterrows():
        doc = docx.Document(word_path)
        for column in columns:
            pattern = f"<<{column}>>"
            outpath = f"./output/{row[columns[0]]} - {row[columns[1]]}.docx"
            for paragraph in doc.paragraphs:
                if pattern in paragraph.text:
                    paragraph.text = paragraph.text.replace(pattern, str(row[column]))
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        if pattern in cell.text:
                            cell.text = cell.text.replace(pattern, str(row[column]))
        doc.save(outpath)
        print(outpath + " has been written")

generate_docs("./excel/path/here", "./word/path/here")

