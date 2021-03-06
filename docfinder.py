from docx2python import docx2python
from docx import Document
from docx import table

from functions import writeToFile

document1 = docx2python("D:\\PycharmProjects\\docfilefinder\\MASTERRYANLABELFILE.docx")
document2 = docx2python("D:\\PycharmProjects\\docfilefinder\\RYANSTORELABELS.docx")

doc2body = document2.body
body = document1.body
count = 0
doc1list = dict()
for innerbody in body:
    for row in innerbody:
        if row[0][0] == "":
            continue
        doc1list[row[0][0]] = row[0]
        doc1list[row[2][0]] = row[2]
        doc1list[row[4][0]] = row
        count += 1
print("Document 1: {MASTERRYANLABELFILE} List", len(doc1list), "Count : ", count, "Check", count * 3)

# for Document 2

count = 0
doc2list = dict()
for innerbody in doc2body:
    for row in innerbody:
        if row[0][0] == "":
            continue
        doc2list[row[0][0]] = row[0]
        doc2list[row[2][0]] = row[2]
        doc2list[row[4][0]] = row[4]
        count += 1

notpresent = list()
present = list()
for record in doc2list.values():
    if doc1list.get(record[0]):
        for row in doc1list:
            del row[3]
            del row[4]
            del row[5]
            present.append(record)
    else:
        notpresent.append(record)

document1 = Document(docx="D:\\PycharmProjects\\docfilefinder\\MASTERRYANLABELFILE.docx")
tables = document1.tables
for tab in tables:
    print("Columns: ", len(tab.columns), "Rows: ", len(tab.rows))
tab = tables[0]

writeToFile(notpresent, document1, tab)



