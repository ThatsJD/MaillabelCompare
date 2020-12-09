from docx import Document

def writeToFile(notpresentdata,document1,tab):
    cols = [1, 3, 5]
    temp = 0
    count = len(notpresentdata)
    while range(len(tab.rows), len(tab.rows) + int(len(notpresentdata) / 3)):
        newRow = tab.add_row()
        colnumber = 1;
        for cell in newRow.cells:
            if colnumber in cols:
                cell.text = notpresentdata[temp][0] + "\n" + notpresentdata[temp][1] + "\n" + notpresentdata[temp][2]
                temp = temp + 1
        if temp == count + 1:
            break
    print("value added", temp)
    colnumber += 1
    document1.save("hello112233.docx")