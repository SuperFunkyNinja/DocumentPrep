import fitz
import os
import openpyxl

os.chdir('E:/Desktop')

arr = os.listdir()

for i in arr:
  if '.pdf' in i or '.PDF' in i:
    pos = i.find('.pdf')
    excelFileName = i[0:pos]+'.xlsx'
    wb = openpyxl.Workbook()
    sheet = wb.active
    doc = fitz.open(i)
    rowNum = 1

    for j in range(doc.pageCount):
      page = doc[j]
      for annot in page.annots():
        if annot.info["content"] != '':
          print(annot.info["content"])
          sheet.cell(row = rowNum, column = 1).value = annot.info["content"]
          rowNum += 1
    
    wb.save(excelFileName)
    wb.close()