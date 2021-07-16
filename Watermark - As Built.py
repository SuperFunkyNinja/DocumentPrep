import fitz
import os

os.chdir("E:/Desktop")

arr = os.listdir()

for i in arr:
    if ".pdf" in i and "VDRL" in i:
        pos = i.find(" ")

        doc = fitz.open(i)

        print(i[5:pos])

        # print('Revision number:')

        # revNo = input()
        revNo = "00"

        # print('Start page:')

        # startPage = int(input())
        startPage = 4

        p = fitz.Point(6, 13)
        text = str(
            "Contractor Document Number: "
            + i[5:pos]
            + "\nContractor Document Revision: "
            + revNo.upper()
        )

        for j in range(startPage - 1, len(doc)):
            page = doc[j]

            if not page._isWrapped:
                page._wrapContents()

            page.insertText(
                p,  # bottom-left of 1st char
                text,  # the text (honors '\n')
                fontsize=10,  # the default font size
            )

        doc.save(i[5:pos] + ".pdf")
