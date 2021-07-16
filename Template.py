import os, fitz, openpyxl

TEMPLATE = "F:\\Engineering\\Eli Saunders\\Python\\"
DRAWINGS = "E:\\Desktop"

os.chdir(DRAWINGS)
wb = openpyxl.load_workbook(TEMPLATE + "array.xlsx")

sheet = wb["Sheet1"]
descriptions = []
numbers = []
files = os.listdir(DRAWINGS)

for i in range(1, sheet.max_column + 1):
    descriptions.append(sheet.cell(1, i).value)

for i in range(2, sheet.max_row + 1):
    numbers.append(sheet.cell(i, 1).value)

place = fitz.Rect(25, 25, 1200, 950)


def getPoint(search):

    sheet = wb["Sheet2"]
    points = {}

    for row in range(2, sheet.max_row + 1):

        ref = sheet["A" + str(row)].value
        x = sheet["B" + str(row)].value
        y = sheet["C" + str(row)].value
        f = sheet["D" + str(row)].value

        points.setdefault(ref, {"x": 0, "y": 0, "f": 0})

        points[ref]["x"] += int(x)
        points[ref]["y"] += int(y)
        points[ref]["f"] += int(f)

    return points[search]


def getString(index, search):

    sheet = wb["Sheet1"]
    strings = {}

    for column in range(1, sheet.max_column + 1):
        ref = sheet.cell(1, column).value
        string = sheet.cell(index, column).value

        strings[ref] = string

    return strings[search]


def getPage(index, pix, pageNo):

    tempDoc = fitz.open(TEMPLATE + "blank.pdf")
    tempPage = tempDoc[0]

    if not tempPage._isWrapped:
        tempPage._wrapContents()

    tempPage.insertImage(place, pixmap=pix)

    for j in descriptions:
        if j == "StartPage" or j == "TitleSize":
            continue

        point = getPoint(j)
        p = fitz.Point(point["x"], point["y"])
        f = point["f"]

        if j == "Description":
            f = getString(index, "TitleSize")

        text = getString(index, j)
        t = str(text)

        if text is not None:
            tempPage.insertText(p, t, fontsize=f)

    pPage = (1375, 984)

    tempPage.insertText(pPage, str(pageNo), fontsize=6.5)

    return tempDoc


def newPages(src, index, ref):

    sourceDoc = fitz.open(src)
    newDoc = fitz.open(src)

    pageStart = getString(index, "StartPage")
    oldLen = int(len(sourceDoc))
    newLen = int(len(sourceDoc) + len(sourceDoc) - pageStart)

    selectPages = list(range(0, pageStart - 1)) + list(range(oldLen, newLen + 1))

    for i in range(pageStart - 1, len(sourceDoc)):

        sourcePage = sourceDoc[i]

        zoom_x = 2.0  # horizontal zoom
        zomm_y = 2.0  # vertical zoom
        mat = fitz.Matrix(zoom_x, zomm_y)  # zoom factor 2 in each dimension
        pix = sourcePage.getPixmap(
            matrix=mat
        )  # use 'mat' instead of the identity matrix

        pageNo = i - pageStart + 2

        tempDoc2 = getPage(index, pix, pageNo)

        newDoc.insertPDF(tempDoc2)

    newDoc.select(selectPages)
    newDoc.save(ref + ".pdf", garbage=3)


for i in files:
    if ".pdf" and "VDRL" in i:
        pos = i.find(" ")
        ref = i[5:pos]

        if ref in numbers:
            print(ref)
            index = numbers.index(ref) + 2
            newPages(i, index, ref)
        else:
            print(i + " not found in array")
