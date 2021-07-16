import fitz
import os

os.chdir('E:/Desktop')

arr = os.listdir()

with open('directory.txt', 'w') as f:
    for item in arr:
        pos = item.find('.pdf')
        fileName = item[0:pos]
        f.write("%s\n" % fileName)


