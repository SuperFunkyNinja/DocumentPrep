import os
import shutil

os.chdir('E:/Desktop')

arr = os.listdir()

for i in arr:
    if i.startswith('VDRL') and '.pdf' in i:
        src = i
        pos = i.find(' ')
        dst = i[5:pos]+str('.pdf')
        print(dst)
        shutil.copy(src,dst)
