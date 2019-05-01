import zipfile
import re
import os

directory = '.'
files = os.listdir(directory)
inputfile_pattern = '.*\.(xlsx|xls)$'


def zip_zip(file):
    try:
        zip_file = re.sub(r'\.(xlsx|xls)$', '.zip', file)
        zippedFile = zipfile.ZipFile(zip_file, 'w')
        zippedFile.write(file)
        print("Архив {} создан".format(zip_file))
        zippedFile.close()
    except:
        print("Ошибка создания архива!")


for file in files:
    if re.match(inputfile_pattern, file):
        zip_zip(file)
print(files)



