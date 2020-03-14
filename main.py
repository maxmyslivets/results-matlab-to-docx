"""Запуск MATLAB-скрипта с помощью Python и автоматизация записи результатов вычислений в docx-файл"""


import scipy.io as sio
from docxtpl import DocxTemplate, Listing
import os
import time

# Вычисление начальной дирректории
main_path = os.path.dirname(os.path.abspath(__file__))

student_name = input('Ф.И.О. студента: ')

print("[ INFO ]\tFile creation: MatlabFunc.m")

# Чтение скрипта из matlab-script.txt
with open('matlab-script.txt', 'r', encoding="utf-8") as ms:
    mat_s = ms.read()

# Запись скрипта в MATLAB-скрипт .m
with open(main_path+'/MatlabFunc.m', 'w') as func:
    func.write(mat_s.format(input('Номер варианта: ')))

# Проверка на наличие старого файла results.mat в дирректории
# Если есть, то удаляем 
if os.path.exists(main_path+'/results.mat'):
    os.remove(main_path+'/results.mat')

# Создание .bat-скрипта для запуска .m-скрипта в MATLAB
print("[ INFO ]\tFile creation: script.bat")
with open(main_path+'/script.bat', 'w') as bat:
    bat.write("matlab.exe -nodesktop -minimize -r run('"+main_path+"/MatlabFunc.m');quit")

# запуск .bat-скрипта
print("[ INFO ]\tStart MatlabFunc.m in MATLAB")
os.startfile(main_path+'/script.bat')

# Пока results.mat не создан, программа простаивает
while not os.path.exists(main_path+'/results.mat'):
    pass

print("[ INFO ]\tFile created: results.mat")

# Чтение файла .mat с помощью библиотеки scipy
print(main_path+'/results.mat')
test = sio.loadmat(main_path+'/results.mat')

print('[ INFO ]\tTransforming data from a results.mat to variable')
context = {}
for key in test:
    if key in ['__header__', '__version__', '__globals__', 'n', 'ans']:
        continue
    context[key] = ''
    for rows in test[key]:
        context[key] += str(rows) + '\n'

# Экранируем текст в переменной
print('[ INFO ]\tSetting listing for .docx')
for key in context:
    context[key] = Listing(context[key])

# Добавляем имя студента и номер варианта в переменную
context['n_var'] = test['n'][0][0]
context['student_name'] = student_name

print('[ INFO ]\tTemplate: Template.docx')
doc = DocxTemplate(main_path+"/Template.docx")    # os.path.relpath("Template.docx", start=None) - путь к файлу относительно стартового пути

print('[ INFO ]\tWrite the result data to a file: Template.docx')
doc.render(context)

print("[ INFO ]\tSaved file: Result-var-{}.docx".format(context['n_var']))
doc.save(main_path+"/Result-var-{}.docx".format(context['n_var']))

print('[ INFO ]\tDeleted files')
os.remove(main_path+'/results.mat')
os.remove(main_path+'/script.bat')
os.remove(main_path+'/MatlabFunc.m')

input('Press for quit...')