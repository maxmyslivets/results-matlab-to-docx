"""Автозаполнение .docx файла данными,
полученными с помощью matlab.exe"""


import scipy.io as sio
from docxtpl import DocxTemplate, Listing
import os
import time

main_path = os.path.dirname(os.path.abspath(__file__))

student_name = input('Ф.И.О. студента: ')

print("[ INFO ]\tFile creation: lab1Func.m")

# Запись скрипта в MATLAB-скрипт .m
with open(main_path+'/lab1Func.m', 'w') as func:
    func.write(
        """
function lab1Func()

'    n - вариант'
n = {}

'    Получаем A'
A = randn(100, 4);
A = A(n*10-9:n*10,:)

'    Получаем l'
N = A' * A;
l = randn(100,1);
l = l(n*10-9:n*10)

b = A'*l

'    Определяем определитель матриц системы'
det(N)

'    Определяем ранг системы'
rank(N)

'    Проверяем на совместимость систему на основе теоремы Кронекера-Капелли'
Nz = [N b];
rank(Nz)

'    Автоматическое разложение матрцы системы уравнений 4 видов'

[L U] = lu(N)	% LU-разложение
L*U-N;

L1 = chol(N)	% Разложение Холецкого
L1'*L1-N;

[V S W] = svd(N)	% Сингулярное разложение

[Q R] = qr(N)	% QR-разложение
Q * R - N;

x4 = (Q' * b) / R(4,4)

'    Контроль (полный х)'
x = inv(N)*b

N11 = N(1:2,1:2);
N12 = N(1:2,3:4);
N21 = N(3:4, 1:2);
N22 = N(3:4,3:4);
b1 = b(1:2);
b2 = b(3:4);

Z = zeros(2,2);
E = eye(2,2);

'    к верхней треугольной матрице'
FL = [E Z
-N21*inv(N11) E];

Np = FL*N
bp = FL*b

'    Np*x2=bp => x2 = inv(Np)*bp'
x2 = inv(Np(3:4,3:4))*bp(3:4)


'    к нижней треугольной матрице'
FU = [E -N12*inv(N22)
Z E];


Np2 = FU*N
bp2 = FU*b
'    Np2*x1=bp2 => x1 = inv(Np2)*bp2'
x1 = inv(Np2(1:2,1:2))*bp2(1:2)

'    Верх Гаусcа GL'

GL = [inv(N11) Z
-N21*inv(N11) E];

Np3 = GL*N
bp3 = GL*b

'    Np3*xp=bp3 => x2 = inv(Np3)*bp3'
x3 = inv(Np3(3:4,3:4))*bp3(3:4)

'    Низ Гаусса GU'

GU = [E -N12*inv(N22)
Z inv(N22)];

Np4 = GU*N
bp4 = GU*b

'    Np4*x4=bp4 => x4 = inv(Np4)*bp4'
x4 = inv(Np4(1:2,1:2))*bp4(1:2)

'    Блочное для Гаусса GL GU'

% теперь Np3 вместо N

Np3_11 = Np3(1:2,1:2);
Np3_12 = Np3(1:2,3:4);
Np3_21 = Np3(3:4, 1:2);
Np3_22 = Np3(3:4,3:4);
bp3_1 = bp3(1:2);
bp3_2 = bp3(3:4);

GU2 = [E -Np3_12*inv(Np3_22)
Z inv(Np3_22)];

Np5 = GU2*Np3
bp5 = GU2*bp3

x5 = inv(Np5(3:4,3:4))*bp5(3:4)
x6 = inv(Np5(1:2,1:2))*bp5(1:2)

'    Псевдовеса'
A1 = A(:,1:2);
A2 = A(:,3:4);

N11 = A1'*A1;
N22 = A2'*A2;

P1 = eye(10)-A1*inv(N11)*A1';
% A2'*P1*A2 *x2 = A2'*P1*l
x7 = inv(A2'*P1*A2)*(A2'*P1*l)

P2 = eye(10)-A2*inv(N22)*A2';
% A1'*P2*A1 *x1 = A1'*P2*l
x8 = inv(A1'*P2*A1)*(A1'*P2*l)

save('var_lab1')
""".format(input('Номер варианта: '))
    )

# Проверка на наличие старого файла var_lab1.mat в дирректории
# Если есть, то удаляем 
if os.path.exists(main_path+'/var_lab1.mat'):
    os.remove(main_path+'/var_lab1.mat')

# Создание .bat-скрипта для запуска .m-скрипта в MATLAB
print("[ INFO ]\tFile creation: script.bat")
with open(main_path+'/script.bat', 'w') as bat:
    bat.write("matlab.exe -nodesktop -minimize -r run('"+main_path+"/lab1Func.m');quit")

# запуск .bat-скрипта
print("[ INFO ]\tStart lab1Func.m in MATLAB")
os.startfile(main_path+'/script.bat')

# Пока var_lab1.mat не создан, программа простаивает
while not os.path.exists(main_path+'/var_lab1.mat'):
    pass

print("[ INFO ]\tFile created: var_lab1.mat")

# Чтение файла .mat с помощью библиотеки scipy
test = sio.loadmat(main_path+'/var_lab1.mat')

print('[ INFO ]\tTransforming data from a var_lab1.mat to variable')
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

print('[ INFO ]\tTemplate: lab1.docx')
doc = DocxTemplate(main_path+"/lab1.docx")    # os.path.relpath("lab1.docx", start=None) - путь к файлу относительно стартового пути

print('[ INFO ]\tWrite the result data to a file: lab1.docx')
doc.render(context)

print("[ INFO ]\tSaved file: lab1_var-{}.docx".format(context['n_var']))
doc.save(main_path+"/lab1_var-{}.docx".format(context['n_var']))

print('[ INFO ]\tDeleted files')
os.remove(main_path+'/var_lab1.mat')
os.remove(main_path+'/script.bat')
os.remove(main_path+'/lab1Func.m')

input('Press for quit...')