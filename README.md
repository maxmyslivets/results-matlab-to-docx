# results-matlab-to-docx
Запуск MATLAB-скрипта с помощью Python и автоматизация записи результатов вычислений в docx-файл

### Основные используемые библиотеки:
docxtpl - для изменения .docx файла по его шаблону
scipy - для чтения данных из .mat файла (MATLAB)

### Структура программы:
1. Чтение .txt файла, который можно изменять, вручную записывая в него скрипт MATLAB-функции
2. Перезапись этого скрипта в .m файл с некоторыми изменениями переменных (исх. данных)
3. Создание .bat файла для запуска MATLAB-скрипта через matlab.exe (предполагается, что matlab установлен в системе)
4. Запуск .m файла и ожидание до создания matlab-ом .mat файла
5. После появления .mat файла в каталоге, его чтение с помощью scipy.io
6. Перестройка считанных с .mat файла данных в тип str()
7. Редактирование .docx шаблона и его запись в новый .docx файл
8. Удаление временных файлов из каталога.