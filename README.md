Скрипт выводит путь до файла .doc .docx и количество страниц в нём, а так же сумму страниц всех найденных документов в каталогах

Для выполнения скрипта:
1) Переместить файл в корневую папку
2) ПКМ->"Выполнить с помощью PowerShell"

Или
1) в строке `$pages = Get-ChildItem -Path "<путь к корневой папке>" -Recurse -Include *.doc, *.docx |` прописать путь к корневой папке
2) ПКМ->"Выполнить с помощью PowerShell"

Возможные ошибки:
1) Невозможно выполнить сценарий в указанной системе.  
   
   Необходимо запустить PowerShell от имени администратора и ввести следующую команду `Set-ExecutionPolicy Unrestricted -Scope CurrentUser`

2) Ошибка кодировки.  
   
   Открываем скрипт при помощи NotePad++ выбираем вкладку "Кодировка" и меняем кодировку на "Кодировка UTF-8 с BOM"
