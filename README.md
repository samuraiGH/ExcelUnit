# ExcelUnit
Минимальная программа:
```
## uses ExcelUnit;//подключаем модуль
var app:= new ExcelApp('path');//создаём новый экземпляр приложения с нужной книгой
Print(app.CellOne[1, 1].Val);//выводим значение ячейки А1 на экран
app[1, 2].Val:= 4;//задаём значение ячейки В1
app.Save;//сохраняем изменения и закрываем книгу
app.Close;//закрываем Excel
```
Установка:
1. Скачать [ExcelUnit](https://github.com/samuraiGH/ExcelUnit/releases/download/v1.0/ExcelUnit.pas)
2. Поместить файл в папку с вашей программой

Ограничения:
- Должен быть установлен Excel
