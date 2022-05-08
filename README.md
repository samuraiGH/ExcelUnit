P.S. В связи с [недавними блокировками](https://habr.com/ru/news/t/661113/) GiHub'ом русский репозиториев, дальнейшие изменения будут размещаться на [GitFlic](https://habr.com/ru/news/t/661113/)

# ExcelUnit

## Описание

`ExcelUnit` - модуль для работы с Excel в PascalABC.NET

## Примеры

```pas
## uses ExcelUnit; // Подключаем модуль

var app := new ExcelApp('path'); // Создаём новый экземпляр приложения с нужной книгой
Print(app.CellOne[1, 1].Val); // Выводим значение ячейки А1 на экран
app[1, 2].Val := 4; // Задаём значение ячейки В1
app.Save; // Сохраняем изменения и закрываем книгу
app.Close; // Закрываем Excel
```

## Установка

1. [Скачать](https://github.com/samuraiGH/ExcelUnit/releases/download/v1.0/ExcelUnit.pas) `ExcelUnit`
2. Поместить скачанный .pas файл в папку с вашей программой

## Ограничения

Должен быть установлен Excel
