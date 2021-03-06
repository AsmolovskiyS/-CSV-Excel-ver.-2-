# Загрузка больших CSV файлов в Excel (ver. 2)

![](https://raw.githubusercontent.com/AsmolovskiyS/-CSV-Excel-ver.-2-/master/Screenshot/Scrn_01.png)

## Описание
Программа написана на VBA.  
Программа предназначена для загрузки в Excel больших CSV файлов, с более чем 1048576 строк.   
Перед началом использования данной программы убедитесь в отсутствии пустых строк в начале и в конце вашего CSV файла!  
Первая строка CSV  повторяется как шапка на всех последующих листах. 
В зависимости от настроек данная программа может загружать файл в память, как частями, так и целиком (что позволяет работать с 32-х битными версиями Excel).  
Следует учитывать, что для 32-х битной версии Excel 2010 - доступно около 2 ГБ памяти, а для 32-х битной версии Excel 2013 и 2016 - доступно около 3 ГБ памяти.
Рекомендуется использовать 64-х битные версии Excel так как у них не должно быть ограничений по используемой памяти.


## Запуск программы
### Вариант 1
Скачать репозиторий, импортировать форму и модуль в Excel.  
Для запуска запустить процедуру "Start".  
### Вариант 2
Скачать готовый Excel файл: [https://drive.google.com/file/d/1cfMr1DP_GuHq6w6hTwHHo6tp068alSQi/view?usp=sharing](https://drive.google.com/file/d/1cfMr1DP_GuHq6w6hTwHHo6tp068alSQi/view?usp=sharing).  
Запустить его и нажать кнопку "Старт".   

## Контакты
Если вам понравилась программа или у Вас есть вопросы, можете писать сюда: Asmolovskiy.Sergey@gmail.com