Набор утилит для расчета НКУ (низковольтных комплетных устройств)
=
Эти утилиты будут полезны при расчете электрошкафов.

# 1 Описание функций
## 1.1	Заполнение паспортов
Данная функция предназначена для автоматизированного заполнения паспортов к НКУ выпускаемых ООО «____». Заполнятся паспорта выделенных строк, столбцы могут быть произвольными.
Выбор шаблона напольный/навесной идет исходя из высоты шкафа, прописанного в столбце N Журнала учета НКУ, по умолчанию ниже 1500 мм выбирается шаблон навесных шкафов, выше и включительно 1500 мм выбирается шаблон напольных корпусов. Возможно, настроить границу выбора шаблонов в файле AppSettings.json см. пункт 3.
Возможные проблемы: 
+ Некорректно заполнен столбец «А». В этом случае макрос не сможет создать папку для сохранения. Рекомендуется после появления проблемы проверить через Диспетчер задач наличие процесса Word, если процесс открыт, его необходимо завершить.
+ Некорректно заполняется паспорт. Рекомендуется проверить корректность заполнения полей в Журнале учета НКУ Рекомендуется проверить корректность и наличие шаблона в папке «Template» (см. пункт 1.22)
Ограничения: запускается только в журнале НКУ текущего года (настройка журнала текущего года см. пункт 3). 
## 1.2	Удалить формулы
Данная функция удаляет формулы на выделенном диапазоне и оставляет значения. После окончания работы функции выделение устанавливается на ячейку A1. Графического интерфейса функция не имеет. Ограничений на атрибуты Excel файла нет.
Графического интерфейса функция не имеет.
## 1.3	Удалить все формулы
Данная функция удаляет формулы и оставляет значения в диапазоне В2:G500 на всех листах кроме первого. Предполагается использовать данную функцию для удаления формул при расчете НКУ, для уменьшения размера файла.  Графического интерфейса функция не имеет. Ограничений на атрибуты Excel файла нет.
Графического интерфейса функция не имеет.
## 1.4	Корпуса щитов
Заполняет атрибуты корпусов НКУ (ширина, высота, глубина, IP и т.д.) в Журнале учета НКУ при занесении записей в журнал. Данные извлекаются из базы данных (БД) по столбцу Z журнала (артикул корпуса щита). Если артикул отсутствует в БД, ячейка в Журнале учета НКУ подсвечивается оранжевым.
Ограничения: запускается только в журнале НКУ текущего года (настройка журнала текущего года см. пункт 3.1), с проверкой по имени файла.
Графического интерфейса функция не имеет. Есть информационные окна.
## 1.5	Корпуса в базу
Добавляет запись в БД если он там отсутствует. Работает только с одной активной строкой, даже если выделено более.
Для добавления необходимо заполнить поля в Журнале учета НКУ:
+ Степень защиты, IP – только числовое значение;
+ Климатическое исполнение – текстовое значение;
+ Высота – числовое значение, возможно текстовое для шкафов IT;
+ Ширина – только числовое значение;
+ Глубина – только числовое значение;
+ Масса – текстовое значение, практически не используется, рекомендуется ставить символ ‘-’ прочерк.
+ Тип шкафа – текстовое значение (Металл/Пластик)
Не рекомендуется записывать в базу напольные шкафы, состоящие из нескольких панелей.
Возможные проблемы: 
+ Текстовое значение в столбце «ширина», это вызовет ошибку, запись не будет добавлена.
Ограничения: запускается только в журнале НКУ текущего года (настройка журнала текущего года см. пункт 3) и успешно выполняет задачу если в БД нет записи о корпусе (проверка проходит по артикулу). Работает только с одной активной строкой, даже если выделено более.
Графического интерфейса функция не имеет. Есть информационные окна.
## 1.6	Исправить запись в БД
Внесение изменений в БД корпусов НКУ, если есть ошибочные данные по характеристикам шкафов.
Для внесения изменений необходимо чтобы были заполнены поля в Журнале учета НКУ:
+ Степень защиты, IP – только числовое значение;
+ Климатическое исполнение – текстовое значение;
+ Высота – числовое значение, возможно текстовое для шкафов IT;
+ Ширина – только числовое значение;
+ Глубина – только числовое значение;
+ Масса – текстовое значение, практически не используется, рекомендуется ставить символ ‘-’ прочерк.
+ Тип шкафа – текстовое значение (Металл/Пластик)
Возможные проблемы: 
+ Текстовое значение в столбце «ширина», это вызовет ошибку, запись не будет изменена.
Ограничения: запускается только в журнале НКУ текущего года (настройка журнала текущего года см. пункт 3.1). 
Графического интерфейса функция не имеет. Есть информационные окна.
## 1.7	Разметка листов
Функция предназначена для генерации шаблона расчета, второй и далее страницы, где поэлементно расписана комплектация шкафов НКУ. При генерации шаблона, автоматически присваивается номер листа, второй лист книги получает номер 1 (первый шкаф), третий лист получает номер 2 (второй шкаф) и т.д. Если номер уже был присвоен на предыдущих листах, то присвоение номера листа игнорируется.  Шаблон устанавливает ширину ячеек Excel и все границы в диапазоне A1:I11.
Графического интерфейса функция не имеет.
## 1.8	Границы
Размечает границы в выделенном диапазоне, является аналогом встроенной функции «Все границы» на вкладке Главная.
## 1.9	Шрифт
Устанавливает шрифт, определенный в файле AppSettings.json выделенном диапазоне. По умолчанию использует шрифт “Calibri” 11пт. Применяется для выравнивания шрифта и установки на листе Excel.
Графического интерфейса функция не имеет, есть информационные окна.
## 1.10	Расчет
Функция предназначена для генерации списка рассчитываемых НКУ в качестве первой страницы.
Если есть данные в области генерации на листе Excel A1-H9, то функция не будет работать и выдаст предупреждение о наличии данных в этой области.
Графического интерфейса функция не имеет.
## 1.11	Причесать расчет
Функция добавляет в расчет столбец «Кратность» и приводит шрифт к установленному в AppSettings.json. Функция изменяет второй и последующие листы. Функция необходима для совместимости со старыми расчетами.
Графического интерфейса функция не имеет.
## 1.12 Объединить ячейки
Функция объединяет ячейки с данными в вертикальном столбце и вставлет значение этих ячеек в вверхнюю с разделителем - один пробел.
## 1.13	Поиск в Яндексе
Функция предназначена для поиска содержимого ячейки в поисковой системе Яндекс. Открывается браузер, который установлен в системе по умолчанию. При нажатии на значок функции автоматически открывается браузер с поисковой выдачей значения ячейки. Функция не работает с объединенными ячейками или с диапазоном ячеек. Если невозможно преобразовать значение в поисковый запрос, то нажатие на значок функции игнорируется.
## 1.14	Поиск в Google
Функция предназначена для поиска содержимого ячейки в поисковой системе Google. Открывается браузер, который установлен в системе по умолчанию. При нажатии на значок функции автоматически открывается браузер с поисковой выдачей значения ячейки. Функция не работает с объединенными ячейками или с диапазоном ячеек. Если невозможно преобразовать значение в поисковый запрос, то нажатие на значок функции игнорируется.
## 1.15	Формула ВПР IEK
Функция вставляет подготовленную формулу ВПР для поиска данных в связанной таблице (прайсе производителя IEK) в столбцы:
+ Столбец “B” Описание – описание артикула согласно прайс-листу вендора;
+ Столбец “D” Кратность – кратность товара (штуки или упаковки) согласно прайс-листу вендора;
+ Столбец “E” Пр-ль – вендор, в данном случае IEK;
+ Столбец “G” Цена – цена с НДС за кратность товара;
+ Столбец “F” Скидка – текущая скидка данного вендора;
+ Столбец “H” Цена со скидкой – вычисление стандартной скидки предоставляемой дистрибьютером по формуле =G(номер строки)*(100-F(номер строки))/100;
+ Столбец “I” Ст-ть – вычисление стоимости текущей позиции (количество, умноженное на цену со скидкой) по формуле =H(номер строки)*C(номер строки).
Для настройки данных для вставки предназначена функция «Настройка формул» (см. пункт 1.24 текущей документации).
Графического интерфейса функция не имеет.
## 1.16	Формула ВПР EKF
Функция вставляет подготовленную формулу ВПР для поиска данных в связанной таблице (прайсе производителя EKF) в столбцы:
+ Столбец “B” Описание – описание артикула согласно прайс-листу вендора;
+ Столбец “D” Кратность – кратность товара (штуки или упаковки) согласно прайс-листу вендора;
+ Столбец “E” Пр-ль – вендор, в данном случае EKF;
+ Столбец “G” Цена – цена с НДС за кратность товара;
+ Столбец “F” Скидка – текущая скидка данного вендора;
+ Столбец “H” Цена со скидкой – вычисление стандартной скидки предоставляемой дистрибьютером по формуле =G(номер строки)*(100-F(номер строки))/100;
+ Столбец “I” Ст-ть – вычисление стоимости текущей позиции (количество, умноженное на цену со скидкой) по формуле =H(номер строки)*C(номер строки).
Для настройки данных для вставки предназначена функция «Настройка формул» (см. пункт 1.24 текущей документации).
Графического интерфейса функция не имеет.
## 1.17	Формула ВПР DKC
Функция вставляет подготовленную формулу ВПР для поиска данных в связанной таблице (прайсе производителя DKC) в столбцы:
+ Столбец “B” Описание – описание артикула согласно прайс-листу вендора;
+ Столбец “D” Кратность – кратность товара (штуки или упаковки) согласно прайс-листу вендора;
+ Столбец “E” Пр-ль – вендор, в данном случае DKC;
+ Столбец “G” Цена – цена с НДС за кратность товара;
+ Столбец “F” Скидка – текущая скидка данного вендора;
+ Столбец “H” Цена со скидкой – вычисление стандартной скидки предоставляемой дистрибьютером по формуле =G(номер строки)*(100-F(номер строки))/100;
+ Столбец “I” Ст-ть – вычисление стоимости текущей позиции (количество, умноженное на цену со скидкой) по формуле =H(номер строки)*C(номер строки).
Для настройки данных для вставки предназначена функция «Настройка формул» (см. пункт 1.24 текущей документации).
Графического интерфейса функция не имеет.
## 1.18	Формула ВПР KEAZ
Функция вставляет подготовленную формулу ВПР для поиска данных в связанной таблице (прайсе производителя KEAZ) в столбцы:
+ Столбец “B” Описание – описание артикула согласно прайс-листу вендора;
+ Столбец “D” Кратность – кратность товара (штуки или упаковки) согласно прайс-листу вендора;
+ Столбец “E” Пр-ль – вендор, в данном случае KEAZ;
+ Столбец “G” Цена – цена с НДС за кратность товара;
+ Столбец “F” Скидка – текущая скидка данного вендора;
+ Столбец “H” Цена со скидкой – вычисление стандартной скидки предоставляемой дистрибьютером по формуле =G(номер строки)*(100-F(номер строки))/100;
+ Столбец “I” Ст-ть – вычисление стоимости текущей позиции (количество, умноженное на цену со скидкой) по формуле =H(номер строки)*C(номер строки).
Для настройки данных для вставки предназначена функция «Настройка формул» (см. пункт 1.24 текущей документации).
Графического интерфейса функция не имеет.
## 1.19	Формула ВПР DEKraft
Функция вставляет подготовленную формулу ВПР для поиска данных в связанной таблице (прайсе производителя DEKraft) в столбцы:
+ Столбец “B” Описание – описание артикула согласно прайс-листу вендора;
+ Столбец “D” Кратность – кратность товара (штуки или упаковки) согласно прайс-листу вендора;
+ Столбец “E” Пр-ль – вендор, в данном случае DEKraft;
+ Столбец “G” Цена – цена с НДС за кратность товара;
+ Столбец “F” Скидка – текущая скидка данного вендора;
+ Столбец “H” Цена со скидкой – вычисление стандартной скидки предоставляемой дистрибьютером по формуле =G(номер строки)*(100-F(номер строки))/100;
+ Столбец “I” Ст-ть – вычисление стоимости текущей позиции (количество, умноженное на цену со скидкой) по формуле =H(номер строки)*C(номер строки).
Для настройки данных для вставки предназначена функция «Настройка формул» (см. пункт 1.24 текущей документации).
Графического интерфейса функция не имеет.
## 1.20	Формула ВПР Chint
Функция вставляет подготовленную формулу ВПР для поиска данных в связанной таблице (прайсе производителя Chint) в столбцы:
+ Столбец “B” Описание – описание артикула согласно прайс-листу вендора;
+ Столбец “D” Кратность – кратность товара (штуки или упаковки) согласно прайс-листу вендора;
+ Столбец “E” Пр-ль – вендор, в данном случае Chint;
+ Столбец “G” Цена – цена с НДС за кратность товара;
+ Столбец “F” Скидка – текущая скидка данного вендора;
+ Столбец “H” Цена со скидкой – вычисление стандартной скидки предоставляемой дистрибьютером по формуле =G(номер строки)*(100-F(номер строки))/100;
+ Столбец “I” Ст-ть – вычисление стоимости текущей позиции (количество, умноженное на цену со скидкой) по формуле =H(номер строки)*C(номер строки).
Для настройки данных для вставки предназначена функция «Настройка формул» (см. пункт 1.24 текущей документации).
Графического интерфейса функция не имеет.
## 1.21	Модульные аппараты
Функция предназначена для подбора артикулов модульных автоматических выключателей и выключателей нагрузки с дальнейшим выводом результата на лист Excel.
Особенности использования функции:
+ Цветовое поле сигнализирует о существовании модульного аппарата с данными характеристиками. Если поле зеленое, то этот аппарат существует, если красное, то либо он не выбран чекбоксом либо он не существует;
+ Нажатие на цветовое поле модульных автоматических выключателей переносит выбранные характеристики (кроме тока) на нижележащие строки, на выключателях нагрузки данная функциональность не реализована;
+ При закрытии окна функции сохраняется настройки характеристик из первой строки и при повторном вызове этой функции они распространяются на остальные строки;
+ Сохранение характеристик аппаратов происходит до закрытия программы Excel, при повторном запуске Excel характеристики сбрасываются на первоначальные.
Проблема данной функции в том, что база данных модульных аппаратов может не отражать актуального состояния и снятых с производства артикулов.
## 1.22	Трансформаторы тока
Функция предназначена для подбора трансформаторов тока по параметрам. При нажатии на кнопку «В буфер» в буфер обмена Windows копируется артикул соответствующего вендора. А при нажатии кнопки «На лист» на лист Excel в текущую позицию копируется артикул трансформатора тока и вставляется формула ВПР согласно пунктам 1.15-1.20 настоящей документации.
## 1.23	Рубильники TwinBlock
Функция предназначена для подбора рубильников TwinBlock и аксессуаров к ним. При нажатии кнопки «На лист» на лист Excel в текущую позицию копируется артикул рубильника и аксессуаров если они были выбраны, после вставляются формулы ВПР согласно пунктам 1.15-1.20 настоящей документации.
Если к рубильнику нет совместимых аксессуаров, то они автоматически исключаются из результата работы программы.
Рисунок 6. Окно выбора рубильников TwinBlock: 
## 1.24	Настройка формул
Данная утилита предназначена для ввода и сохранения формул ВПР в макросы. 
Работа с данной утилитой происходит следующим образом:
1.	В открытом листе Excel наводим на строку, где прописана формула ВПР, это может быть любой расчет или заготовка для формулы ВПР;
2.	Переводим фокус на окно Settings;
3.	Нажимаем кнопку Cчитать! на необходимом вендоре;
4.	Далее нажимаем кнопку Сохранить, происходит процесс сохранения в память компьютера и выводится дата сохранения.
## 1.25	About
Окно «О программе». Несет информационную цель. 
## 1.26	Открыть папку
Функция открывает папку с расположением файлов программ.

Настройка программы, файл AppSettings.json
Файл находится в папке */Config/AppSettings.json и имеет следующую структуру:
{
  "Resources": {
    "NameFileJournal": "_Журнал учета НКУ_2022.xlsx",
    "HeightMaxBox": 1500,
    "TemplateWall": "Паспорт_навесные.docx",
    "TemplateFloor": "Паспорт_напольные.docx"
  },
  "CorrectFontResources": {
    "NameFont": "Calibri",
    "SizeFont": 11
  },
  "FormSettings": {
    "FormTopMost": true
  },
  "GlobalDateBaseLocation": "//192.168.100.100/ftp/Info_A/FTP/Производство Абиэлт/Инженеры/База данных/1.7.5/"
}
Значения полей:
+ NameFileJournal – Полное имя Журнала учета НКУ в котором предполагается работа;
+ HeightMaxBox – высота шкафов при которой программа заполнения паспортов изделий автоматически их считает напольными и выбирает соответствующий шаблон;
+ TemplateWall – имя файла шаблона паспорта изделия навесного исполнения в папке файлов программы, с расположением */Template/;
+ TemplateFloor – имя файла шаблона паспорта изделия напольного исполнения в папке файлов программы, с расположением */Template/;
+ NameFont – название используемого шрифта для функций «Разметка листов», «Причесать расчет» и «Шрифт».
+ SizeFont – размер используемого шрифта для функций «Разметка листов», «Причесать расчет» и «Шрифт».
+ FormTopMost – вывод форм «Модульные аппараты», «Трансформаторы тока» и «Рубильники TwinBlock» всегда по верх листа Excel, если значение true и обычное поведение окон если значение false.
+ GlobalDateBaseLocation – указывает на расположение общей базы данных на FTP сервере. Если данный путь не доступен программе, то будет использоваться локальная база данных.
Данный файл редактируется блокнотом и изменения вступают в силу после перезапуска Excel.
 
# 2	Небезопасные функции (Not Safe)
Использование функции: «Удалить формулы», «Удалить все формулы», «Корпуса щитов», «Разметка листов», «Причесать расчет», «Шрифт», «Формула ВПР IEK», «Формула ВПР EKF», «Формула ВПР DKC», «Формула ВПР KEAZ», «Формула ВПР DEK», «Модульные аппараты», «Трансформатор тока» и «Рубильники TwinBlock», небезопасно по отношению к имеющимся данным на листе. При выполнении этих функций на текущий лист Excel записываются новые данные по верх имеющихся, при этом стандартная функция Excel ‘Отменить ввод’ не работает и восстановить прежние данные невозможно.
До использования данных функций рекомендуется сделать сохранение текущего документа, и закрыть без сохранения этот документ если в процессе работы функции были повреждены нужные данные. 
 
# 4	Удаление надстройки ExcelMacro Add-in

## Как корректно удалить надстройку
Для удаления надстройки требуется выполнить следующие действия:
1. Закрыть все файлы Microsoft Office Excel и выйти из программы
2. Открыть меню «Пуск», перейти в «Панель управления» и найти пункт меню, отвечающий за изменение/удаление программ.
3. Найти в списке установленных программ надстройку ExcelMacro Add-in
4. Удалить надстройку
## Что делать в случае неправильного удаления
Иногда надстройка удаляется не корректно, в этом случае повторная ее установка будет невозможна. Для
решения данной проблемы необходимо:
+ Убедиться, что надстройка удалена
+ Открыть редактор реестра (regedit.exe)
+ Перейти по следующему пути \HKEY_CURRENT_USER\Software\Microsoft\
+ Найти раздел «VSTO»
+ Внутри раздела «VSTO» будет 2 подраздела «Security» и «SolutionMetadata»
+ Перейти в раздел «Security», а затем в «Inclusion», внутри которого будут разделы с названием вида «45e59586-ebfa-4f61-9b1d-d83ef1b5fca0»
+ Удалить каждый такой раздел, внутри которого в ключе «Url» содержится строка «ExcelMacro.vsto»
+ Перейти в раздел «SolutionMetadata», внутри которого будут разделы с названием вида «45e59586-ebfa-4f61-9b1d-d83ef1b5fca0»
+ Удалить каждый такой раздел, внутри которого в ключе «addInName» содержится строка «ExcelMacro»
После выполнения всех пунктов установить надстройку.
Если проблема не была решена, попробуйте произвести следующие действия:
+ Перейти в каталог C:\Users\UserName\AppData\Local\Apps\2.0
+ Удалить все папки, внутри которых содержится информация, относящаяся к
надстройке ExcelMacro Add-In
+ Произвести повторную чистку реестра
+ Установить надстройку
Если после этого при установке надстройки возникает ошибка System.Runtime.InteropServices.COMException (0x800736B3), то следует провести следующие действия: 
+ Открыть блокнотом файл ExcelMacroAdd.vsto, найти значение publicKeyToken=”xxxxxxx” и скопировать в буфер обмена значение токена.
+ Открыть редактор реестра и через меню «Правка-Найти» найти и удалить все разделы с данным токеном.
