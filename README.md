При выгрузке статистики выполнения работ сотрудников столкнулись со следующей проблемой. Выгружался общий файл на 75 человек (таблица excel с четырьмя листами), содержащий данные по выработке и зарплатам. Для того что-то бы ознакомить каждого сотрудника с данными, касающимися непосредственно его, приходилось вручную выделять их в отдельную таблицу. Для оптимизации этого процесса был написан этот скрипт. 

Логика работы:

- принимает на входе общий файл;
- собирает все уникальные значения в солбце "***user_name***";
- копирует данные по каждому "***user_name***" в отдельную таблицу, в название файла которой так же записывается "***user_name***";
- работает со всеми таблицами подходящего формата, независимо от количества листов в них.

--- 
When exporting the performance statistics of employees, we encountered the following problem. A general file for 75 people was being exported (an Excel table with four sheets), containing data on productivity and salaries. To familiarize each employee with the data relevant to them, it was necessary to manually highlight their information into a separate table. To optimize this process, this script was written.

Working logic:

- Takes the general file as input;
- Collects all unique values in the "***user_name***" column;
- Copies data for each "***user_name***" into a separate table, with the file name also including "***user_name***";
- Works with all tables of the appropriate format, regardless of the number of sheets in them.
