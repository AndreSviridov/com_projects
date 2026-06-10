# com_projects
Репозиторий создан для хранения коммерческих проектов.

| #    | Наименование проекта                | Описание                                                     |
| ---- | ----------------------------------------------------- | ------------------------------------------------------------ |
| 1.   | [Comments_tone_classifier](https://github.com/AndreSviridov/com_projects/blob/main/Comments_tone_classifier/Comments_tone_classifier.ipynb) | Программа для определения тональности комментариев по загруженному файлу на базе модели Rubert c отображением статистики и последующей выгрузкой файла с негативными комментариями. Примененные библиотеки и команды: pandas, rubert, time, matplotlib, pd.read_excel, astype(), tolist(), list comprehension, groupby, count(), reset_index(), round(), pivot_table, for, iterrows(), pie, plot, to_excel().
| 2.   | [Geo_cluster](https://github.com/AndreSviridov/com_projects/blob/main/Geo_cluster/geo.ipynb) | Создание folium-карты для геоаналитики кофейных заведений СПБ: pandas, folium, CustomIcon, read_excel, срезы, def, map, MarkerCluster, marker.
| 3.   | [ipoteka_obrabotka](https://github.com/AndreSviridov/com_projects/blob/main/ipoteka/ipoteka_obrabotka.py) | Загрузка, парсинг и обогащение данных из txt в структурированный Excel‑отчёт с автоматическим поиском файлов действующих графиков в сетевых папках.
- Парсинг txt (cp1251, разделитель ';') и xls/xlsx реестров через pandas + openpyxl.
- Обход сетевых папок (Path.rglob) с фильтрацией файлов по ключевым словам (более 500 директорий c более 1000 файлов на сетевом драйве)
- Извлечение транша через regex: поиск паттерна (\d+)\s*транш с числовой сортировкой
- Парсинг из нужных строк в xls (xlrd по заданной логике)
- Финальное автоматическое форматирование Excel: настройка ширины колонок, цветовая маркировка шапки, создание гиперссылок, подсветка ячеек, отличающихся от моды по месяцу платежа.
- Логирование через QueueHandler в отдельный поток — неблокирующий GUI
| 4.   | [ipoteka_raznesenie](https://github.com/AndreSviridov/com_projects/blob/main/ipoteka/ipoteka_raznesenie.py) | Распределение обработанных данных по 4 целевым реестрам с сохранением полного следа изменений.
- Умное суммирование: преобразование чисел в формулы =старое+новое с хранением полной истории в cell.comment;
- Накопление дат через запятую с нормализацией форматов (datetime, YYYY-MM-DD, DD.MM.YYYY → ДД-ММ-ГГГГ);
- Автоматическая простановка статуса «должник» при обнаружении минуса ('-' in str(value)) с блокировкой дальнейшего разнесения;
- Механизм резервного копирования: shutil.copy2 всех 5 файлов перед записью, восстановление из любой точки через выбор папки;
- Многопоточность (threading) — интерфейс не зависает во время обработки.
