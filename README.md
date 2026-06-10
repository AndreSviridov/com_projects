# com_projects
Репозиторий создан для хранения коммерческих проектов.

| № | Проект | Описание | Стек |
|---|--------|----------|------|
| 1 | [Comments Tone Classifier](https://github.com/AndreSviridov/com_projects/blob/main/Comments_tone_classifier/Comments_tone_classifier.ipynb) | Анализ тональности комментариев на базе модели Rubert. Загрузка файла, определение негативных отзывов, визуализация статистики, выгрузка результатов. | `pandas` `rubert` `matplotlib` `time` |
| 2 | [Geo Cluster](https://github.com/AndreSviridov/com_projects/blob/main/Geo_cluster/geo.ipynb) | Геоаналитика кофейных заведений Санкт-Петербурга. Интерактивная карта с кластеризацией точек. | `pandas` `folium` `MarkerCluster` |
| 3 | [Ипотека-обработка](https://github.com/AndreSviridov/com_projects/blob/main/ipoteka/ipoteka_obrabotka.py) | Парсинг банковских выгрузок (txt) в структурированный Excel-отчёт. Автоматический поиск файлов графиков в сетевой папке (>500 директорий, >1000 файлов). | `pandas` `openpyxl` `xlrd` `re` `logging` `customtkinter` |
| 4 | [Ипотека-разнесение](https://github.com/AndreSviridov/com_projects/blob/main/ipoteka/ipoteka_raznesenie.py) | Распределение данных по 4 Excel-реестрам с сохранением истории изменений. Умное суммирование (формулы + комментарии), бекапы, автономный .exe. | `pandas` `openpyxl` `shutil` `threading` `customtkinter` `PyInstaller` |
