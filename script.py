from bs4 import BeautifulSoup
import pandas as pd

# Указываем путь к файлу, содержащему исходный код страницы
file_path = 'html.txt'  # Замените на путь к вашему файлу

# Чтение содержимого файла
with open(file_path, 'r', encoding='utf-8') as file:
    html_content = file.read()

# Парсим HTML-код
soup = BeautifulSoup(html_content, 'html.parser')

# Ищем таблицу с классом 'user-grade'
table = soup.find('table', {'class': 'user-grade'})

if table:
    print("Таблица с классом 'user-grade' найдена!")

    # Заголовки таблицы: название курса, оценка, диапазон оценок и комментарий
    headers = ['course_name', 'grade', 'range', 'feedback']
    
    # Список для хранения данных таблицы
    data = []
    
    # Считываем строки таблицы (начиная со второй строки, так как первая — это заголовки)
    rows = table.find_all('tr')

    for row in rows[1:]:  # Пропускаем заголовок
        # Извлекаем данные по каждому столбцу
        course_name_tag = row.find('a', {'class': 'gradeitemheader'})
        grade_tag = row.find('td', {'class': 'column-grade'})
        range_tag = row.find('td', {'class': 'column-range'})
        feedback_tag = row.find('td', {'class': 'column-feedback'})
        
        # Проверяем, есть ли все необходимые теги, и извлекаем текст
        course_name = course_name_tag.text.strip() if course_name_tag else ''
        grade = grade_tag.text.strip() if grade_tag else ''
        range_value = range_tag.text.strip() if range_tag else ''
        feedback = feedback_tag.text.strip() if feedback_tag else ''
        
        # Добавляем собранные данные в список
        data.append([course_name, grade, range_value, feedback])

    # Если данные собраны, сохраняем их в Excel
    if data:
        # Создаём DataFrame
        df = pd.DataFrame(data, columns=headers)

        # Сохраняем DataFrame в Excel
        output_file = 'moodle_grades.xlsx'
        df.to_excel(output_file, index=False)  # Убираем аргумент encoding

        print(f"Таблица сохранена в файл {output_file}")
    else:
        print("Данные таблицы не были найдены.")
else:
    print("Таблица с классом 'user-grade' не найдена.")
