import requests
import os

# Список ссылок на PDF-файлы
pdf_links = [
"https://link1.pdf",
"https://link2.pdf",
...,
    # Добавьте остальные ссылки
]

# Список наименований, в том порядке, в котором они соответствуют PDF-файлам
pdf_names = [
"name1",
"name2",
    # Добавьте остальные наименования
]

# Заголовок User-Agent
headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/116.0.0.0 Safari/537.36'
}

# Создание папки для сохранения файлов, если её нет
if not os.path.exists(r"C:\Users\xxx_pdf"):
    os.makedirs(r"C:\Users\xxx_pdf")

# Скачивание и сохранение файлов
for link, name in zip(pdf_links, pdf_names):
    response = requests.get(link, headers=headers, verify=False)
    file_path = os.path.join(r"C:\Users\xxx_pdf", f"{name}.pdf")

    # Проверка, существует ли уже файл с таким именем, и создание копии при необходимости
    counter = 1
    while os.path.exists(file_path):
        file_path = os.path.join(r"C:\Users\xxx_pdf", f"{name}_{counter}.pdf")
        counter += 1

    with open(file_path, "wb") as f:
        f.write(response.content)
