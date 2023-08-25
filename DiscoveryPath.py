from pdfminer.high_level import extract_text

path = "C:/Users/pc/Downloads/forABBmailsorted/"  + "Черный Сергей Николаевич 19.10.1990"

text = extract_text(path)

path = "C:/Users/pc/Downloads/forABBmailsorted/" + matches[0]

if os.path.exists(path):
    print('Перемещаю файл:', X)
    os.replace(path, path + "/" + X)
else:
    os.mkdir(path) # создаем папку
    os.replace(path, path + "/" + X)
    print('\n СОЗДАЮ ПАПКУ. ПУТЬ К ПАПКЕ: ')
    print(path2)
    print('Перемещаю файл:', X)

