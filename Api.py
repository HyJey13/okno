import re
import requests
from tkinter import *
from docx import Document

# паттерны для разных валидаций
# r'^[\w\.-]+@[\w\.-]+\.\w+$' - почта
# r'^\+7\(\d{3}\)\d{3}-\d{2}-\d{2}$' - телефон для формата +7(999)123-45-67
# r'^(?=.*[A-Za-z])(?=.*\d)[A-Za-z\d]{8,}$' - пароль 8 символов 1 буква 1 цифра
# r'^https?://[\w\.-]+\.\w+$'  - url ссылка
# r'^\d{2}\.\d{2}\.\d{4}$'  - формат даты но не проверяет наличие даты в календаре
# r'^(\d{1,3}\.){3}\d{1,3}$' - ip адресс
# r'^([0-9A-Fa-f]{2}:){5}[0-9A-Fa-f]{2}$' - mac адрес
# r'^\d{4} \d{6}$' - паспорт

def get_check():
    url = "http://127.0.0.1:4444/TransferSimulator/fullName"
    a = requests.get(url).json()['value']

    lbl_fio.config(text=a)
    pattern = r'[а-аА-ЯёЁ] + [а - яА - ЯёЁ] + [а-яА-ЯёЁ]'
    check = re.fullmatch(pattern, a)

    if check != None:
        lbl_check.config(text="Валидные данные")
    else:
        lbl_check.config(text = 'ФИО содержит запрещенные символы')

def send_result():
    doc = Document("ПУ/ТестКейс.docx")
    table = doc.tables[0]
    row = table.rows().cells

    row[0].text =  lbl_fio.cget("text")
    row[1].text = lbl_check.cget("text")
    row[2].text = "Успешно"
    doc.save("ПУ/ТестКейс.docx")


win = Tk()
win.title("Валидация данных")
win.geometry("600x200")
win.resizable(False, False)

bth_get = Button(win, text = "Получить данные", width=25,font = "Arial, 12", command=get_check)
btn_test = Button(win, text= "Отправить результат тестов", width=25, font = "Arial, 12")
lbl_fio= Label(win, text="", font = "Arial, 12")
lbl_check = Label(win, text="", font = "Arial, 12")

bth_get.grid(row=0, column=0, padx=10, pady=10)
btn_test.grid(row=1, column=0, padx=10, pady=10)
lbl_fio.grid(row=0, column=1, padx=10, pady=10)
lbl_check.grid(row=1, column=1, padx=10, pady=10)

win.mainloop()