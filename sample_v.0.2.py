import pandas as pd
from tkinter import filedialog
from tkinter import *
from tkinter import ttk
from tkinter import messagebox
import time
import requests
import tkinter as tk

window = Tk()

notebook = ttk.Notebook(window)

tab2 = ttk.Frame(notebook)
tab3 = ttk.Frame(notebook)

notebook.add(tab2, text='EXCEL SELECTED')
notebook.add(tab3, text='ADD IP INFO')

notebook.pack(expand=True, fill='both')

def open_file():
    global input_file_path
    input_file_path = filedialog.askopenfilename()
    print(f"File:{input_file_path}")

def save_file():
    global output_file_path 
    output_file_path = filedialog.asksaveasfilename()
    print(f"File saved:{output_file_path}")

def find_and_save_matches():
    global input_file_path
    if input_file_path is None:
        print("Не выбран входной файл.")
        show_misstake_4()
        return
    
    search_column_index_value = search_column_index_1.get()
    search_value_value = search_value.get()

    if not search_column_index_value.isdigit():
        print("Индекс столбца должен быть числом.")
        show_misstake_5()
        return

    search_column_index_value = int(search_column_index_value)

    find_and_save_matches_internal(input_file_path, search_column_index_value, search_value_value, output_file_path)

def show_success_message():
    messagebox.showinfo("Успешно", "Операция завершена успешно!")
def show_misstake_1():
    messagebox.showinfo("Ошибка","DataFrame пуст. Нет данных для обработки.")
def show_misstake_2():
    messagebox.showinfo("Ошибка","Столбец с введенным индексом не существует в DataFrame.")
def show_misstake_3():
    messagebox.showinfo("Ошибка","Нет совпадений для записи в новый файл.")
def show_misstake_4():
    messagebox.showinfo("Ошибка","Не выбран входной файл.")
def show_misstake_5():
    messagebox.showinfo("Ошибка","Индекс столбца должен быть числом.")

def find_and_save_matches_internal(input_file_path, search_column_index, search_value, output_file_path):
    df = pd.read_excel(input_file_path)
    if df.empty:
        print("DataFrame пуст. Нет данных для обработки.")
        show_misstake_1()
        return
    if search_column_index >= len(df.columns):
        print(f"Столбец с индексом {search_column_index} не существует в DataFrame.")
        show_misstake_2()
        return
    
    matches = df.iloc[:, search_column_index].astype(str).str.contains(search_value, case=False, na=False)
    print(matches)

    if not matches.any():
        print("Нет совпадений для записи в новый файл.")
        show_misstake_3()
        return
    
    matching_rows = df[matches]

    print("Найдены совпадения:")
    print(matching_rows)

    matching_rows.to_excel(output_file_path, index=False)
    print(f"Результат записан в файл: {output_file_path}")
    show_success_message()


def find_and_save_ip():

    global input_file_path, output_file_path
    if input_file_path is None:
        print("Не выбран входной файл.")
        show_misstake_4()
        return
    
    search_column_index_ip = search_column_index_11.get()

    print(search_column_index_ip)

    if not search_column_index_ip.isdigit():
        print("Индекс столбца должен быть числом.")
        show_misstake_5()
        return

    search_column_index_ip = int(search_column_index_ip)

    process_excel_file(input_file_path, output_file_path, search_column_index_ip)

def get_ip_info(ip_address):
    try:
        response = requests.get(f'http://ip-api.com/json/{ip_address}')
        response.raise_for_status()
        data = response.json()
        return {
            'IP_Address': ip_address,
            'Страна': data.get('country', ''),
            'Город': data.get('city', ''),
            'Информация о провайдере': data.get('isp', '')
        }
    except requests.RequestException as e:
        print(f"Произошла ошибка при выполнении запроса для IP-адреса {ip_address}: {e}")
        return None

def process_excel_file(input_file_path, output_file_path, search_column_index_ip):
    df = pd.read_excel(input_file_path)
    print(1)
    print(search_column_index_ip)
    print(2)

    for index, row in df.iterrows():
        ip_address = row.iloc[search_column_index_ip].replace("dstip=", "")
        print(ip_address)
        ip_info = get_ip_info(ip_address)
        if ip_info:
            df.at[index, 'IP_Address'] = ip_info['IP_Address']
            df.at[index, 'Страна'] = ip_info['Страна']
            df.at[index, 'Город'] = ip_info['Город']
            df.at[index, 'Информация о провайдере'] = ip_info['Информация о провайдере']
        time.sleep(1)  
    df.to_excel(output_file_path, index=False)
    show_success_message()
    

# def csv_helper():
#     global input_file_path,output_file_path
#     if input_file_path is None:
#         print("Не выбран входной файл.")
#         show_misstake_4()
#         return
    
#     search_razdelitel=search_razdelitel_1.get()
#     print(search_razdelitel)
#     print('------------------------------------')
#     print(input_file_path)
#     print('------------------------------------')
#     print(output_file_path)

#     convert_csv_to_excel(input_file_path,output_file_path,search_razdelitel)

# def convert_csv_to_excel(input_file_path, output_file_path,search_razdelitel):
#     encodings_to_try = ['cp1251']

#     for encoding in encodings_to_try:
#         try:
#             df=pd.read_csv(input_file_path, sep=search_razdelitel,encoding=encoding, engine="python")
#         except Exception as e:
#             print(f"Mistake:{e}")
#             show_misstake_1()
#         if df is not None:
#             df.to_excel(output_file_path, index=False,engine='openpyxl')
#             show_success_message()
#         else:
#             print("huy")
    


# #Разделение CSV файла
# greeting_csv = ttk.Label(tab1,
#     text='Укажите разделитель',
#     width=60,
# )
# search_razdelitel_1=ttk.Entry(tab1)
# button_csv_1 = ttk.Button(tab1,
#     text="Open file",
#     command=open_file,
#     width=60
# )
# button_csv_2 = ttk.Button(tab1,
#     text="Save file",
#     command=save_file,
#     width=60
# )
# button_csv_3 = ttk.Button(tab1,
#     text="Start",
#     command=csv_helper,
#     width=60
# )


#Поиск строк по значению
search_column_index_ip = ttk.Entry(tab3)
search_column_index_1 = ttk.Entry(tab2)
search_value = ttk.Entry(tab2)
greeting_3 = ttk.Label(tab2,
    text='Укажите индекс ячейки',
    width=60,
)
greeting_4 = ttk.Label(tab2,
    text='Укажите триггер',
    width=60,
)
button_1 = ttk.Button(tab2,
    text="Open file",
    command=open_file,
    width=60
)
button_2 = ttk.Button(tab2,
    text="Save file",
    command=save_file,
    width=60
)
button_3 = ttk.Button(tab2,
    text="Start",
    command=find_and_save_matches,
    width=60
)


#Добавление информации по IP
search_column_index_11 = ttk.Entry(tab3)
greeting_ip = ttk.Label(tab3,
    text='Укажите индекс ячейки',
    width=60,
)
button_ip_1 = ttk.Button(tab3,
    text="Open file",
    command=open_file,
    width=60
)
button_ip_2 = ttk.Button(tab3,
    text="Save file",
    command=save_file,
    width=60
)
button_ip_3 = ttk.Button(tab3,
    text="Start",
    command=find_and_save_ip,
    width=60
)

greeting_ip.pack(fill=tk.BOTH, expand=True)
search_column_index_11.pack(fill=tk.BOTH, expand=True)
button_ip_1.pack(fill=tk.BOTH, expand=True, padx=6, pady=6)
button_ip_2.pack(fill=tk.BOTH, expand=True, padx=6, pady=6)
button_ip_3.pack(fill=tk.BOTH, expand=True, padx=6, pady=6)


greeting_3.pack(fill=tk.BOTH, expand=True)
search_column_index_1.pack(fill=tk.BOTH, expand=True)
greeting_4.pack(fill=tk.BOTH, expand=True)
search_value.pack(fill=tk.BOTH, expand=True)
button_1.pack(fill=tk.BOTH, expand=True, padx=6, pady=6)
button_2.pack(fill=tk.BOTH, expand=True, padx=6, pady=6)
button_3.pack(fill=tk.BOTH, expand=True, padx=6, pady=6)

window.mainloop()