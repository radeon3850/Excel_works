import glob
from pathlib import Path
import pandas as pd
import time


while True:

    try:
        print(" ")
        path_qwery = input("Укажите папку для объединения Excel файлов: ")+("/").replace('/', "\\")
        print(" ")
        path_save_file = input("Скопируйте путь для сохранения файла: ")+("/").replace('/', "\\")
        print(" ")
        file_name = input("Укажите название итогового файла объединёных данных: ")
        new_name_file = ''.join(file_name.split())
        print(" ")
        print("Идёт процес объединения данных файлов Excel.....")

        list_failes_for_query = [item for item in glob.glob(f"{path_qwery}*{'.xlsx'}")]
        print("Список обьединяемых файлов")
        for elem in list_failes_for_query:
            print(elem)


        path = Path(path_qwery)
        min_excel_file_size = 100

        df = pd.concat([pd.read_excel(f)
                        for f in path.glob("*.xlsx")
                        if f.stat().st_size >= min_excel_file_size], ignore_index=True)

        df.to_excel(path_save_file+new_name_file +'.xlsx')


        print(" ")
        print("Программа завершила объединение файлов")
        print(" ")
        print(f"Перейдите в директорию сохранения файла: {file_name}.xlsx")

    except Exception as ex:
        print(ex)
        print("Не верно введены данные для чтения или сохранения файлов")

    time.sleep(3)

    print(" ")
    exit_program = input("Продолжить работу програмы? Yes/No(нажмите y/n + ENTER: ").lower()
    if exit_program =="n":
        break

