# DiplomPRJ
import pandas as pd
import openpyxl
import datetime
import warnings


def import_to_main_base_csv(path_file):
    """функция импортирует из файла csv данные и записывает из в excel файл"""
    data_csv = pd.read_csv(path_file, index_col=None, header=None)
    writer = pd.ExcelWriter('base_main.xlsx', mode="a", engine='openpyxl', if_sheet_exists='overlay')
    data_csv.to_excel(writer, 'marks', index=False, header=False)
    writer.save()
    writer.close()


def import_to_main_base_excel(path_file):
    """функция импортирует из файла excel данные и записывает из в excel файл"""
    data_json = pd.read_excel(path_file, index_col=None, header=None)
    writer = pd.ExcelWriter('base_main.xlsx', mode="a", engine='openpyxl', if_sheet_exists='overlay')
    data_json.to_excel(writer, 'marks', index=False, header=False)
    writer.save()
    writer.close()


def enter_data_from_keyboard(list):
    """Функция ввода данных с клаыиатуры"""
    df = pd.DataFrame(list, index=[0])
    df.head()
    excel_data_df = pd.read_excel('base_main.xlsx', index_col=0, sheet_name='marks')
    df_unit = pd.concat([excel_data_df, df], ignore_index=True)
    print(df_unit)
    writer = pd.ExcelWriter('base_main.xlsx', mode="a", engine='openpyxl', if_sheet_exists='overlay')
    df_unit.to_excel(writer, 'marks')
    writer.save()
    writer.close()


def caunter_age(year_naw):
    """Отчет формирования возраста по всей базе с указанием года формирования отчета"""
    excel_data_df = pd.read_excel('base_main.xlsx', index_col=0, sheet_name='marks')
    excel_data_df.insert(10, "Year_Calc", year_naw)
    excel_data_df["Age"] = (excel_data_df["Year_Calc"] - excel_data_df["birt_year"])
    excel_data_df_live = excel_data_df.loc[excel_data_df['live'] == "Да"]
    excel_data_df_death = excel_data_df.loc[excel_data_df['live'] == "Нет"]
    df_unit = pd.concat([excel_data_df_live, excel_data_df_death], ignore_index=True)
    pd.set_option('display.max_rows', None)
    pd.set_option('display.max_columns', None)
    pd.set_option('display.max_colwidth', None)
    return df_unit


def caunter_age_live(year_naw):
    """Отчет формирования возраста по живым"""
    excel_data_df = pd.read_excel('base_main.xlsx', index_col=0, sheet_name='marks')
    excel_data_df.insert(10, "Year_Calc", year_naw)
    excel_data_df["Age"] = (excel_data_df["Year_Calc"] - excel_data_df["birt_year"])
    excel_data_df_live = excel_data_df.loc[excel_data_df['live'] == "Да"]
    excel_data_df_live = excel_data_df_live.drop(columns=['death_year', 'death_month', 'death_day'], axis=1)
    pd.set_option('display.max_rows', None)
    pd.set_option('display.max_columns', None)
    pd.set_option('display.max_colwidth', None)
    return excel_data_df_live


def caunter_age_death(year_naw):
    """Отчет формирования по возраста по мертвым"""
    excel_data_df = pd.read_excel('base_main.xlsx', index_col=0, sheet_name='marks')
    excel_data_df.insert(10, "Year_Calc", year_naw)
    excel_data_df["Age"] = (excel_data_df["Year_Calc"] - excel_data_df["birt_year"])
    excel_data_df_death = excel_data_df.loc[excel_data_df['live'] == "Нет"]
    pd.set_option('display.max_rows', None)
    pd.set_option('display.max_columns', None)
    pd.set_option('display.max_colwidth', None)
    return excel_data_df_death


def seacher_text(txt):
    """Отчет поиска по ключивым словам по столбцам фамилия, имя, отчество"""
    excel_data_df = pd.read_excel('base_main.xlsx', index_col=0, sheet_name='marks')
    search_surname_df = excel_data_df[excel_data_df['surname'].str.contains(txt)]
    search_name_df = excel_data_df[excel_data_df['name'].str.contains(txt)]
    search_patronymic_df = excel_data_df[excel_data_df['patronymic'].str.contains(txt)]
    df_unit1 = pd.concat([search_surname_df, search_name_df], ignore_index=True)
    df_unit2 = (pd.concat([df_unit1, search_patronymic_df], ignore_index=True)).drop_duplicates()
    pd.set_option('display.max_rows', None)
    pd.set_option('display.max_columns', None)
    pd.set_option('display.max_colwidth', None)
    return df_unit2


def caunter_age_live_auto(year_naw1):
    """Отчет формирования возраста автоматически по текущему году"""
    excel_data_df = pd.read_excel('base_main.xlsx', index_col=0, sheet_name='marks')
    excel_data_df.insert(10, "Year_Calc", year_naw1)
    excel_data_df["Age"] = (excel_data_df["Year_Calc"] - excel_data_df["birt_year"])
    excel_data_df_live = excel_data_df.loc[excel_data_df['live'] == "Да"]
    pd.set_option('display.max_rows', None)
    pd.set_option('display.max_columns', None)
    pd.set_option('display.max_colwidth', None)
    return excel_data_df_live


def caunter_adult(year_naw):
    """Отчет по совершенолетним """
    excel_data_df = pd.read_excel('base_main.xlsx', index_col=0, sheet_name='marks')
    excel_data_df.insert(10, "Year_Calc", year_naw)
    excel_data_df["Age"] = (excel_data_df["Year_Calc"] - excel_data_df["birt_year"])
    excel_data_df_live = excel_data_df.loc[excel_data_df['live'] == "Да"]
    excel_data_df_live = excel_data_df_live.loc[excel_data_df_live['Age'] > 18]
    excel_data_df_live = excel_data_df_live.drop(columns=['death_year', 'death_month', 'death_day'], axis=1)
    pd.set_option('display.max_rows', None)
    pd.set_option('display.max_columns', None)
    pd.set_option('display.max_colwidth', None)
    return excel_data_df_live


def caunter_age_live_pensioners(year_naw):
    """Отчет в разрезе живой/возраст больше 60 те пенсионеры"""
    excel_data_df = pd.read_excel('base_main.xlsx', index_col=0, sheet_name='marks')
    excel_data_df.insert(10, "Year_Calc", year_naw)
    excel_data_df["Age"] = (excel_data_df["Year_Calc"] - excel_data_df["birt_year"])
    excel_data_df_live = excel_data_df.loc[excel_data_df['live'] == "Да"]
    excel_data_df_live = excel_data_df_live.loc[excel_data_df_live['Age'] > 60]
    excel_data_df_live = excel_data_df_live.drop(columns=['death_year', 'death_month', 'death_day'], axis=1)
    pd.set_option('display.max_rows', None)
    pd.set_option('display.max_columns', None)
    pd.set_option('display.max_colwidth', None)
    return excel_data_df_live


def seacher_relatives(year_naw):
    """Отчет поиска потенциальных близких родственников в разрезе фамилия и отчество"""
    excel_data_df = pd.read_excel('base_main.xlsx', index_col=0, sheet_name='marks')
    excel_data_df.insert(10, "Year_Calc", year_naw)
    excel_data_df["Age"] = (excel_data_df["Year_Calc"] - excel_data_df["birt_year"])
    excel_data_df_live = excel_data_df.loc[excel_data_df['live'] == "Да"]
    excel_data_df_death = excel_data_df.loc[excel_data_df['live'] == "Нет"]
    df_unit = pd.concat([excel_data_df_live, excel_data_df_death], ignore_index=True)
    df_unit = df_unit[df_unit.duplicated(['surname', 'patronymic'])]
    pd.set_option('display.max_rows', None)
    pd.set_option('display.max_columns', None)
    pd.set_option('display.max_colwidth', None)
    return df_unit

"""Старт программы"""
warnings.simplefilter(action='ignore', category=FutureWarning)
"""Убирает информационное сообщение о не потдерже ряда функционалов в будующих версиях"""
try:
    """Перехват ошибки или проверка на наличие существующего файла"""
    file = open('base_main.xlsx')
except IOError as e:
    """Создание нового файла в случае отсутсвия """
    df = pd.DataFrame(columns=(
        ['surname', 'name', 'patronymic', 'birt_year', 'birth_month', 'birth_day', 'live', 'death_year',
         'death_month',
         'death_day', 'gender']))
    writer = pd.ExcelWriter('base_main.xlsx', mode="w", engine='openpyxl')
    df.to_excel(writer, 'marks')
    writer.save()
    writer.close()
else:
    """Перезапись существующего файла"""
    with file:
        df = pd.DataFrame(columns=(
            ['surname', 'name', 'patronymic', 'birt_year', 'birth_month', 'birth_day', 'live', 'death_day',
             'death_month',
             'death_year', 'gender']))
        writer = pd.ExcelWriter('base_main.xlsx', mode="a", engine='openpyxl', if_sheet_exists='overlay')
        df.to_excel(writer, 'marks')
        writer.save()
        writer.close()
while True:
    """Основной цикл программы"""
    base_ask = int(input(
        "Вас приветствует программа обработки персональных данных\n Для ввода данных нажмите 1 \n Для обработки персональных данных нажмите 2\n"))
    if base_ask == 1:
        choice = int(input(
            "Приветствую Вас в программе обработки персональных данных. Для продолжения выберите пункт из меню \n Ввод или загрузка данных нажмите 1 \n Ввод вручную нажмите 2\n"))
        if choice == 1:
            """выполнение загрузки данных вызов функци вып загрузку данных"""
            type_import_file = int(
                input("Введите тип файла загрузжаеиого в базу \n 1 файл csv \n 2 файл excel\n"))
            if type_import_file == 1:
                """Загрузка файла csv"""
                path_file_scv = input("Введите путь к данному файлу \n")
                import_to_main_base_csv(path_file_scv)
            if type_import_file == 2:
                """загрузка файла excel"""
                path_file_excel = input("Введите путь к данному файлу \n")
                import_to_main_base_excel(path_file_excel)
        if choice == 2:
            """Ввод данных с клавиатуры"""
            choice_var = int(
                input(
                    "выбиретет вариант ввода данных \n 1- ввод данных в одну строку \n 2-ввод данных последовательно\n"))
            if choice_var == 2:
                """Ввод данных последовательно"""
                live_inp = input("Человек на данный момент жив? \n Веедите Да или Нет\n")
                if live_inp == "Нет":
                    while True:
                        """Валидация введенных данных а именно года"""
                        death_year_inp = input("Введите год смерти \n")
                        len_inp_death_year = len(death_year_inp)
                        test_isdigit_death_year = death_year_inp.isdigit()
                        if not test_isdigit_death_year:
                            print("вы ввели не число, повторите ввод")
                            continue
                        death_year_inp = int(death_year_inp)
                        if len_inp_death_year > 4:
                            print(("вы ввели некорректный год, цифр должно быть 4"))
                            continue
                        else:
                            break
                    while True:
                        """Валидация введнных данных месяца"""
                        death_month_inp = input("Введите месяц кончины \n")
                        len_death_month_inp = len(death_month_inp)
                        test_isdigit_death_year = death_month_inp.isdigit()
                        if not test_isdigit_death_year:
                            print("вы ввели не число, повторите ввод")
                            continue
                        death_month_inp = int(death_month_inp)
                        if death_month_inp > 12:
                            print(("вы ввели некорректный месяц, в году месяцев 12"))
                            continue
                        else:
                            break
                    while True:
                        """Валидация введенных данных дня"""
                        death_day_inp = input("Введите день в какой приставился\n")
                        len_death_day_inp = len(death_day_inp)
                        test_isdigit_death_day_inp = death_day_inp.isdigit()
                        if not test_isdigit_death_day_inp:
                            print("вы ввели не число, повторите ввод")
                            continue
                        death_day_inp_int = int(death_day_inp)
                        if death_day_inp_int > 31:
                            print(("вы ввели некорректный день, в месяце максимум 31 день"))
                            continue
                        else:
                            break
                surname_inp = input(("Введите фамилию \n").title())
                name_inp = input(("Введите имя \n").title())
                patronymic_inp = input(("Введите отчество \n").title())
                gender_inp = input("Введите пол: М- мужской, Ж-женский")
                while True:
                    """Валидация данных год"""
                    birt_year_inp = input("Введите год рождения \n")
                    len_birt_year_inp = len(birt_year_inp)
                    test_birt_year_inp = birt_year_inp.isdigit()
                    if not test_birt_year_inp:
                        print("вы ввели не число, повторите ввод")
                        continue
                    birt_year_inp = int(birt_year_inp)
                    if len_birt_year_inp > 4:
                        print(("вы ввели некорректный год, цифр должно быть 4"))
                        continue
                    else:
                        break
                while True:
                    """Валидация данных месяц"""
                    birth_month_inp = input("Введите месяц рождения \n")
                    test_isdigit_birth_month_inp = birth_month_inp.isdigit()
                    if not test_isdigit_birth_month_inp:
                        print("вы ввели не число, повторите ввод")
                        continue
                    birth_month_inp = int(birth_month_inp)
                    if birth_month_inp > 12:
                        print(("вы ввели некорректный месяц, в году месяцев 12"))
                        continue
                    else:
                        break
                while True:
                    """Валидация данных день"""
                    birth_day_inp = input("Введите день рождения \n")
                    test_isdigit_birth_day_inp = birth_day_inp.isdigit()
                    if not test_isdigit_birth_day_inp:
                        print("вы ввели не число, повторите ввод")
                        continue
                    birth_day_inp_int = int(birth_day_inp)
                    if birth_day_inp_int > 31:
                        print(("вы ввели некорректный день, в месяце максимум 31 день"))
                        continue
                    else:
                        break
                if live_inp == "Да":
                    death_year_inp = None
                    death_month_inp = None
                    death_day_inp = None
            if choice_var == 1:
                """Ввод данных в одну строку"""
                live_inp = input("Человек на данный момент жив? \n Веедите Да или Нет\n")
                if live_inp == "Да":
                    var_birthday_input = input(
                            'Выберете даты рождения  в следующем формате: \n  дд.мм.год  или \n дд мм год или \n дд/мм/год\n дд-мм-год')
                    birth_day_inp = int(var_birthday_input[0:2])
                    birth_month_inp = int(var_birthday_input[3:5])
                    birt_year_inp = int(var_birthday_input[6:10])
                    death_year_inp = None
                    death_month_inp = None
                    death_day_inp = None
                    surname_inp = input(("Введите фамилию \n").title())
                    name_inp = input(("Введите имя \n").title())
                    patronymic_inp = input(("Введите отчество \n").title())
                    gender_inp = input("Введите пол: М- мужской, Ж-женский\n")
                if live_inp == "Нет":
                    var_birthday_input = input(
                        "Выберете даты рождения  в следующем формате: \n  дд.мм.год  или \n дд мм год или \n дд/мм/год\n дд-мм-год\n")
                    birth_day_inp = int(var_birthday_input[0:2])
                    birth_month_inp = int(var_birthday_input[3:5])
                    birt_year_inp = int(var_birthday_input[6:10])
                    var_death_input1 = input(
                        "Выберете дату смерти  в следующем формате: \n  дд.мм.год  или \n дд мм год или \n дд/мм/год\n дд-мм-год\n")
                    death_year_inp = int(var_death_input1[0:2])
                    death_month_inp = int(var_death_input1[3:5])
                    death_day_inp = int(var_death_input1[6:10])
                    surname_inp = input(("Введите фамилию \n").title())
                    name_inp = input(("Введите имя \n").title())
                    patronymic_inp = input(("Введите отчество \n").title())
                    gender_inp = input("Введите пол: М- мужской, Ж-женский\n")
            else:
                break
            n = (dict(surname=surname_inp, name=name_inp, patronymic=patronymic_inp, birt_year=birt_year_inp,
                      birth_month=birth_month_inp, birth_day=birth_day_inp, live=live_inp, death_year=death_year_inp,
                      death_month=death_month_inp, death_day=death_day_inp, gender=gender_inp))
            enter_data_from_keyboard(n)
    if base_ask == 2:
        choice_rep = int(input(
            "Приветствую Вас в системе обработки данных и их вывода \n 1- вывести возраст людей внесенных в базу\n 2- вывести данне живущих людей \n 3 -вывести данные мертвых \n 4- поиск и вывод данных по всей базе \n 5 -посчитать возраст на данный год без ввода \n 6-вывести всех совершенолетних \n 7- вывести всех пенсионного возарста\n 8- поиск потенциальных близких родственников \n"))
        if choice_rep == 1:
            yaer_rep = int(input("Введите год на который нужно узнать возраст всех людей в базе\n"))
            rep1 = caunter_age(yaer_rep)
            print(rep1)
        if choice_rep == 2:
            yaer_rep = int(input("Введите год на который нужно узнать возраст живых людей\n"))
            rep1 = caunter_age_live(yaer_rep)
            print(rep1)
        if choice_rep == 3:
            yaer_rep = int(input("Введите год на который нужно узнать возраст не живых людей\n"))
            rep1 = caunter_age_death(yaer_rep)
            print(rep1)
        if choice_rep == 4:
            txt_rep = input("введите слово или часть слова для поиска по базе \n")
            rep1 = seacher_text(txt_rep)
            print(rep1)
        if choice_rep == 5:
            year_naw1 = (datetime.datetime.now()).year
            rep1 = caunter_age_live_auto(year_naw1)
            print(rep1)
        if choice_rep == 6:
            year_naw1 = (datetime.datetime.now()).year
            rep1 = caunter_adult(year_naw1)
            print(rep1)
        if choice_rep == 7:
            year_naw1 = (datetime.datetime.now()).year
            rep1 = caunter_age_live_pensioners(year_naw1)
            print(rep1)
        if choice_rep == 8:
            year_naw1 = (datetime.datetime.now()).year
            rep1 = seacher_relatives(year_naw1)
            print(rep1)

    ask = input(
        'Выбиретет вариант сценария\n наберите Выход-  для выхода из программы \nИ нажмите любую клавишу для продолжения программы \n')
    if ask == "Выход":
        break
    else:
        continue
"""Конец программы"""