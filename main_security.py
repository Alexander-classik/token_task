import pandas as pd
import mysql.connector
import PySimpleGUI as sg
import random
import aspose.words as aw
import ctypes, sys

# Подключение к БД
conn = mysql.connector.connect(user='root', password='root', host='127.0.0.1', database='token_task')
cursor = conn.cursor(buffered=True)

# Импрот
def import_token(result_out, name):

    doc = aw.Document()
    builder = aw.DocumentBuilder(doc)
    for i in range(0, len(result_out)):
        builder.writeln("Тип работы: " + result_out[i][0])
        builder.writeln("Тема работы: " + result_out[i][1])
        builder.writeln("Задание: " + result_out[i][2])
    doc.save(str(name)+'.docx')


def pars(file):
    # Чтение excel
    df = pd.read_excel(io=file, engine='openpyxl', sheet_name='Лист1')

    # Парс excel
    result = []

    # Парс ТЕМ
    for i in range(0, len(df['Темы'].tolist())):
        add_ser = 'INSERT INTO `themas` (`name`) VALUES (%s)'
        result.append(df['Темы'].tolist()[i])
        check_input = 'SELECT * FROM `themas` WHERE `name` = %s'
        cursor.execute(check_input, result)
        if cursor.fetchone() == None:
            cursor.execute(add_ser, result)
            conn.commit()
            result = []
        else:
            result = []

    # Парс ЗАДАНИЙ
    for i in range(0, len(df['Задание'].tolist())):
        add_ser = 'INSERT INTO `tasks` (`name`) VALUES (%s)'
        result.append(df['Задание'].tolist()[i])
        check_input = 'SELECT * FROM `tasks` WHERE `name` = %s'
        cursor.execute(check_input, result)
        if cursor.fetchone() == None:
            cursor.execute(add_ser, result)
            conn.commit()
            result = []
        else:
            result = []

    # Парс ТИПОВ ЗАДАНИЙ
    for i in range(0, len(df['Тип задания'].tolist())):
        add_ser = 'INSERT INTO `type_tasks` (`name`) VALUES (%s)'
        result.append(df['Тип задания'].tolist()[i])
        check_input = 'SELECT * FROM `type_tasks` WHERE `name` = %s'
        cursor.execute(check_input, result)
        if cursor.fetchone() == None:
            cursor.execute(add_ser, result)
            conn.commit()
            result = []
        else:
            result = []

    # Парс БИЛЕТОВ
    for i in range(0, len(df['Темы'].tolist())):
        add_ser = 'INSERT INTO `tokens` (`themas_id`, `tasks_id`, `type_tasks_id`) VALUES (%s, %s, %s)'
        for j in range(0, len(df.values.tolist()[i])):
            data_db = []
            if j == 0:
                data_db.append(df.values.tolist()[i][j])
                check_input = 'SELECT `id` FROM `themas` WHERE `name` = %s'
                cursor.execute(check_input, data_db)
                result.append(cursor.fetchone()[0])
            elif j == 1:
                data_db.append(df.values.tolist()[i][j])
                check_input = 'SELECT `id` FROM `tasks` WHERE `name` = %s'
                cursor.execute(check_input, data_db)
                result.append(cursor.fetchone()[0])
            elif j == 2:
                data_db.append(df.values.tolist()[i][j])
                check_input = 'SELECT `id` FROM `type_tasks` WHERE `name` = %s'
                cursor.execute(check_input, data_db)
                result.append(cursor.fetchone()[0])
        check_input = 'SELECT * FROM `tokens` WHERE `themas_id` = %s AND `tasks_id` = %s AND `type_tasks_id` = %s'
        cursor.execute(check_input, result)
        if cursor.fetchone() == None:
            cursor.execute(add_ser, result)
            conn.commit()
            result = []
        else:
            result = []
sg.theme('dark grey 9')


def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False


def pars_win():
    if is_admin():
    # Code of your program here
        text = sg.popup_get_file('Please enter a file name')
        pars(text)
    else:
        # Re-run the program with admin rights
        ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, __file__, None, 1)


def import_(result_out):
    layout = [[sg.Text('Введите название документа'), sg.InputText()],
              [sg.Button('Ок', key='ok')]]
    window = sg.Window('Выгрузка билета', layout)
    while True:
        event, values = window.read()
        if event == sg.WIN_CLOSED or event == 'Cancel':  # if user closes window or clicks cancel
            break
        if event == 'ok':
            import_token(result_out, values[0])


def main():
    cursor.execute('SELECT COUNT(`name`) FROM `themas`')
    N = cursor.fetchone()
    dropdown = []
    index = []
    for i in range(1, N[0]+1):
        sel_them = 'SELECT `name` FROM `themas` WHERE `id` = %s'
        index.append(i)
        cursor.execute(sel_them, index)
        dropdown.append(cursor.fetchone())
        index = []
    layout = [
        [sg.InputCombo(dropdown, key='thema'), sg.Button('Сгенерировать', key='gen')],
        [sg.Output(size=(88, 20))],
        [sg.Button('Загрузить', key='open_p'), sg.Button('Выгрузить', key='im')]
    ]
    window = sg.Window('Генератор билетов', layout)
    while True:
        event, values = window.read()
        if event in (None, 'Exit'):
            break
        if event == 'open_p':
            pars_win()
        if event == 'gen':
            cursor.execute('SELECT COUNT(`name`) FROM `themas`')
            N = cursor.fetchone()
            result_out = []
            for i in range(1, N[0] + 1):
                result = []
                val = []
                right = False
                sel_them = 'SELECT `id`, `name` FROM `themas` WHERE `id` = %s'
                index.append(i)
                cursor.execute(sel_them, index)
                val.append(cursor.fetchone())
                index = []
                write = 0
                if values['thema'][0] == str(val[0][1]):
                    sel_teor = "SELECT id FROM `type_tasks` WHERE `name` = 'теория'"
                    cursor.execute(sel_teor)
                    index.append(cursor.fetchone()[0])
                    sel_count = 'SELECT MAX(tasks_id) FROM `tokens` ' \
                                'WHERE type_tasks_id = %s '\
                                'AND themas_id = %s '
                    index.append(val[0][0])
                    result.append(index[0])
                    cursor.execute(sel_count, index)
                    tt_count = cursor.fetchone()[0]
                    t_count = []
                    t_count.append(index[0])
                    sel_count = 'SELECT COUNT(type_tasks_id) FROM `tokens`' \
                                'WHERE type_tasks_id = %s'
                    cursor.execute(sel_count, t_count)
                    control = 0
                    for j in range(0, cursor.fetchone()[0]):
                        task = random.randint(1, tt_count)
                        if write == 2:
                            break
                        for n in range(0, tt_count):
                            if task != control:
                                sel_token = 'SELECT ' \
                                            '(SELECT `name` FROM `type_tasks` ' \
                                            'WHERE type_tasks.id=%s), ' \
                                            '(SELECT `name` FROM `themas` ' \
                                            'WHERE themas.id=%s), ' \
                                            '(SELECT `name` FROM `tasks` ' \
                                            'WHERE tasks.id=%s) ' \
                                            'FROM `tokens`'
                                result.append(val[0][0])
                                result.append(task)
                                control = task
                                check_sel = 'SELECT * FROM `tokens` ' \
                                            'WHERE type_tasks_id = %s ' \
                                            'AND themas_id = %s ' \
                                            'AND tasks_id = %s'
                                cursor.execute(check_sel, result)
                                if cursor.fetchone() != None:
                                    cursor.execute(sel_token, result)
                                    result_out.append(cursor.fetchone())
                                    result = []
                                    result.append(index[0])
                                    write += 1
                                    break
                                else:
                                    result = []
                                    result.append(index[0])
                                    j -= 1
                                    break
                            else:
                                j -= 1
                                break
                if write == 2:
                    result = []
                    index = []
                    if values['thema'][0] == str(val[0][1]):
                        sel_teor = "SELECT id FROM `type_tasks` WHERE `name` = 'практика'"
                        cursor.execute(sel_teor)
                        index.append(cursor.fetchone()[0])
                        sel_count = 'SELECT MAX(tasks_id) FROM `tokens` ' \
                                    'WHERE type_tasks_id = %s ' \
                                    'AND themas_id = %s '
                        index.append(val[0][0])
                        result.append(index[0])
                        cursor.execute(sel_count, index)
                        tt_count = cursor.fetchone()[0]
                        p_count = []
                        p_count.append(index[0])
                        sel_count = 'SELECT COUNT(type_tasks_id) FROM `tokens` WHERE type_tasks_id=%s'
                        cursor.execute(sel_count, p_count)
                        for p in range(0, cursor.fetchone()[0]):
                            task = random.randint(0, tt_count)
                            for n in range(0, tt_count):
                                if right == False:
                                    sel_token = 'SELECT ' \
                                                '(SELECT `name` FROM `type_tasks` ' \
                                                'WHERE type_tasks.id=%s), ' \
                                                '(SELECT `name` FROM `themas` ' \
                                                'WHERE themas.id=%s), ' \
                                                '(SELECT `name` FROM `tasks` ' \
                                                'WHERE tasks.id=%s) ' \
                                                'FROM `tokens`'
                                    result.append(val[0][0])
                                    result.append(task)
                                    check_sel = 'SELECT * FROM `tokens` ' \
                                                'WHERE type_tasks_id = %s ' \
                                                'AND themas_id = %s ' \
                                                'AND tasks_id = %s'
                                    cursor.execute(check_sel, result)
                                    if cursor.fetchone() != None:
                                        cursor.execute(sel_token, result)
                                        result_out.append(cursor.fetchone())
                                        result = []
                                        result.append(index[0])
                                        right = True
                                        break
                                    else:
                                        result = []
                                        result.append(index[0])
                                        break
                                else:
                                    break
                if right == True:
                    break
            for m in range(0, len(result_out)):
                print("Тип работы: "+result_out[m][0])
                print("Тема работы: "+result_out[m][1])
                print("Задание: "+result_out[m][2])
                print('')
        if event == 'im':
            import_(result_out)
    window.close()


if __name__ == "__main__":
    main()
