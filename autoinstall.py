import os
import shutil
import pythonping
import datetime


def nowversion(arm, user, pasw, name_programm):
    """
    Узнаём версию программы установленной сейчас на удаленном компьютере

    """
    cmd = f'wmic /NODE:\"{arm}\" /USER:\"{user}\" /password: \"{pasw}\" product get name| findstr \"{name_programm}\"'
    output = os.popen(cmd, 'r')
    templist = []
    for row in output:
        row = row.encode('cp1251', 'replace').decode('cp866').strip()
        templist.append(row)
        return templist[0]


def if_arm_online(arm):
    try:
        response = pythonping.ping(arm, count=1)
        if response.success():
            print(f"{arm} доступен")
            return True
        else:
            print(f"{arm} не доступен")
            return False
    except:
        print(f"{arm} не доступен")
        return False


def logwritting(data):
    with open(f'{way_to}\\log.txt', 'a')as logfile:
        logfile.write(data + '\n')


def copying(arm, way_to):
    """
    Копируемнеобходимый дистрибутив на удаленный компьютер
    """
    print(f'Копирование файлов на {arm}')
    try:
        shutil.copytree(f'{way_to}\\soft', f'\\\\{arm}\\c$\\psexec')
        print('Копирование завршено')
    except FileExistsError:
        print('Папка уже существует, попытка удаления, затем копирования...')
        shutil.rmtree(f'\\\\{arm}\\c$\\psexec')
        shutil.copytree(f'{way_to}\\soft', f'\\\\{arm}\\c$\\psexec')
        print('Копирование завршено')


def taskkill(arm, user, pasw):
    """
    Прекращаем выполнение программы на удаленном компьютере
    """
    cmd = f'taskkill.exe /s {arm} /u {user} /p {pasw} /F /T /IM  DispatchTerminal.exe'
    output = os.popen(cmd)
    temp_list = []
    for row in output:
        row = row.encode('cp1251', 'replace').decode('cp866').strip()
        temp_list.append(row)
    print(f'На {arm}:')
    print(*temp_list)


def makedir_cfg(arm, way_to_copy_xml, user_sttings_inuse):
    """
    Создаём необходимые директории и копируем файл настроек
    """
    print('Создаём необходимые директории и копируем файл настроек')
    dirprotei = f'\\\\{arm}\\{way_to_copy_xml[:21]}'
    dirdispach = f'\\\\{arm}\\{way_to_copy_xml[:38]}'
    try:
        os.mkdir(dirprotei)
        print(f'Директория {dirprotei} создана')
    except OSError:
        print(f'Создать директорию {dirprotei} не удалось')

    try:
        os.mkdir(dirdispach)
        print(f'Директория {dirdispach} создана')
    except OSError:
        print(f'Создать директорию {dirdispach} не удалось')

    print(f'Копируем файл настроек в {dirdispach}')
    try:
        shutil.copyfile(user_sttings_inuse, f'\\\\{arm}\\{way_to_copy_xml}')
        print('Скопировано успешно')
    except OSError:
        print('Не удалось скопировать файл настроек')


def uninstallprogramm(arm, user, pasw, tmpprogramm):
    os.system(f'wmic /NODE:\"{arm}\" /USER:\"{user}\" /password: \"{pasw}\" product where description=\"{tmpprogramm}\" uninstall')
    print(f'{tmpprogramm} is uninstalled')


def installprogramm(way_to, arm, user, pasw, dt_version):
    print('installing new version of soft')
    os.system(
        f'{way_to}\\psexec.exe \\\\{arm} -u {user} -p {pasw} -h msiexec.exe /i \"C:\\psexec\\{dt_version}\"')



name_programm = 'Дежурн'  # поиск программы ведется по этому словосочетанию
DT_version = 'DT-7.21.1-release-Spb-38533.msi'  # фактическое название файла утсановки + визуально понятно какая версия

# динамичное определение папки, где всё хранится
way_to = os.path.abspath(__file__)
way_to = os.path.dirname(way_to)

way_to_copy_xml = 'c$\\ProgramData\\Protei\\DispatchTerminal\\UserSettings.xml'  # где лежит файл сеттингов на удаленной ммашине
oper112_xml = f'{way_to}\\usersetting\\oper112\\UserSettings.xml'
nachsmen_xml = f'{way_to}\\usersetting\\nachsmen\\UserSettings.xml'
zamnach_xml = f'{way_to}\\usersetting\\zamnach\\UserSettings.xml'
zamnach_cout_xml = f'{way_to}\\usersetting\\zamnach_cout\\UserSettings.xml'
user_sttings_inuse = zamnach_xml
is_change_roll = False  # ставим флаг на смену роли
first_install = False  # первая установка нужно копирование файла настройки
no_install = False  # просто меняем роль без установки
linebreake = '********************'

# подгатавливаем файл логов к записи событий данной сессии
with open(f'{way_to}\\log.txt', 'w+') as logfile:
    logfile.write(str(datetime.datetime.now()) + '\n')

# печатаем список компьютеров где будутпроводиться работы
with open(f'{way_to}\\list.txt') as list_of_arms:
    print(
        f'Внимание! Будет установлен {DT_version} '
        f'на список АРМ-ов ниже:'
    )
    for row_t in list_of_arms:
        print(row_t)

print(f'Путь к UserSettings: {user_sttings_inuse}')
if first_install:
    print('Первая установка, без смены роли')
    is_change_roll = False  # глупо менять роль при первой установке
elif no_install:
    print('Без установки, только смена роли')
else:
    print('Старое СПО будет удалено!')
    print('Процесс Dispatch.exe будет остановлен!')
print(f'Смена роли: {is_change_roll}')
question_yn = input('Do you want to continue? Y/N ')
if question_yn.lower() == 'y':
    domen = input('Введите домен ')
    user = f'{domen}\\' + input('Введите логин ')
    pasw = input('Введите пароль ')

    # читаем названия армов из файла list.txt
    with open(f'{way_to}\\list.txt') as list_of_arms:
        for arm in list_of_arms:
            arm = arm.strip()
            logwritting(linebreake)
            logwritting(arm)

            # отделить строчкой утсановку разных армов
            print(linebreake)

            # определяем, доступен ли АРМ
            if not if_arm_online(arm):
                logwritting('Не доступен')
                continue
            # если настройка "только смена роли"
            if no_install:
                makedir_cfg(arm, way_to_copy_xml, user_sttings_inuse)
                continue

            # копирование дистрибутива
            copying(arm, way_to)

            # если флаг первой установки False
            if not first_install:
                # узнаём версию установленного ПО
                print('Проверка установленой версии...')
                tmpprogramm = nowversion(arm, user, pasw, name_programm)
                if tmpprogramm:
                    print(f'Найдена {tmpprogramm}, удаляем...')
                    taskkill(arm, user, pasw)

                    # удаляем установленную старую версию
                    uninstallprogramm(arm, user, pasw, tmpprogramm)
                else:
                    print('Программа не найдена')

            else:  # не первая установка - создаём директории настроек,
                # копируем файл настроек
                makedir_cfg(arm, way_to_copy_xml, user_sttings_inuse)

            # устанавливаем новую версию
            installprogramm(way_to, arm, user, pasw, DT_version)

            # удаляем файлы дистрибутива
            print('Удаляем файлы дистрибутива с удаленного АРМ-а')
            shutil.rmtree(f'\\\\{arm}\\c$\\psexec')

            # если флаг смены роли, перезаписываем файл настроек
            if is_change_roll:
                print(
                    f'Копируем следующие файлы настроек: {user_sttings_inuse}'
                )
                shutil.copyfile(
                    user_sttings_inuse, f'\\\\{arm}\\{way_to_copy_xml}')

            # проверяем версию установленной программы
            print('Проверяем версию установленной программы...')
            nowinstall = nowversion(arm, user, pasw, name_programm)
            if nowinstall:
                print(f'Установка {nowinstall} на {arm} завершена')
            else:
                print('Программы не найдено')
            logwritting(nowinstall)

else:
    print('Выход из программы инсталяции...')

# зрительно отделяем последнюю этерацию
print(linebreake)
print(linebreake)

