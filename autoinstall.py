import os
import shutil
import pythonping
import datetime


# берем версию программы установленной сейчас
def nowversion():
    os.system(f'wmic /NODE:\"{arm}\" /USER:\"{user}\" /password: \"{pasw}\" product get name| findstr \"{name_programm}\">\"{way_to}\\temp\\\"{arm}.txt')
    if os.stat(f'{way_to}\\temp\\{arm}.txt').st_size == 0:
        print('Программы не найдено')
        tmp = ''
    with open(f'{way_to}\\temp\\{arm}.txt') as t:
        for tmp in t:
            tmp = tmp.encode('cp1251').decode('cp866').strip()

            break
    return tmp


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


def copying():
    print(f'Copy distrib file to {arm}')
    shutil.copytree(f'{way_to}\\soft', f'\\\\{arm}\\c$\\psexec')
    print('file copyng complete')


def taskkill():
    print(f'close  DispatchTerminal program on {arm}')
    #os.system('taskkill.exe /s ' + row + ' /u ' + user + ' /p ' + pasw + '  /F /T /IM  DispatchTerminal.exe')
    os.system(f'taskkill.exe /s {arm} /u {user} /p {pasw} /F /T /IM  DispatchTerminal.exe')


def makedir_cfg():
    print('Создаём необходимые директории и копируем файл настроек')
    dirprotei = f'\\\\{arm}\\c$\\ProgramData\\Protei'
    dirdispach = f'\\\\{arm}\\c$\\ProgramData\\Protei\\DispatchTerminal'
    try:
        os.mkdir(dirprotei)
    except OSError:
        print(f'Создать директорию {dirprotei} не удалось')
    try:
        os.mkdir(dirdispach)
    except OSError:
        print(f'Создать директорию {dirdispach} не удалось')
    try:
        shutil.copyfile(user_sttings_inuse, f'\\\\{arm}\\{way_to_copy_xml}')
    except OSError:
        print(f'Не удалось скопировать файл настроек')


def makdir_files(way_to):
    # проверяем наличие temp папки
    dirtemp = f'{way_to}\\temp'
    try:
        os.mkdir(dirtemp)
        print('make temp directory')
    except OSError:
        print(f'cant make Directory {dirtemp}, allready exist? ')


def uninstallprogramm(arm, user, pasw, tmpprogramm):
    # os.system('wmic /NODE:\"' + arm + '\" /USER:\"' + user + '\" /password: \"' + pasw + '\" product where description=\"' + tmpprogramm + '\" uninstall')
    os.system(f'wmic /NODE:\"{arm}\" /USER:\"{user}\" /password: \"{pasw}\" product where description=\"{tmpprogramm}\" uninstall')
    print(f'{tmpprogramm} is uninstalled')


def installprogramm(way_to, arm, user, pasw, dt_version):
    print('installing new version of soft')
    os.system(
        f'{way_to}\\psexec.exe \\\\{arm} -u {user} -p {pasw} -h msiexec.exe /i \"C:\\psexec\\{dt_version}\"')


# заводим переменные
name_programm = 'Дежурн'  # поиск программы ведется по этому словосочетанию
DT_version = 'DT-7.11.13-release-Spb-37706.msi'  # фактическое название файла утсановки + визуально понятно какая версия
# динамичное определение папки, где всё хранится
way_to = os.path.abspath(__file__)
way_to = os.path.dirname(way_to)
way_to_copy_xml = 'c$\\ProgramData\\Protei\\DispatchTerminal\\UserSettings.xml'  # где лежит файл сеттингов на удаленной ммашине
oper112_xml = f'{way_to}\\usersetting\\oper112\\UserSettings.xml'
nachsmen_xml = f'{way_to}\\usersetting\\nachsmen\\UserSettings.xml'
user_sttings_inuse = nachsmen_xml
is_change_roll = False  # ставим флаг на смену роли
first_install = False
linebreake = '********************'

# подгатавливаем файл логов к записи событий данной сессии
with open(f'{way_to}\\log.txt', 'w+') as logfile:
    logfile.write(str(datetime.datetime.now()) + '\n')
makdir_files(way_to)
with open(f'{way_to}\\list.txt') as list_of_arms:
    print(f'Warning! automatic install {DT_version} will be iniciated on hosts below!')
    for row_t in list_of_arms:
        print(row_t)
print(f'way to UserSettings {user_sttings_inuse}')
if first_install:
    print('first instalation')
    is_change_roll = False  # глупо менять роль при первой установке
else:
    print('old version DT will be removed')
    print('Dispatch.exe process will be stoped')
print(f'change role = {is_change_roll}')
question_yn = input('Do you want to continue? Y/N ')
if question_yn.lower() == 'y':
    user = 'gmc\\' + input('type username ')
    pasw = input('type password ')
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
                logwritting('не доступен')
                continue
            # копирование дистрибутива
            try:
                copying()
            except FileExistsError:
                print('Folder allready exist, deleting')
                shutil.rmtree(f'\\\\{arm}\\c$\\psexec')
                copying()

            if not first_install:
                # узнаём версию установленного ПО
                print('searching old soft')
                tmpprogramm = nowversion()
                if tmpprogramm != '':
                    print(f'find {tmpprogramm}, uninstalling')
                    taskkill()
                    # удаляем установленную старую версию
                    uninstallprogramm(arm, user, pasw, tmpprogramm)
                else:
                    taskkill()
            # создаём директории настроек, копируем файл настроек
            else:
                makedir_cfg()
            # устанавливаем новую версию
            # print('installing new version of soft')
            # os.system(f'{way_to}\\psexec.exe \\\\{arm} -u {user} -p {pasw} -h msiexec.exe /i \"C:\\psexec\\{DT_version}\"')
            installprogramm(way_to, arm, user, pasw, DT_version)
            print('delete remote folder with distrib file')
            shutil.rmtree(f'\\\\{arm}\\c$\\psexec')
            if is_change_roll:
                print(f'copying role of {user_sttings_inuse}')
                shutil.copyfile(
                    user_sttings_inuse, f'\\\\{arm}\\{way_to_copy_xml}')
            print('checking version')
            nowinstall = nowversion()
            print(f'installing {nowinstall} on {arm} complete')
            logwritting(nowinstall)

else:
    print('escaping instalation...')
# зрительно отделяем последнюю этерацию
print(linebreake)
print(linebreake)

# циклим чтобы при исполнении ехе не закрывался терминал
'''while True:
    pass'''
