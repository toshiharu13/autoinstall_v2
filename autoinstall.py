import os
import shutil
import pythonping
import datetime


# берем версию программы установленной сейчас
def nowversion():
    os.system('wmic /NODE:\"' + row + '\" /USER:\"' + user + '\" /password: \"' + pasw + '\" product get name| findstr \"' + name_programm + '\">\"' + way_to + '\\temp\\\"' + row + '.txt')
    if os.stat(f'{way_to}\\temp\\{row}.txt').st_size == 0:
        print('Программы не найдено')
        tmp = ''
    with open(f'{way_to}\\temp\\{row}.txt') as t:
        for tmp in t:
            tmp = tmp.encode('cp1251').decode('cp866').strip()

            break
    return tmp

def if_ARM_online(arm):
    try:
        response = pythonping.ping(arm, count=1)
        if response.success():
            print(f"{arm} доступен")
        else:
            print(f"{arm} не доступен")
            return False
    except:
        print(f"{arm} не доступен")
        return False

def logwritting(data):
    with open(f'{way_to}\\log.txt', 'a')as logfile:
        logfile.write(data + '\n')


def Copying():
    print(f'Copy distrib file to {row}')
    shutil.copytree(way_to + "\soft", '\\\\' + row + "\c$\psexec")
    print('file copyng complete')

def taskkill():
    print(f'close  DispatchTerminal program on {row}')
    os.system('taskkill.exe /s ' + row + ' /u ' + user + ' /p ' + pasw + '  /F /T /IM  DispatchTerminal.exe')

def makedir():
    print('Создаём необходимые директории и копируем файл настроек')
    dirprotei = f'\\\\{row}\\c$\\ProgramData\\Protei'
    dirdispach = f'\\\\{row}\\c$\\ProgramData\\Protei\\DispatchTerminal'
    try:
        os.mkdir(dirprotei)
    except OSError:
        print(f'Создать директорию {dirprotei} не удалось')
    try:
        os.mkdir(dirdispach)
    except OSError:
        print(f'Создать директорию {dirdispach} не удалось')
    try:
        shutil.copyfile(user_sttings_inuse, f'\\\\{row}\\{way_to_copy_xml}')
    except OSError:
        print(f'Не удалось скопировать файл настроек')


# заводим переменные
name_programm = 'Дежурн' # поиск программы ведется по этому словосочетанию
DT_version = 'DT-7.11.12-release-Spb-37552.msi' # фактическое название файла утсановки + визуально понятно какая версия
# динамичное определение папки, где всё хранится
way_to = os.path.abspath(__file__)
way_to = os.path.dirname(way_to)
way_to_copy_xml = 'c$\\ProgramData\\Protei\\DispatchTerminal\\UserSettings.xml' # где лежит файл сеттингов на удаленной ммашине
oper112_xml = way_to + '\\usersetting\\oper112\\UserSettings.xml'
zamnachsmen_xml = way_to + '\\usersetting\\zamnachsmen\\UserSettings.xml'
user_sttings_inuse = oper112_xml
is_change_roll = False # ставим флаг на смену роли
first_install = True
linebreake = '********************'

# подгатавливаем файл логов к записи событий данной сессии
with open(f'{way_to}\\log.txt', 'w') as logfile:
    logfile.write(str(datetime.datetime.now()) + '\n')

with open(f'{way_to}\list.txt') as list_of_arms:
    print(f'Warning! automatic install {DT_version} will be iniciated on hosts below!')
    for row_t in list_of_arms:
        print(row_t)
print(f'way to UserSettings {user_sttings_inuse}')
if first_install:
    print('first instalation')
else:
    print('old version DT will be removed')
    print('Dispatch.exe process will be stoped')
print(f'change role = {is_change_roll}')
question_yn = input('Do you want to continue? Y/N ')
if question_yn.lower() == 'y':
    user = 'gmc\\' + input('type username ')
    pasw = input('type password ')
    # читаем названия армов из файла list.txt
    with open(f'{way_to}\list.txt') as list_of_arms:
        for row in list_of_arms:
            row = row.strip()
            logwritting(linebreake)
            logwritting(row)
            # отделить строчкой утсановку разных армов
            print(linebreake)

            # определяем, доступен ли АРМ
            if  if_ARM_online(row) == False:
                logwritting('не доступен')
                continue
            # копирование дистрибутива
            try:
                Copying()
            except FileExistsError:
                print('Folder allready exist, deleting')
                shutil.rmtree(f'\\\\{row}\\c$\\psexec')
                Copying()

            # узнаём версию установленного ПО
            print('searching old soft')
            tmpprogramm = nowversion()
            if tmpprogramm != '':
                print(f'find {tmpprogramm}, uninstalling')
                taskkill()
                # удаляем установленную старую версию
                os.system('wmic /NODE:\"' + row + '\" /USER:\"' +user + '\" /password: \"' + pasw + '\" product where description=\"' + tmpprogramm +'\" uninstall')
                print(f'{tmpprogramm} is uninstalled')
            else:
                taskkill()
            # создаём директории настроек, копируем файл настроек
            if first_install:
                makedir()
            # устанавливаем новую версию
            print('installing new version of soft')
            os.system(way_to + '\\psexec.exe \\\\' + row + ' -u ' + user + ' -p '+ pasw + ' -h msiexec.exe /i \"C:\\psexec\\' + DT_version + '\"')
            print('delete remote folder with distrib file')
            shutil.rmtree(f'\\\\{row}\\c$\\psexec')
            if is_change_roll== True:
                print(f'copying role of {user_sttings_inuse}')
                shutil.copyfile(user_sttings_inuse, f'\\\\{row}\\{way_to_copy_xml}')
            print('checking version')
            nowinstall = nowversion()
            print(f'installing {nowinstall} on {row} complete')
            logwritting(nowinstall)



else:
    print('escaping instalation...')
# зрительно отделяем последнюю этерацию
print(linebreake)
print(linebreake)

# циклим чтобы при исполнении ехе не закрывался терминал
'''while True:
    pass'''
