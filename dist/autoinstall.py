import os
import shutil
import pythonping
import datetime

def nowversion(): #берем версию программы установленной сейчас
    os.system('wmic /NODE:\"' + row + '\" /USER:\"' + user + '\" /password: \"' + pasw + '\" product get name| findstr \"' + name_programm + '\">\"' + way_to + '\\temp\\\"' + row + '.txt')
    if os.stat(way_to + '\\temp\\' + row + '.txt').st_size == 0:
        print('Программы не найдено')
        tmp = ''
    with open(way_to + '\\temp\\' + row + '.txt') as t:
        for tmp in t:
            tmp = tmp.encode('cp1251').decode('cp866').strip()

            break
    return tmp

def if_ARM_online(arm):
    try:
        response = pythonping.ping(arm, count=1)
        if response.success():
            print(arm + ' доступен')
        else:
            print(arm + ' не доступен')
            return False
    except:
        print(arm + ' не доступен')
        return False

def logwritting(data):
    with open(way_to + '\\log.txt', 'a')as logfile:
        logfile.write(data + '\n')

def Copying():
    print('Copy distrib file to ' + row)
    shutil.copytree(way_to + "\soft", '\\\\' + row + "\c$\psexec")
    print('file copyng complete')


## заводим переменные
name_programm = 'Дежурн' #поиск программы ведется по этому словосочетанию
DT_version = 'DT-5.14.7-release-Spb-35245.msi' #фактическое название файла утсановки + визуально понятно какая версия
#динамичное определение папки, где всё хранится
way_to = os.path.abspath(__file__)
way_to = os.path.dirname(way_to)

way_to_copy_xml = 'c$\\ProgramData\\Protei\\DispatchTerminal\\UserSettings.xml' # где лежит файл сеттингов на удаленной ммашине
oper112_xml = way_to + '\\usersetting\\oper112\\UserSettings.xml'
zamnachsmen_xml = way_to + '\\usersetting\\zamnachsmen\\UserSettings.xml'
user_sttings_inuse = oper112_xml
linebreake = '********************'

#подгатавливаем файл логов к записи событий данной сессии
with open(way_to + '\\log.txt', 'w') as logfile:
    logfile.write(str(datetime.datetime.now()) + '\n')

with open(way_to + "\list.txt") as list_of_arms:
    print('Warning! automatic install' + DT_version + ' will be iniciated on hosts below!')
    for row_t in list_of_arms:
        print(row_t)
print('way to UserSettings ' + user_sttings_inuse)
print('old version DT will be removed')
print('Dispatch.exe process will be stoped')
question_yn = input('Do you want to continue? Y/N ')
if question_yn.lower() == 'y':
    user = 'gmc\\' + input('type username ')
    pasw = input('type password ')

    with open(way_to + "\list.txt") as list_of_arms: #читаем названия армов из файла list.txt
        for row in list_of_arms:
            row = row.strip()
            logwritting(linebreake)
            logwritting(row)
            print(linebreake)  # отделить строчкой утсановку разных армов

            #определяем, доступен ли АРМ
            if  if_ARM_online(row) == False:
                logwritting('не доступен')
                continue
            try:#копирование дистрибутива
                Copying()
            except FileExistsError:
                print('Folder allready exist, deleting')
                shutil.rmtree('\\\\' + row + "\c$\psexec")
                Copying()

            print('close  DispatchTerminal program on ' + row)
            os.system('taskkill.exe /s ' + row + ' /u ' + user + ' /p ' + pasw + '  /F /T /IM  DispatchTerminal.exe')

            #узнаём версию установленного ПО
            print('searching old soft')
            tmpprogramm = nowversion()
            if tmpprogramm != '':
                print('uninstalling ' + tmpprogramm)

                #удаляем установленную старую версию
                os.system('wmic /NODE:\"' + row + '\" /USER:\"' +user + '\" /password: \"' + pasw + '\" product where description=\"' + tmpprogramm +'\" uninstall')
                print(tmpprogramm + 'is uninstalled')
            else:
                # если программы нет, предполагаем, что первый раз, копируем файл настроек
                print('copyng xml file of user settings')
                shutil.copyfile(user_sttings_inuse, '\\\\' + row +'\\' + way_to_copy_xml)

            #устанавливаем новую версию
            print('installing new version of soft')
            os.system(way_to + '\\psexec.exe \\\\' + row + ' -u ' + user + ' -p '+ pasw + ' -h msiexec.exe /i \"C:\\psexec\\' + DT_version + '\"')
            print('delete remote folder with distrib file')
            shutil.rmtree('\\\\' + row + "\c$\psexec")
            print('checking version')
            nowinstall = nowversion()
            print('installing ', nowinstall, 'on ' + row +' complete')
            logwritting(nowinstall)



else:
    print('escaping instalation...')
#зрительно отделяем последнюю этерацию
print(linebreake)
print(linebreake)

#циклим чтобы при исполнении ехе не закрывался терминал
while True:
    pass



