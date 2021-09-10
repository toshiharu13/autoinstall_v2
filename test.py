import autoinstall

def test_ip_pinging(testing_func):
    """Тест функции доступности АРМа."""
    # Напечатать информацию о тесте из docstring
    print(f'{test_ip_pinging.__doc__}')

    # Напечатать имя тестируемой функции
    print(f'Тестирование алгоритма {testing_func.__name__}')

    arm = 'kv-c112-181'

    assert testing_func(arm) == True, ('Что-то пошло не так')

    print(f'Тест для {testing_func.__name__} пройден успешно')


test_ip_pinging(autoinstall.if_arm_online)