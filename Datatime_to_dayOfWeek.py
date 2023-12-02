import datetime
while True:
    date_input = input('введи дату: ')
    date_split = date_input.split('.')
    try:
        exec(f'date = datetime.date({date_split[2]}, {date_split[1]}, {date_split[0]})')
        if date.weekday() == 0:
            print('Понедельник')
        elif date.weekday() == 1:
            print('Вторник')
        elif date.weekday() == 2:
            print('Среда')
        elif date.weekday() == 3:
            print('Четверг')
        elif date.weekday() == 4:
            print('Пятница')
        elif date.weekday() == 5:
            print('Суббота')
        elif date.weekday() == 6:
            print('Воскресенье')
    except:
        print('неверный формат!')
    otv = input('еще раз? y/n?: ')
    if otv == 'y':
        continue
    else:
        print('ну и пиздуй')
        exit()
