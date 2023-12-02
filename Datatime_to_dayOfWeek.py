import datetime


def dow(date_input_1):
    date = datetime.datetime.strptime(date_input_1, '%d.%m.%Y')
    exec(f'date = datetime.date({date.year}, {date.month}, {date.day})')
    
    if date.weekday() == 0:
        day = 'Понедельник'
    elif date.weekday() == 1:
        day = 'Вторник'
    elif date.weekday() == 2:
        day = 'Среда'
    elif date.weekday() == 3:
        day = 'Четверг'
    elif date.weekday() == 4:
        day = 'Пятница'
    elif date.weekday() == 5:
        day = 'Суббота'
    elif date.weekday() == 6:
        day = 'Воскресенье'
    return day


if __name__ == '__main__':
    date_input = input('введи дату: ')
    print(dow(date_input))
