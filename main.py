from argparse import ArgumentParser
from datetime import datetime, timedelta
from docxtpl import DocxTemplate
from random import randint


BREAK_INTERVAL = (90, 120)  # перерыв между патрулями в минутах
STEP = 5    # шаг для округления времени патруля в минутах
COLUMNS_NUMBER = 2  # количество колонок в таблице с патрулями


def get_next_patrol(prev_patrol):
    break_time = timedelta(minutes=STEP*round(randint(90, 120)/STEP))
    return prev_patrol + break_time


def get_patrol_columns(date, start_time, period):
    full_start_time = datetime.strptime(f"{date}, {start_time}", '%d.%m.%Y, %H:%M')
    end_time = full_start_time + timedelta(hours=period)
    first_patrol = full_start_time + timedelta(minutes=(randint(0, 15)))

    patrol_columns = []
    patrols = []

    patrol_time = first_patrol
    while patrol_time < end_time:
        patrols.append(str(patrol_time))
        patrol_time = get_next_patrol(patrol_time)

    patrols_in_column = int(len(patrols) / COLUMNS_NUMBER)
    for column in range(COLUMNS_NUMBER):
        patrol_columns.append(
            patrols[(column*patrols_in_column):(column*patrols_in_column+patrols_in_column)]
        )

    return patrol_columns


def get_schedules(storage_objects_names, date, start_time, period):
    schedules = []
    for storage_object_number, storage_object_name in enumerate(storage_objects_names):
        storage_object = {
            'number': storage_object_number+1,
            'name': storage_object_name,
            'patrol_columns': get_patrol_columns(date, start_time, period),
        }
        schedules.append(storage_object)
    return schedules


def create_documents():
    template_doc_name = 'schedule_template.docx'
    args = create_parser().parse_args()
    date = args.date
    schedules_names = args.schedules.split(',')

    for doc in range(args.number):
        template = DocxTemplate(template_doc_name)
        formated_date = datetime.strptime(date, '%d.%m.%Y')

        context = {
            'schedules': get_schedules(schedules_names, date, args.start_time, args.period),
            'date': date,
            'next_day': datetime.strftime(formated_date + timedelta(days=1), '%d.%m.%Y'),
        }

        template.render(context)
        template.save(f'schedule{date}.docx')
        date = datetime.strftime(formated_date + timedelta(days=1), '%d.%m.%Y')


def create_parser():
    parser = ArgumentParser(description='description')
    parser.add_argument(
        '-d',
        '--date',
        help='Дата, для которой создаётся расписание, в формате: 01.01.2022. По умолчанию - завтра',
        type=str,
        default=datetime.strftime(datetime.today()+timedelta(days=1), '%d.%m.%Y')
    ),
    parser.add_argument(
        '-n',
        '--number',
        help='Количество дней подряд, для которых будут созданы расписания (начиная от указанного в параметре -d дня)',
        type=int,
        default=1,
    ),
    parser.add_argument(
        '-st',
        '--start_time',
        help='Время начала смены в формате: 08:00',
        type=str,
        default='08:00',
    ),
    parser.add_argument(
        '-p',
        '--period',
        help='Продолжительность смены в часах, например: 24',
        type=float,
        default=24.0,
    )
    parser.add_argument(
        '-s',
        '--schedules',
        help='Названия всех расписаний (например, по объекту склада) через запятую без пробела: ГСМ,Стоянка,Ангар',
        type=str,
        default='ГСМ,Стоянка,Ангар'
    )
    return parser


if __name__ == '__main__':
    create_documents()
