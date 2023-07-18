# encoding: windows-1251

import requests
import pandas as pd
from dotenv import dotenv_values
from datetime import timedelta

env_vars = dotenv_values('variables.env')
tasks_webhook = str(env_vars['tasks_webhook'])
employees_webhook = str(env_vars['employees_webhook'])
task_description_webhook = str(env_vars['task_description_webhook'])


def count_pages(date, next_date):
    request = requests.get(f'https://{tasks_webhook}/task.elapseditem.getlist.json?order[ID]=ASC&filter[>CREATED_DATE]={date}&filter[<CREATED_DATE]={next_date}&select[]=*&PARAMS[NAV_PARAMS][nPageSize]=50&PARAMS[NAV_PARAMS][iNumPage]=1')
    request_json = request.json()
    total = request_json['total']
    pages = (int(total) // 50) + 1
    return pages


def create_dataframe(date, employees):
    next_date = date + timedelta(days=1)
    pages = count_pages(date, next_date)
    df = pd.DataFrame(columns=['Сотрудник', 'Время, часы'])

    for num_page in range(1, pages + 1):
        task_url = f'https://{tasks_webhook}/task.elapseditem.getlist.json?order[ID]=ASC&filter[>CREATED_DATE]={date}&filter[<CREATED_DATE]={next_date}&select[]=*&PARAMS[NAV_PARAMS][nPageSize]=50&PARAMS[NAV_PARAMS][iNumPage]={num_page}'
        tasks_response = requests.get(task_url)
        tasks_json = tasks_response.json()

        for task in tasks_json['result']:
            user_id = task['USER_ID']
            time_spent = task['SECONDS']
            task_id = task['TASK_ID']

            task_description = requests.get(f'https://{task_description_webhook}/task.item.getdata.json?task_id={task_id}')
            if 'error' in task_description.json().keys():
                continue

            user_response = requests.get(f'https://{employees_webhook}/user.get.json?ID={user_id}')
            user_json = user_response.json()
            user_first_name = ''
            user_last_name = ''
            for field in user_json['result']:
                user_first_name = field['NAME']
                user_last_name = field['LAST_NAME']
            full_name = f'{user_first_name} {user_last_name}'

            if full_name in employees:
                temp_df = pd.DataFrame({'Сотрудник': [f'{user_first_name} {user_last_name}'],
                                        'Время, часы': [int(time_spent) / 3600]})
                df = pd.concat([df, temp_df], ignore_index=True)

    return df
