# encoding: windows-1251

import requests
import pandas as pd
from dotenv import dotenv_values
from datetime import datetime, timedelta

env_vars = dotenv_values('variables.env')
tasks_webhook = str(env_vars['tasks_webhook'])
employees_webhook = str(env_vars['employees_webhook'])


def create_dataframe(date):
    next_date = date + timedelta(days=1)

    task_url = f'https://{tasks_webhook}/task.elapseditem.getlist.json?order[ID]=ASC&filter[>CREATED_DATE]={date}&filter[<CREATED_DATE]={next_date}&select[]=*&PARAMS[NAV_PARAMS][nPageSize]=50&PARAMS[NAV_PARAMS][iNumPage]=1'
    tasks_response = requests.get(task_url)
    tasks_json = tasks_response.json()
    df = pd.DataFrame(columns=['Сотрудник', 'Время, часы'])

    for task in tasks_json['result']:
        user_id = task['USER_ID']
        time_spent = task['SECONDS']

        user_response = requests.get(f'https://{employees_webhook}/user.get.json?ID={user_id}')
        user_json = user_response.json()
        user_first_name = ''
        user_last_name = ''
        for field in user_json['result']:
            user_first_name = field['NAME']
            user_last_name = field['LAST_NAME']

        temp_df = pd.DataFrame({'Сотрудник': [f'{user_first_name} {user_last_name}'],
                                'Время, часы': [int(time_spent) / 3600]})

        df = pd.concat([df, temp_df], ignore_index=True)

    return df
