import asyncio

import requests
import pandas as pd
from dotenv import dotenv_values
from datetime import timedelta, datetime
from fast_bitrix24 import Bitrix
from fast_bitrix24.server_response import ErrorInServerResponseException

env_vars = dotenv_values('variables.env')
tasks_webhook = str(env_vars['tasks_webhook'])
employees_webhook = str(env_vars['employees_webhook'])
task_description_webhook = str(env_vars['task_description_webhook'])


def get_date_statistic(date):
    formatted_date = date.strftime("%d.%m.%Y")
    url = f'https://production-calendar.ru/get/ru/{formatted_date}/json?region=16'
    request = requests.get(url)
    request_json = request.json()
    return {'holiday': request_json["statistic"]['holidays'], 'working_hours': request_json["statistic"]['working_hours'], 'weekend': request_json["statistic"]['weekends']}


def get_month_statistic(date):
    formatted_date = date.strftime("%m.%Y")
    url = f'https://production-calendar.ru/get/ru/{formatted_date}/json?region=16'
    request = requests.get(url)
    request_json = request.json()
    month = {}
    for day in range(len(request_json['days'])):
        month[int(request_json['days'][day]['date'].split('.')[0])] = request_json['days'][day]['week_day']
    return month


def count_pages(date, next_date):
    request = requests.get(f'https://{tasks_webhook}/task.elapseditem.getlist.json?order[ID]=ASC&filter[>CREATED_DATE]={date}&filter[<CREATED_DATE]={next_date}&select[]=*&PARAMS[NAV_PARAMS][nPageSize]=50&PARAMS[NAV_PARAMS][iNumPage]=1')
    request_json = request.json()
    total = request_json['total']
    pages = (int(total) // 51) + 1
    return pages


async def create_dataframe(date, employees):
    next_date = date + timedelta(days=1)
    pages = count_pages(date, next_date)
    df = pd.DataFrame(columns=['Сотрудник', 'Время, часы'])

    for num_page in range(1, pages + 1):
        task_url = f'https://{tasks_webhook}/task.elapseditem.getlist.json?order[ID]=ASC&filter[>CREATED_DATE]={date}&filter[<CREATED_DATE]={next_date}&select[]=*&PARAMS[NAV_PARAMS][nPageSize]=50&PARAMS[NAV_PARAMS][iNumPage]={num_page}'
        tasks_response = requests.get(task_url)
        tasks_json = tasks_response.json()
        elapsed_time = {}
        users_id = []
        tasks_id = []

        for task in tasks_json['result']:
            user_id = task['USER_ID']
            time_spent = task['SECONDS']
            task_id = task['TASK_ID']
            users_id.append(user_id)
            tasks_id.append(task_id)
            if user_id not in elapsed_time:
                elapsed_time[user_id] = {}
            elapsed_time[user_id][task_id] = time_spent

        set(users_id)

        b_user = Bitrix(f'https://{employees_webhook}')
        b_task = Bitrix(f'https://{tasks_webhook}')

        if not users_id:
            continue

        users = await b_user.get_by_ID(
            'user.get',
            {str(user_id) for user_id in users_id})
        if isinstance(users, list) and len(users) == 1:
            users = {users[0]['ID']: [users[0]]}

        print(users)
        errors = []

        async def get_task_by_ID(b_task, task_id):
            try:
                task = await b_task.get_by_ID('task.item.getdata', {str(task_id)})
                return task
            except ErrorInServerResponseException as e:
                print("Error in server response:", e)
                errors.append(task_id)

        tasks_coroutines = [get_task_by_ID(b_task, task_id) for task_id in tasks_id]
        tasks = await asyncio.gather(*tasks_coroutines)

        print(elapsed_time)

        for item in elapsed_time.keys():
            for task in elapsed_time[item].keys():
                if task in errors:
                    continue
                user_json = users[item][0]
                user_first_name = user_json['NAME']
                user_last_name = user_json['LAST_NAME']
                full_name = f'{user_first_name} {user_last_name}'

                if full_name in employees:
                    temp_df = pd.DataFrame({'Сотрудник': [f'{user_first_name} {user_last_name}'],
                                            'Время, часы': [int(elapsed_time[item][task]) / 3600]})
                    df = pd.concat([df, temp_df], ignore_index=True)

    return df
