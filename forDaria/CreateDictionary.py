import asyncio
import json
import requests
from dotenv import dotenv_values
from datetime import timedelta
import re
from fast_bitrix24 import Bitrix
from fast_bitrix24.server_response import ErrorInServerResponseException


env_vars = dotenv_values('variables.env')

tasks_webhook = str(env_vars['tasks_webhook'])
employees_webhook = str(env_vars['employees_webhook'])
task_description_webhook = str(env_vars['task_description_webhook'])
group_webhook = str(env_vars['group_webhook'])
group_list_webhook = str(env_vars['group_list_webhook'])


def get_month_statistic(date):
    formatted_date = date.strftime("%m.%Y")
    url = f'https://production-calendar.ru/get/ru/{formatted_date}/json?region=16'
    request = requests.get(url)
    request_json = request.json()
    month = {}
    for day in range(len(request_json['days'])):
        month[int(request_json['days'][day]['date'].split('.')[0])] = request_json['days'][day]['week_day']
    return month


def get_date_statistic(date):
    formatted_date = date.strftime("%d.%m.%Y")
    url = f'https://production-calendar.ru/get/ru/{formatted_date}/json?region=16'
    request = requests.get(url)
    request_json = request.json()
    return {'day': int(request_json['days'][0]['date'].split('.')[0]), 'week_day': request_json['days'][0]['week_day']}


def count_pages(url):
    request = requests.get(url)
    request_json = request.json()
    total = request_json['total']
    pages = (int(total) // 50) + 1
    return pages


def create_groups_dict():
    pages = count_pages(f"https://{group_webhook}/sonet_group.get.json")
    group_names = {}
    for num_page in range(1, pages + 1):
        request = requests.get(f"https://{group_webhook}/sonet_group.get.json?PARAMS[NAV_PARAMS][nPageSize]=50&PARAMS[NAV_PARAMS][iNumPage]={num_page}")
        request_json = request.json()
        total = int(request_json['total'])
        groups_count = total % 50
        for i in range(0, groups_count):
            id = request_json['result'][i]['ID']
            name = request_json['result'][i]['NAME']
            group_names[id] = name
    print(group_names)
    return group_names


async def get_projects():
    with open('projects.txt', 'r') as file:
        content = file.read()
    needed_groups = content.split(',')
    webhook = f'https://{group_list_webhook}'
    b = Bitrix(webhook)
    group_names = create_groups_dict()
    groups_info = {}
    for group in needed_groups:
        group_information = await b.get_all(
            'tasks.task.list',
            params={
                'select': ['TITLE'],
                'filter': {'GROUP_ID': group, 'UF_AUTO_649598802341': 'ТИТУЛ'}
            },
        )
        groups_info[group_names[group]] = group_information
    return groups_info


async def create_dictionary(date, employees):
    next_date = date + timedelta(days=1)
    pages = count_pages(f"https://{tasks_webhook}/task.elapseditem.getlist.json?order[ID]=ASC&filter[>CREATED_DATE]={date}&filter[<CREATED_DATE]={next_date}&SELECT[0]=*&PARAMS[NAV_PARAMS][nPageSize]=50&PARAMS[NAV_PARAMS][iNumPage]=1")
    group_names = create_groups_dict()

    result = {}
    for num_page in range(1, pages + 1):
        request = requests.get(f"https://{tasks_webhook}/task.elapseditem.getlist.json?order[ID]=ASC&filter[>CREATED_DATE]={date}&filter[<CREATED_DATE]={next_date}&SELECT[0]=*&PARAMS[NAV_PARAMS][nPageSize]=50&PARAMS[NAV_PARAMS][iNumPage]={num_page})")
        request_json = request.json()
        elapsed_time = {}
        users_id = []
        tasks_id = []
        for field in request_json["result"]:
            user_id = field['USER_ID']
            task_id = field['TASK_ID']
            users_id.append(user_id)
            tasks_id.append(task_id)
            if user_id not in elapsed_time:
                elapsed_time[user_id] = {}
            elapsed_time[user_id][task_id] = field['SECONDS']
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
        errors = []

        async def get_task_by_ID(b_task, task_id):
            try:
                task = await b_task.get_by_ID('task.item.getdata', {str(task_id)})
                return task

            except ErrorInServerResponseException as e:
                errors.append(task_id)

        tasks_coroutines = [get_task_by_ID(b_task, task_id) for task_id in tasks_id]
        tasks_result = await asyncio.gather(*tasks_coroutines)
        tasks = {task_id: task for task_id, task in zip(tasks_id, tasks_result)}

        for item in elapsed_time.keys():
            user_json = users[item][0]
            if 'PERSONAL_CITY' in user_json and user_json['PERSONAL_CITY'] != '':
                user_city = re.sub(r'\([^)]*\)', '', user_json['PERSONAL_CITY'])
                user_city = user_city.replace(' ', '')
            else:
                user_city = 'Нет поля город'

            if 'WORK_POSITION' in user_json:
                user_work_position = user_json['WORK_POSITION']
            else:
                user_work_position = 'Нет должности'
            user_first_name = user_json['NAME']
            user_last_name = user_json['LAST_NAME']
            full_name = f'{user_first_name} {user_last_name}'

            if full_name not in employees:
                continue

            for task in elapsed_time[item].keys():
                if task in errors:
                    continue
                task_json = tasks[task][task]

                group_id = task_json['GROUP_ID']

                parent_id = task_json['PARENT_ID']

                if group_id in group_names.keys():
                    name = group_names[group_id]
                else:
                    name = str(group_id)

                while parent_id != 0 and parent_id is not None:
                    parent_task = requests.get(f'https://{task_description_webhook}/task.item.getdata.json?task_id={parent_id}').json()
                    if 'error' in parent_task:
                        break
                    if parent_task['result']['UF_AUTO_649598802341'] == 'ТИТУЛ':
                        name += f' {parent_task["result"]["TITLE"].split(" - ")[0]}'
                        break
                    parent_id = parent_task['result']['PARENT_ID']

                if full_name not in result:
                    result[full_name] = {'work_position': '', 'group_and_hours': {}, 'total_time': 0, 'city': ''}

                if name not in result[full_name]['group_and_hours']:
                    result[full_name]['group_and_hours'][name] = 0

                result[full_name]['work_position'] = user_work_position
                result[full_name]['group_and_hours'][name] += int(elapsed_time[item][task]) / 3600
                result[full_name]['total_time'] += int(elapsed_time[item][task]) / 3600
                result[full_name]['city'] = user_city

    print(json.dumps(result, indent=4, ensure_ascii=False))

    return result