import math
import time

import numpy as np
import pandas as pd
import requests
from yandex_geocoder import Client
from sqlalchemy import create_engine
from datetime import datetime
import uuid
import openpyxl


class TaskDistributor:
    def __init__(self, points_task_file, employees_file, task_file, task_type, yandex_api_key, openroute_api_key):
        self.task = task_file
        self.task_type = task_type[['uuid', 'priority']]
        self.employees = employees_file
        self.points_task = points_task_file

        self.yandex_api_key = yandex_api_key
        self.openroute_api_key = openroute_api_key
        self.client = Client(yandex_api_key)
        self.token = openroute_api_key

        self.cnt = 0

    @staticmethod
    def set_task(point):
        """
        Установка для каждой задачи приоритета

        :param point: задача
        :return: приоритет: 1-высокий, 2-средний, 3-низкий
        """

        # try:
        #     point['num_given_cards'] / point['approved_requests']
        # except ZeroDivisionError:
        #         return 100

        last_card_given = (datetime.now().date() - point["last_card_given"]).days

        # Новым задачам выдаем тип
        if point["agent_point_date"] == "вчера" or not point["materials"]:
            return 3
        elif point['num_given_cards'] / point['approved_requests'] <= 0.5:
            return 2
        elif (last_card_given >= 7 and point['approved_requests'] >= 1) or last_card_given >= 14:
            return 1
        else:
            return 4  # Задачи, которые не прошли ни одно условие, считаем как выполненные

    @staticmethod
    def set_grade(employee_grade):
        """
        Установка каждому сотруднику списка приоритетов (для джуна [3] и так далее)

        :param employee_grade: сотрудник
        :return: список приоритетов, который сотрудник может выполнить
        """
        employee_grade = int(employee_grade)
        match employee_grade:
            case 3:
                return [1, 2, 3]
            case 2:
                return [2, 3]
            case 1:
                return [3]

    def matrix(self, address: list[str, str], profile=0):
        """
        Расчёт времени между адресами

        :param address: 2 адреса на входе
        :param profile: 0 - на машине; 1 - пешком
        :return: durations
        """

        coord_1 = self.client.coordinates(address[0])
        coord_2 = self.client.coordinates(address[1])

        locations = [[float(coord_1[1]), float(coord_1[0])], [float(coord_2[1]), float(coord_2[0])]]

        headers = {
            'Content-Type': 'application/json; charset=utf-8',
            'Accept': 'application/json',
            'Authorization': self.token
        }
        profile_dict = {
            0: 'driving-car',
            1: 'foot-walking'
        }
        data = {"locations": [i[::-1] for i in locations], "metrics": ["distance", "duration"], "units": "m"}
        res = requests.post(f'https://api.openrouteservice.org/v2/matrix/{profile_dict[profile]}',
                            headers=headers,
                            json=data).json()
        ps = dict(durations=res['durations'][0][1], distances=res['distances'][0][1])

        self.cnt += 1
        if self.cnt == 40:
            time.sleep(60)
            self.cnt = 0

        return math.ceil(ps["durations"] / 60)

    def check_point(self):
        """
        Перепроверяет выполненные задачи

        :return: если, у задачи появилось новое задание, устанавливаем ему тип задачи
        """

        for index, point in self.points_task.iterrows():
            if point["status"] == 4:
                if (datetime.now().date() - point["appointment_date"]).days >= 7:
                    self.points_task.loc[index, '№ Задачи'] = self.set_task(point)
                    if point["№ Задачи"] != 4:
                        self.task.loc[self.task['uuid'] == point['uuid']] = np.NAN

    def distribute_tasks(self):
        """
        Самая главная функция, выполняющая распределение задач между сотрудниками

        :return: заполненный dataframe employees и points_tomorrow
        """

        for index, point in self.points_task.iterrows():
            task_number = point['№ Задачи']

            if point['status'] in [1, 2, 3, 4]:  # Уже выполнялось ранее и менять не надо
                continue

            if task_number == 4:  # Новая задача, которая не подошла ни под одно из условии
                new_data = {'uuid': str(uuid.uuid4()),
                            'appointment_date': datetime.now().date(),
                            'priority': np.NAN,
                            'status': 4,
                            'agent_point_id': point["uuid_x"],
                            'employee_id': np.NAN,
                            'task_type_id': np.NAN
                            }
                self.task.loc[len(self.task)] = new_data
                continue
            elif task_number == 1:
                task_time = 240
            elif task_number == 2:
                task_time = 120
            elif task_number == 3:
                task_time = 90
            else:
                continue

            best_travel_time = 10000
            best_employee_id = None
            origin = 'Краснодар, ' + point['address']

            suitable_employees = self.employees[self.employees['Решаемые задачи'].apply(lambda x: task_number in x)]
            # suitable_employees = suitable_employees.sort_values(by=['Уровень сотрудника'])

            if task_number == 3:
                junior = suitable_employees[suitable_employees["grade"] == 1]
                if junior[junior['Кол-во отработанных минут'] >= 450]['uuid'].count() < len(junior):
                    suitable_employees = junior
                else:
                    suitable_employees = suitable_employees[suitable_employees["Уровень сотрудника"] != 1]

            elif task_number == 2:
                middle = suitable_employees[suitable_employees["grade"] == 2]
                if middle[middle['Кол-во отработанных минут'] >= 420]['uuid'].count() < len(middle) - 1:
                    suitable_employees = middle
                else:
                    suitable_employees = suitable_employees[suitable_employees["grade"] != 1]

            for e_index, employee in suitable_employees.iterrows():

                if employee["Кол-во отработанных минут"] + task_time >= 540:
                    continue

                destination = employee["address"]
                travel_time = self.matrix([origin, destination])

                if travel_time < best_travel_time:
                    best_travel_time = travel_time
                    best_employee_id = e_index

            if best_employee_id is not None:
                self.employees.loc[best_employee_id, "Номер задачи"] += 1

                if point['status'] is not np.NAN:  # Если это 'вчерашняя' задача, меняем ее
                    uuids = point['uuid']
                    index_task = self.task[(self.task['uuid'] == uuids) & (self.task['status'].between(5, 15))].index

                    self.task.loc[index_task, "appointment_date"] = datetime.now().date()
                    self.task.loc[index_task, "priority"] = self.employees.loc[best_employee_id, "Номер задачи"]
                    self.task.loc[index_task, "status"] = 1
                    self.task.loc[index_task, "employee_id"] = self.employees.loc[best_employee_id, "uuid"],

                else:  # Либо создаем
                    new_data = {'uuid': str(uuid.uuid4()),
                                'appointment_date': datetime.now().date(),
                                'priority': self.employees.loc[best_employee_id, "Номер задачи"],
                                'status': 1,
                                'agent_point_id': point["uuid_x"],
                                'employee_id': self.employees.loc[best_employee_id, "uuid"],
                                'task_type_id':
                                    self.task_type[self.task_type['priority'] == point['№ Задачи']]['uuid'].values[0]
                                }
                    # Добавляем запись в DataFrame
                    self.task.loc[len(self.task)] = new_data

                self.employees.loc[best_employee_id, "Кол-во отработанных минут"] += best_travel_time + task_time
                self.employees.loc[best_employee_id, "address"] = origin

            else:
                if point['status'] is not np.NAN:
                    uuids = point['uuid']
                    index_task = self.task[(self.task['uuid'] == uuids) & (self.task['status'].between(5, 15))].index
                    self.task.loc[index_task, "status"] += 1
                else:
                    status = 5
                    new_data = {'uuid': str(uuid.uuid4()),
                                'appointment_date': self.points_task.loc[index, 'appointment_date'],
                                'priority': np.NAN,
                                'status': status,
                                'agent_point_id': point["uuid_x"],
                                'employee_id': np.NAN,
                                'task_type_id':
                                    self.task_type[self.task_type['priority'] == point['№ Задачи']]['uuid'].values[0]
                                }
                    # Добавляем запись в DataFrame
                    self.task.loc[len(self.task)] = new_data

    def run_distribution(self):
        # Перепроверяем выполненные задачи, если прошло более 7 дней
        self.check_point()

        self.points_task["№ Задачи"] = self.points_task.apply(self.set_task,
                                                              axis=1)

        # Сортируем задачи так, чтобы вчерашние задачи выполнялись первее новых
        self.points_task.sort_values(by=['status', '№ Задачи'], inplace=True, ascending=[False, True])

        self.employees["Решаемые задачи"] = self.employees["grade"].apply(self.set_grade)
        self.employees["Кол-во отработанных минут"] = 0
        self.employees["Номер задачи"] = 0

        # Распределяем задачи
        self.distribute_tasks()

        # Вывод
        self.task.to_csv('points.csv', index=False)
        self.task.to_sql('Task', connection_string, index=False, if_exists='replace')


if __name__ == "__main__":
    connection_string = create_engine(
        'postgresql://fzpvuodf:ZJhLvoXnBVUJRGxCkDNpMvWmddimNytV@cornelius.db.elephantsql.com/fzpvuodf')

    query_for_users = 'SELECT * FROM accounts_user'
    query_for_points = 'SELECT * FROM "AgentPoint"'
    query_for_task = 'SELECT * FROM "Task"'
    query_for_task_type = 'SELECT * FROM "TaskType"'

    df_users = pd.read_sql(query_for_users, connection_string)
    df_users = df_users[df_users["role"] == 2]

    df_points = pd.read_sql(query_for_points, connection_string)
    df_task = pd.read_sql(query_for_task, connection_string)
    df_task_type = pd.read_sql(query_for_task_type, connection_string)

    df_points['uuid'] = df_points['uuid'].astype(str)
    df_task['agent_point_id'] = df_task['agent_point_id'].astype(str)
    points_task = pd.merge(df_points, df_task,
                           left_on='uuid', right_on='agent_point_id', how='outer', suffixes=('_x', ''))

    task_distributor = TaskDistributor(points_task, df_users, df_task, df_task_type,
                                       'your_yandex_api',
                                       'your_openroute_api')
    task_distributor.run_distribution()

    connection_string.dispose()
