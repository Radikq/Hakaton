import math
import time

import openpyxl
import pandas as pd
import requests
from yandex_geocoder import Client


class TaskDistributor:
    def __init__(self, points_file, employees_file, yandex_api_key, openroute_api_key):
        self.points = pd.read_excel(points_file, sheet_name='Входные данные для анализа', usecols=range(7))
        self.points.dropna(inplace=True)

        self.employees = pd.read_excel(employees_file, sheet_name='Справочник сотрудников')
        self.points_tomorrow = {
            key: [] for key in self.points.columns  # Создание датафрейма, с задачами которые не успели сегодня
        }
        self.points_tomorrow = pd.DataFrame(self.points_tomorrow)

        self.yandex_api_key = yandex_api_key
        self.openroute_api_key = openroute_api_key
        self.client = Client(yandex_api_key)
        self.token = openroute_api_key

        self.cnt = 0

    def set_task(self, point):
        """
        Установка для каждой задачи приоритета

        :param point: задача
        :return: приоритет: 1-высокий, 2-средний, 3-низкий
        """

        if point["Когда подключена точка?"] == "вчера" or point["Карты и материалы доставлены?"] == "нет":
            return 3
        elif point['Кол-во выданных карт'] / point['Кол-во одобренных заявок'] <= 0.5:
            return 2
        elif (point["Кол-во дней после выдачи последней карты"] >= 7 and point['Кол-во одобренных заявок'] >= 1) or \
                point[
                    "Кол-во дней после выдачи последней карты"] >= 14:
            return 1
        else:
            return 0

    def set_grade(self, employee_grade):
        """
        Установка каждому сотруднику списка приоритетов (для джуна [3] и так далее)

        :param employee_grade: сотрудник
        :return: список приоритетов, который сотрудник может выполнить
        """

        match employee_grade:
            case "Синьор":
                return [1, 2, 3]
            case "Мидл":
                return [2, 3]
            case "Джун":
                return [3]

    def matrix(self, address: list[str, str], profile=0):
        """
        Расчёт времени между адресами

        :param address: 2 адреса
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

    def distribute_tasks(self):
        """
        Самая главная функция, выполняющая распределение задач между сотрудниками

        :return: заполненный dataframe employees и points_tomorrow
        """

        for index, point in self.points.iterrows():
            task_number = point['№ Задачи']
            if task_number == 0:
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
            origin = 'Краснодар, ' + point['Адрес точки, г. Краснодар']

            suitable_employees = self.employees[self.employees['Решаемые задачи'].apply(lambda x: task_number in x)]
            suitable_employees = suitable_employees.sort_values(by=['Грейд'])

            if task_number == 3:
                junior = suitable_employees[suitable_employees["Грейд"] == "Джун"]
                if junior[junior['Кол-во отработанных минут'] >= 450]['ФИО'].count() < len(junior):
                    suitable_employees = junior
                else:
                    suitable_employees = suitable_employees[suitable_employees["Грейд"] != "Джун"]

            elif task_number == 2:
                middle = suitable_employees[suitable_employees["Грейд"] == "Мидл"]
                if middle[middle['Кол-во отработанных минут'] >= 420]['ФИО'].count() < len(middle) - 1:
                    suitable_employees = middle
                else:
                    suitable_employees = suitable_employees[suitable_employees["Грейд"] != "Джун"]

            for e_index, employee in suitable_employees.iterrows():

                if employee["Кол-во отработанных минут"] + task_time >= 540:
                    continue

                destination = employee["Адрес локации"]
                travel_time = self.matrix([origin, destination])

                if travel_time < best_travel_time:
                    best_travel_time = travel_time
                    best_employee_id = e_index

            if best_employee_id is None:
                if point["№ Задачи"] != 1:
                    point["№ Задачи"] -= 1
                self.points_tomorrow = pd.concat([self.points_tomorrow, pd.DataFrame([point])], ignore_index=True)
            else:
                self.employees.loc[best_employee_id, "Кол-во отработанных минут"] += best_travel_time + task_time
                self.employees.loc[best_employee_id, "Номера взятых задач"].append(index)
                self.employees.loc[best_employee_id, "Адрес локации"] = origin

    def run_distribution(self):
        self.points["№ Задачи"] = self.points.apply(self.set_task, axis=1)
        self.points.sort_values(by='№ Задачи', inplace=True)

        self.employees["Решаемые задачи"] = self.employees["Грейд"].apply(self.set_grade)
        self.employees["Кол-во отработанных минут"] = 0
        self.employees["Номера взятых задач"] = self.employees["Кол-во отработанных минут"].apply(lambda x: [])

        self.distribute_tasks()

        self.points_tomorrow.to_csv('points_tomorrow.csv', index=False)
        self.employees[['ФИО', "Номера взятых задач"]].to_csv('employees.csv', index=False)


if __name__ == "__main__":
    task_distributor = TaskDistributor('DataSet.xlsx', 'DataSet.xlsx', 'your_yandex_api_key', 'your_openroute_api_key')
    task_distributor.run_distribution()
