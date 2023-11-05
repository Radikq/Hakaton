import math
import pandas as pd
import requests
from yandex_geocoder import Client

points = pd.read_excel('DataSet.xlsx', sheet_name='Входные данные для анализа', usecols=range(7))
points.dropna(inplace=True)


def set_task(point):
    if point["Когда подключена точка?"] == "вчера" or point["Карты и материалы доставлены?"] == "нет":
        return 3
    elif point['Кол-во выданных карт'] / point['Кол-во одобренных заявок'] <= 0.5:
        return 2
    elif (point["Кол-во дней после выдачи последней карты"] >= 7 and point['Кол-во одобренных заявок'] >= 1) or point[
        "Кол-во дней после выдачи последней карты"] >= 14:
        return 1
    else:
        return 0


points["№ Задачи"] = points.apply(set_task, axis=1)
points.sort_values(by='№ Задачи', inplace=True)

client = Client("")     #Yandex api key
token = ''              #OpenRouteService Api Key


def matrix(address: list[str, str], profile=0):
    coord_1 = client.coordinates(address[0])
    coord_2 = client.coordinates(address[1])

    locations = [[float(coord_1[1]), float(coord_1[0])], [float(coord_2[1]), float(coord_2[0])]]

    headers = {
        'Content-Type': 'application/json; charset=utf-8',
        'Accept': 'application/json',
        'Authorization': token
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

    return math.ceil(ps["durations"] / 60)


employees = pd.read_excel('DataSet.xlsx', sheet_name='Справочник сотрудников')
centres = employees["Адрес локации"].unique()


def set_grade(employee_grade):
    match employee_grade:
        case "Синьор":
            return [1, 2, 3]
        case "Мидл":
            return [2, 3]
        case "Джун":
            return [3]


employees["Решаемые задачи"] = employees["Грейд"].apply(set_grade)
employees["Кол-во отработанных минут"] = 0
employees["Номера взятых задач"] = employees["Кол-во отработанных минут"].apply(lambda x: [])

points_tomorrow = {
    key: [] for key in points.columns
}
points_tomorrow = pd.DataFrame(points_tomorrow)


def distribute_tasks():
    global points_tomorrow
    for index, point in points.iterrows():
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

        suitable_employees = employees[employees['Решаемые задачи'].apply(lambda x: task_number in x)]
        suitable_employees = suitable_employees.sort_values(by=['Грейд'])

        if task_number == 3:
            junior = suitable_employees[suitable_employees["Грейд"] == "Джун"]
            if junior[junior['Кол-во отработанных минут'] >= 480]['ФИО'].sum() < len(junior):
                suitable_employees = junior
            else:
                suitable_employees = suitable_employees[suitable_employees["Грейд"] != "Джун"]

        elif task_number == 2:
            middle = suitable_employees[suitable_employees["Грейд"] == "Мидл"]
            if middle[middle['Кол-во отработанных минут'] >= 480]['ФИО'].sum() < len(middle):
                suitable_employees = middle
            else:
                suitable_employees = suitable_employees[suitable_employees["Грейд"] == "Синьор"]

        for e_index, employee in suitable_employees.iterrows():

            if employee["Кол-во отработанных минут"] + task_time >= 480:
                continue

            destination = employee["Адрес локации"]
            travel_time = matrix([origin, destination])

            if travel_time < best_travel_time:
                best_travel_time = travel_time
                best_employee_id = e_index

        if best_employee_id is None:
            if point["№ Задачи"] != 1:
                point["№ Задачи"] -= 1
            points_tomorrow = pd.concat([points_tomorrow, pd.DataFrame([point])], ignore_index=True)
        else:
            employees.loc[best_employee_id, "Кол-во отработанных минут"] += best_travel_time + task_time
            employees.loc[best_employee_id, "Номера взятых задач"].append(index)
            employees.loc[best_employee_id, "Адрес локации"] = origin