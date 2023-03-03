import random
import pandas as pd


def get_student_priority(record: dict):
    record = {key: value for key, value in record.items() if key != 'Name'}
    priority = dict(sorted(record.items(), key=lambda item: item[1]))
    res = list(priority.keys())
    return res


def allocate_site(capacity: dict, priority_order: list):
    for priority in priority_order:
        if capacity.get(priority) != 0:
            return priority


with open('students.csv') as stu:
    students = stu.read().split("\n")

random.shuffle(students)
ranking_df = pd.read_csv('internship_ranking.csv')
capacity: dict = pd.read_csv('intership_capacity.csv').to_dict('records')[0]
internship_sites = list(capacity.keys())
internship_allocation = []

for student in students:
    df = ranking_df.loc[ranking_df['Name'].str.strip().str.lower() == student.lower().strip()]
    record = df.to_dict('records')[0]
    priority = get_student_priority(record)
    allocated_site = allocate_site(capacity, priority)
    capacity.update({allocated_site: capacity.get(allocated_site) - 1})
    print(student, allocated_site)
