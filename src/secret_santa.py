import pandas as pd
import random

employee_data = pd.read_excel('../data/Employee-List.xlsx', sheet_name='Employee-List')
last_year_data = pd.read_excel('../data/Secret-Santa-Game-Result-2023.xlsx', sheet_name='Secret-Santa-Game-Result-2023')

employees = employee_data[['Employee_Name', 'Employee_EmailID']].to_dict('records')
last_year_pairs = last_year_data[['Employee_Name', 'Secret_Child_Name']].to_dict('records')

def generate_secret_santa_pairs(employees, last_year_pairs):
    last_year_lookup = {pair['Employee_Name']: pair['Secret_Child_Name'] for pair in last_year_pairs}
    
    shuffled_employees = employees.copy()
    random.shuffle(shuffled_employees)
    
    assignments = []
    for employee in employees:
        available_choices = [e for e in shuffled_employees if e['Employee_Name'] != employee['Employee_Name']
                             and e['Employee_Name'] != last_year_lookup.get(employee['Employee_Name'])]
        
        if not available_choices:
            raise ValueError("Unable to assign a unique secret child for each employee.")
        
        secret_child = available_choices.pop(0)
        shuffled_employees.remove(secret_child)
        
        assignments.append({
            'Employee_Name': employee['Employee_Name'],
            'Employee_EmailID': employee['Employee_EmailID'],
            'Secret_Child_Name': secret_child['Employee_Name'],
            'Secret_Child_EmailID': secret_child['Employee_EmailID']
        })
    
    return pd.DataFrame(assignments)

result_df = generate_secret_santa_pairs(employees, last_year_pairs)
result_df.to_csv('../data/Secret_Santa_Assignments.csv', index=False)
print("Secret Santa assignments saved to '../data/Secret_Santa_Assignments.csv'")
