# This script matches students to supervisors using their preferences.
# It uses linear programming (PuLP) to maximize an overall satisfaction score.
# It uses data from an Excel file with two sheets: supervisors and students. It writes data in a third sheet: matches
# Every student and every supervisor has a list of supervision areas in order of preference.
# The goal is to assign students to supervisors in a way that maximises overall satisfaction based on preferences and workload constraints.

import pandas as pd
from pulp import LpProblem, LpVariable, LpMaximize, lpSum, LpBinary

# This function removes numbers from a string, which is useful for cleaning up names or areas.
# This useful because the lists of areas from students and supervisors are numbered differently.
import re

def remove_numbers(text):
    if isinstance(text, str):
        return re.sub(r'\d+', '', text).strip()
    return text

def normalize_area(text):
    if isinstance(text, str):
        # Remove all numbers, whitespace, dots, and commas, then lowercase
        return re.sub(r'[\s\d\.,]+', '', text).lower()
    return text

# 1. Load data from Excel
file = file = "dissmatch_data.xlsx" # Insert here the name or path to the Excel file
supervisors = pd.read_excel(file, sheet_name=0)
students = pd.read_excel(file, sheet_name=1)

# Display the dataframes to check if they are loaded correctly
#print(supervisors)
#print(students)

# 2. Prepare data
supervisor_names = supervisors['name'].tolist()
student_names = students['name'].tolist()
workload = dict(zip(supervisors['name'], supervisors['workload']))

# Normalize all area strings for comparison
raw_areas = supervisors[['area1','area2','area3','area4','area5']].values.flatten()
areas = list(set(normalize_area(a) for a in raw_areas if pd.notnull(a)))

# Extraction of preferences for supervisors (normalized)
supervisor_area_cols = [f'area{i}' for i in range(1, 6)]
supervisor_preferences = {
    row['name']: [normalize_area(row[col]) for col in supervisor_area_cols if pd.notnull(row[col])]
    for _, row in supervisors.iterrows()
}

# Extraction of preferences for students (normalized)
student_area_cols = [col for col in students.columns if col.startswith('proposal_')]
student_preferences = {
    row['name']: [normalize_area(row[col]) for col in student_area_cols if pd.notnull(row[col])]
    for _, row in students.iterrows()
}

#The prepared data can be displayed here
#print("Supervisors:", supervisor_names)
#print("Students:", student_names)
#print("Workload:", workload)
#print("Areas:", areas)

#The preferences can be displayed here (names can be changed to match the data)
#print(supervisor_preferences["Prof. Smith"])
#print(student_preferences["Alice"])

# 3. Set up the problem
prob = LpProblem("Supervision_Allocation", LpMaximize)

# 4. Create decision variables for shared areas: x[student, supervisor, area]. This section creates a list of all student/supervisor matches for each area.  
x = {}
for student in student_names:
    for supervisor in supervisor_names:
        for area in areas:
            x[(student, supervisor, area)] = LpVariable(f"x_{student}_{supervisor}_{area}", 0, 1, LpBinary)

# 5. Define satisfaction scores.
# You can tinker with all the return values - these are satisfaction points that contribute to the overall score.
# The higher the score, the more the code will try to create matches to obtain that score i.e. if choice 1 is assigned, it will get 10 points, if choice 2 is assigned, it will get 7 points, etc.
# If the student values are overall higher than supervisor values, then student choices are considered more important. 
def student_points(student, area):
    prefs = student_preferences[student]
    if len(prefs) > 0 and area == prefs[0]: return 10 # Highest satisfaction for first preference
    if len(prefs) > 1 and area == prefs[1]: return 7
    if len(prefs) > 2 and area == prefs[2]: return 5
    return -10  # Penalize non-matching areas. Don't change this value, as it is used to eliminate non-matching areas.

def supervisor_points(supervisor, area):
    prefs = supervisor_preferences[supervisor]
    if len(prefs) > 0 and area == prefs[0]: return 5 # Highest satisfaction for first preference
    if len(prefs) > 1 and area == prefs[1]: return 4
    if len(prefs) > 2 and area == prefs[2]: return 3
    if len(prefs) > 3 and area == prefs[3]: return 2
    if len(prefs) > 4 and area == prefs[4]: return 1
    return -10  # Penalize non-matching areas. Don't change this value, as it is used to eliminate non-matching areas.

# Build satisfaction dictionary
satisfaction = {}
for student in student_names:
    for supervisor in supervisor_names:
        for area in areas:
            satisfaction[(student, supervisor, area)] = student_points(student, area) + supervisor_points(supervisor, area)

# 6. Objective: maximize total satisfaction
prob += lpSum(
    satisfaction[(student, supervisor, area)] * x[(student, supervisor, area)]
    for student in student_names
    for supervisor in supervisor_names
    for area in areas
)

# Check if the satisfaction scores are calculated correctly. 
# print("Satisfaction score:",satisfaction[("Bob", "Prof. Lee", "Human Rights Law")]) 

# 7. Constraint - Each student can only be assigned to one supervisor and one area
for student in student_names:
    prob += lpSum(
        x[(student, supervisor, area)]
        for supervisor in supervisor_names
        for area in areas
    ) <= 1 # Each student must be assigned one supervisor and one area. 
# If no match can be made, the asnwer will show the unmatched students.

# 8. Constraint - No supervisor exceeds their workload
for supervisor in supervisor_names:
    prob += lpSum(
        x[(student, supervisor, area)]
        for student in student_names
        for area in areas
    ) <= workload[supervisor] # Each supervisor's total assignments must not exceed their workload. 
# If there are more students than supervision workload, the answer will show the supervisors with remaining workload.

# 9. Constraint - Only allow matches if area is in both preferences
for student in student_names:
    for supervisor in supervisor_names:
        for area in areas:
            if satisfaction[(student, supervisor, area)] == -20:
                prob += x[(student, supervisor, area)] == 0 # If the area is not in both preferences, the match is not allowed.

# 10. Solve the problem
prob.solve()

# 11. Check the solution status
from pulp import LpStatus
print("Status:", LpStatus[prob.status])

# 12. Display the assignments
for student in student_names:
    for supervisor in supervisor_names:
        for area in areas:
            val = x[(student, supervisor, area)].varValue
            if val is not None and val > 0.5:
                print(f"{student} assigned to {supervisor} for area {area}")


# 13. Export matches to Excel ("matches" sheet) with student choice number
results = []
for student in student_names:
    prefs = student_preferences[student]
    for supervisor in supervisor_names:
        for area in areas:
            val = x[(student, supervisor, area)].varValue
            if val is not None and val > 0.5:
                # Determine student's choice number for this area
                if len(prefs) > 0 and area == prefs[0]:
                    choice = 1
                elif len(prefs) > 1 and area == prefs[1]:
                    choice = 2
                elif len(prefs) > 2 and area == prefs[2]:
                    choice = 3
                else:
                    choice = None
                results.append({
                    'student': student,
                    'supervisor': supervisor,
                    'area': area,
                    'student_choice': choice
                })

results_df = pd.DataFrame(results)
with pd.ExcelWriter(file, mode='a', if_sheet_exists='replace') as writer:
    results_df.to_excel(writer, sheet_name='matches', index=False)


# 14. Test the solution

# Total satisfaction score
total_satisfaction = sum(
    satisfaction[(student, supervisor, area)]
    for student in student_names
    for supervisor in supervisor_names
    for area in areas
    if x[(student, supervisor, area)].varValue is not None and x[(student, supervisor, area)].varValue > 0.5
)
print("Total satisfaction score:", total_satisfaction)

# Supervisor workload check
for supervisor in supervisor_names:
    count = sum(
        x[(student, supervisor, area)].varValue is not None and x[(student, supervisor, area)].varValue > 0.5
        for student in student_names
        for area in areas
    )
    print(f"{supervisor}: assigned {count} students (workload limit: {workload[supervisor]})")

# Student preference satisfaction check
first, second, third, other = 0, 0, 0, 0
for student in student_names:
    for supervisor in supervisor_names:
        for area in areas:
            if x[(student, supervisor, area)].varValue is not None and x[(student, supervisor, area)].varValue > 0.5:
                prefs = student_preferences[student]
                if len(prefs) > 0 and area == prefs[0]:
                    first += 1
                elif len(prefs) > 1 and area == prefs[1]:
                    second += 1
                elif len(prefs) > 2 and area == prefs[2]:
                    third += 1
                else:
                    other += 1
print(f"First student choice: {first}, Second: {second}, Third: {third}, Other: {other}")

# Supervisor preference satisfaction check
first, second, third, fourth, fifth, other = 0, 0, 0, 0, 0, 0
for student in student_names:
    for supervisor in supervisor_names:
        for area in areas:
            if x[(student, supervisor, area)].varValue is not None and x[(student, supervisor, area)].varValue > 0.5:
                prefs = supervisor_preferences[supervisor]
                if len(prefs) > 0 and area == prefs[0]:
                    first += 1
                elif len(prefs) > 1 and area == prefs[1]:
                    second += 1
                elif len(prefs) > 2 and area == prefs[2]:
                    third += 1
                elif len(prefs) > 3 and area == prefs[3]:
                    fourth += 1
                elif len(prefs) > 4 and area == prefs[4]:
                    fifth += 1
                else:
                    other += 1
print(f"First supervisor choice: {first}, Second: {second}, Third: {third}, Fourth: {fourth}, Fifth: {fifth}, Other: {other}")

# All students matched?
assigned_students = set()
for student in student_names:
    for supervisor in supervisor_names:
        for area in areas:
            if x[(student, supervisor, area)].varValue is not None and x[(student, supervisor, area)].varValue > 0.5:
                assigned_students.add(student)
print("All students matched:", len(assigned_students) == len(student_names))

# Identify unmatched students
unmatched = []
for student in student_names:
    assigned = False
    for supervisor in supervisor_names:
        for area in areas:
            val = x[(student, supervisor, area)].varValue
            if val is not None and val > 0.5:
                assigned = True
    if not assigned:
        unmatched.append(student)
print("Unmatched students due to workload or other constraints:", unmatched)

# Identify supervisors with remaining capacity
supervisors_with_capacity = []
for supervisor in supervisor_names:
    assigned_count = sum(
        x[(student, supervisor, area)].varValue is not None and x[(student, supervisor, area)].varValue > 0.5
        for student in student_names
        for area in areas
    )
    if assigned_count < workload[supervisor]:
        supervisors_with_capacity.append((supervisor, workload[supervisor] - assigned_count))
        print(f"{supervisor} has {workload[supervisor] - assigned_count} slots remaining.")

if not supervisors_with_capacity:
    print("All supervisors are at full capacity.")