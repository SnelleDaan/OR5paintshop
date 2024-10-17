import pandas as pd
import matplotlib.pyplot as plt
import random

# Reading Excel data
file = 'PaintShop - November 2024.xlsx'
df = pd.read_excel(file, sheet_name=0)
orders  = df.to_dict(orient='records')

df2 = pd.read_excel(file, sheet_name=1)
machines = {}
for index, row in df2.iterrows():
    machines[row['Machine']] = {'speed': row['Speed']}

df3 = pd.read_excel(file, sheet_name=2)
kleurwissels = list(df3.itertuples(index=False, name=None))
setup_times = {}
for prev_color, current_color, interval in kleurwissels:
    color_pair = (prev_color, current_color)
    setup_times[color_pair] = interval


# Function defining
def painttime(area, machine, machines):
    for machine_name, machine_data in machines.items():
        if machine == machine_name:
            speed = machine_data['speed']
    return area / speed

def switchtime(prev_color, current_color):
    if prev_color is None:
        return 0
    if (prev_color, current_color) in setup_times:
        return setup_times[(prev_color, current_color)]
    return 0

def machine_to_index(Mx):
    if Mx == 'M1':
        return 0
    elif Mx == 'M2':
        return 1
    elif Mx == 'M3':
        return 2
    elif Mx == 'M4':
        return 3

def index_to_machine(index):
    return [f'M{index + 1}']


def schedule_orders(orders, machines):
    schedule_O = []
    machine_states = {machine: {'current_color': None, 'available_time': 0} for machine in machines}
    machine_time = [0, 0, 0, 0]

    for order in orders:
        best_machine = None
        best_time = float('inf')
        best_start_time = None
        
        for machine_name, machine_data in machines.items():
            current_machine = machine_states[machine_name]
            switch_time = switchtime(current_machine['current_color'], order['Colour'])
            paint_time = painttime(order['Surface'], machine_name, machines)
            time_to_complete = switch_time + paint_time
            start_time = machine_time[machine_to_index(machine_name)]
            completion_time = start_time + time_to_complete

            if completion_time < best_time:
                best_time = completion_time
                best_machine = machine_name
                best_start_time = start_time
                best_paint_time = paint_time
                best_switch_time = switch_time

        schedule_O.append([order['Order'], best_machine, best_time, order['Colour'].lower(), best_paint_time, best_start_time + best_switch_time, order['Deadline']])
        machine_time[machine_to_index(best_machine)] = best_time
        machine_states[best_machine]['available_time'] = best_time
        machine_states[best_machine]['current_color'] = order['Colour']
    
    return schedule_O


def convert_sched_O_to_sched_M(schedule_O):
    schedule_M = [[], [], [], []]
    for entry in schedule_O:
        order_index = entry[0]
        end_time = entry[2]
        color = entry[3].lower()
        duration = entry[4]
        start_time = entry[5]
        deadline = entry[6]
        schedule_M[machine_to_index(entry[1])].append((order_index, end_time, color, duration, start_time, deadline))
    return schedule_M


def calculate_penalty(orders, schedule):
    penalty = 0
    for order in orders:
        for entry in schedule:
            time_finish = entry[2]
            if order['Deadline'] < time_finish and entry[0] == order['Order']:
                penalty += order['Penalty'] * (time_finish - order['Deadline'])
    return penalty


def draw_schedule(schedule):
    fig, ax = plt.subplots(figsize=(12, 6))

    for i, machine_schedule in enumerate(schedule):
        for order_num, completion_time, color, duration, start_time, deadline in machine_schedule:
            width = duration
            ax.barh(i + 1, width, left=start_time, color=color, edgecolor='black', label=color if i == 0 else "")
            ax.text(start_time + width / 2, i + 1, order_num, va='center', ha='center', color='white')

    ax.set_yticks(range(1, len(schedule) + 1))
    ax.set_yticklabels([f'Machine {i+1}' for i in range(len(schedule))])
    ax.set_xlabel('Time')
    ax.set_title('Gantt Chart of Orders for Each Machine')

    plt.grid(axis='x', linestyle='--', alpha=0.7)
    plt.tight_layout()
    return plt.show()


# 2-opt Swap Improving Search
def swap_orders_optimization(orders, machines, max_iterations=1000):
    current_orders = orders.copy()
    current_schedule = schedule_orders(current_orders, machines)
    current_penalty = calculate_penalty(current_orders, current_schedule)
    improvement_list = []
    iteration_list = []
    improvement_index = 0
    count_iteration = 0

    for iteration in range(max_iterations):
        improved = False
        count_iteration += 1

        for i in range(len(current_orders)):
            for j in range(i + 1, len(current_orders)):
                current_orders[i], current_orders[j] = current_orders[j], current_orders[i]
                new_schedule = schedule_orders(current_orders, machines)
                new_penalty = calculate_penalty(current_orders, new_schedule)

                if new_penalty < current_penalty:
                    current_penalty = new_penalty
                    current_schedule = new_schedule
                    improvement_list.append(current_penalty)
                    improvement_index += 1
                    iteration_list.append(improvement_index)
                    improved = True
                else:
                    current_orders[i], current_orders[j] = current_orders[j], current_orders[i]
                    improvement_list.append(current_penalty)
                    improvement_index += 1
                    iteration_list.append(improvement_index)

        if not improved:
            break

    return current_schedule, current_penalty, improvement_list, iteration_list, count_iteration


# Tabu Search
def tabu_search_optimization(orders, machines, max_iterations=1000, tabu_tenure=10):
    current_orders = orders.copy()
    current_schedule = schedule_orders(current_orders, machines)
    current_penalty = calculate_penalty(current_orders, current_schedule)

    best_orders = current_orders.copy()
    best_schedule = current_schedule.copy()
    best_penalty = current_penalty

    tabu_list = []
    tabu_tenure_dict = {}
    tabu_length = tabu_tenure

    improvement_list = []
    iteration_list = []
    improvement_index = 0
    count_iteration = 0

    for iteration in range(max_iterations):
        improved = False
        count_iteration += 1
        best_move = None
        best_new_penalty = float('inf')

        for i in range(len(current_orders)):
            for j in range(i + 1, len(current_orders)):
                move = (current_orders[i]['Order'], current_orders[j]['Order'])

                if move in tabu_list and best_penalty <= current_penalty:
                    continue

                current_orders[i], current_orders[j] = current_orders[j], current_orders[i]
                new_schedule = schedule_orders(current_orders, machines)
                new_penalty = calculate_penalty(current_orders, new_schedule)

                if new_penalty < best_new_penalty:
                    best_new_penalty = new_penalty
                    best_move = (i, j)
                    improved = True

                current_orders[i], current_orders[j] = current_orders[j], current_orders[i]

        if best_move:
            i, j = best_move
            current_orders[i], current_orders[j] = current_orders[j], current_orders[i]
            current_schedule = schedule_orders(current_orders, machines)
            current_penalty = best_new_penalty

            tabu_list.append((current_orders[i]['Order'], current_orders[j]['Order']))
            tabu_tenure_dict[(current_orders[i]['Order'], current_orders[j]['Order'])] = iteration

            if len(tabu_list) > tabu_length:
                tabu_list.pop(0)

            if current_penalty < best_penalty:
                best_penalty = current_penalty
                best_orders = current_orders.copy()
                best_schedule = current_schedule.copy()

        improvement_list.append(current_penalty)
        improvement_index += 1
        iteration_list.append(improvement_index)

        tabu_list = [move for move in tabu_list if iteration - tabu_tenure_dict[move] < tabu_tenure]

        if not improved:
            break

    return best_schedule, best_penalty, improvement_list, iteration_list, count_iteration


# Run 2-opt swap
basic_schedule, basic_penalty, basic_improvement, basic_iteration, basic_count = swap_orders_optimization(orders, machines)

# Run Tabu Search
tabu_schedule, tabu_penalty, tabu_improvement, tabu_iteration, tabu_count = tabu_search_optimization(orders, machines)

# Gannt charts
draw_schedule(convert_sched_O_to_sched_M(basic_schedule))  # Basic optimization schedule
draw_schedule(convert_sched_O_to_sched_M(tabu_schedule))   # Tabu Search schedule

# Print penalties
print(f"Basic Swap Optimization Penalty: {basic_penalty}")
print(f"Tabu Search Optimization Penalty: {tabu_penalty}")

# Scatter plot to visualize improvement
plt.figure(figsize=(10, 6))

# 2-opt swap improvements
plt.plot(basic_iteration, basic_improvement, label='2-opt Swap Optimization', marker='o')

# Tabu Search improvements
plt.plot(tabu_iteration, tabu_improvement, label='Tabu Search Optimization', marker='x')

plt.title('Penalty Improvement over Iterations')
plt.xlabel('Iterations')
plt.ylabel('Penalty')
plt.legend()
plt.grid(True)
plt.show()

# Print iteration counts
# print(f"Total iterations for 2-opt Swap Optimization: {basic_count}")
# print(f"Total iterations for Tabu Search Optimization: {tabu_count}")