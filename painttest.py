import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import random

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


def painttime(area, machine, machines):
    for machine_name, machine_data in machines.items():
        if machine == machine_name:
            speed = machine_data['speed']
    return area/ speed

# Calculate the color switch time between previous and current color
def switchtime(prev_color, current_color):
    if prev_color is None:  # No switch time if the machine has not painted yet
        return 0
    if (prev_color, current_color) in setup_times:
        return setup_times[(prev_color, current_color)]
    return 0

def machine_to_index(Mx):
    if Mx.startswith('M') and Mx[1:].isdigit():
        return int(Mx[1:]) - 1
    else:
        raise ValueError("Invalid input format. Expected 'M' followed by a number.")

def index_to_machine(index):
    return [f'M{index + 1}']


def schedule_orders(orders, machines):
    schedule_O = []

    machine_states = {machine: {'current_color': None, 'available_time': 0} for machine in machines}
    machine_time = []
    for machine in machines:
        machine_time.append(0)
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


                
        # Schudule format = 'Order index,   machine,      end time,       colour,          duration,    start time
        schedule_O.append([order['Order'], best_machine, best_time, order['Colour'].lower(), best_paint_time, best_start_time+best_switch_time, order['Deadline']])
        # Update the chosen machine state
        machine_time[machine_to_index(best_machine)] = best_time
        machine_states[best_machine]['available_time'] = best_time
        machine_states[best_machine]['current_color'] = order['Colour']
    return schedule_O


def convert_sched_O_to_sched_M(schedule_O):
    schedule_M = []
    for i in range(len(machines)):
        schedule_M.append([])
    for entry in schedule_O:
        order_index = entry[0]
        end_time = entry[2]
        color = entry[3].lower()
        duration = entry[4]
        start_time = entry[5]
        deadline = entry[6]
        schedule_M[machine_to_index(entry[1])].append((order_index, end_time, color, duration, start_time, deadline))
    return schedule_M

def convert_sched_M_to_sched_O(schedule_M):
    schedule_O = []
    for i, entry in enumerate(schedule_M):
        for order_i, end_time, color, duration, start_time, deadline in entry:
            schedule_O.append([order_i, index_to_machine(i)[0], end_time, color, duration, start_time, deadline])
    schedule_O.sort(key=lambda x: int(x[0][1:]))
    return schedule_O

schedule1_O = schedule_orders(orders, machines)
    
def calculate_penalty(orders, schedule):
    penalty = 0
    penalty_list = []
    for order in orders:
        for entry in schedule:
            time_finish = entry[2]
            if order['Deadline'] < time_finish and entry[0]==order['Order']:
                penalty = penalty + order['Penalty'] * (time_finish - order['Deadline'])
                penalty_list.append((order['Penalty'] * (time_finish - order['Deadline']),order['Order']))
    return penalty
penalty1 = calculate_penalty(orders, schedule1_O)

machine_schedules = convert_sched_O_to_sched_M(schedule1_O)


def draw_schedule(schedule):
    # schedule has to be of machince type
    # Create a figure and axis for plotting
    fig, ax = plt.subplots(figsize=(12, 6))
    
    # Loop through each machine's schedule
    for i, machine_schedule in enumerate(schedule):
        previous_end_time = 0  # Track end time of previous order
        
        for order_num, completion_time, color, duration, start_time,deadline in machine_schedule:
            
            if start_time > previous_end_time:
                gap = start_time - previous_end_time
                ax.barh(i + 1, gap, left=previous_end_time, color='grey', edgecolor='black')
            # Calculate the width (duration) of each order
            width =  duration # The completion time is the total time spent on the order
            ax.barh(i + 1, width, left=start_time, color=color, edgecolor='black', label=color if i == 0 else "")
            ax.text(start_time + width / 2, i + 1, order_num, va='center', ha='center', color='white')  # Add order number to the bar
            
            previous_end_time = start_time + duration

    # Add labels and title
    ax.set_yticks(range(1, len(schedule) + 1))
    ax.set_yticklabels([f'Machine {i+1}' for i in range(len(schedule))])
    ax.set_xlabel('Time')
    ax.set_title('Gantt Chart of Orders for Each Machine')

    # Show the plot
    plt.grid(axis='x', linestyle='--', alpha=0.7)
    plt.tight_layout()
    return plt.show()


def swap_orders_optimization(orders, machines, max_iterations=1000):
    current_orders = orders.copy()
    current_schedule = schedule_orders(current_orders, machines)
    current_penalty = calculate_penalty(current_orders, current_schedule)
    improvement_list= []
    iteration_list = []
    improvement_index = 0
    count_iteration = 0
    for iteration in range(max_iterations):
        improved = False
        count_iteration += 1
        # Try swapping each pair of orders
        for i in range(len(current_orders)):
            for j in range(i + 1, len(current_orders)):
                # Swap orders
                current_orders[i], current_orders[j] = current_orders[j], current_orders[i]

                # Calculate new schedule and penalty
                new_schedule = schedule_orders(current_orders, machines)
                new_penalty = calculate_penalty(current_orders, new_schedule)

                # Accept the swap if it improves the penalty
                if new_penalty < current_penalty:
                    current_penalty = new_penalty
                    current_schedule = new_schedule
                    improvement_list.append(current_penalty)
                    improvement_index += 1
                    iteration_list.append(improvement_index)
                    improved = True
                else:
                    # Swap back if not improving
                    current_orders[i], current_orders[j] = current_orders[j], current_orders[i]
                    improvement_list.append(current_penalty)
                    improvement_index += 1
                    iteration_list.append(improvement_index)

        # If no improvements found in an iteration, we can stop
        if not improved:
            break

    return current_schedule, current_penalty, improvement_list, iteration_list, count_iteration

def random_schedule(orders):
    
    random_orders = random.sample(orders, len(orders))
    return random_orders
# Run the swap orders optimization
optimized_schedule, optimized_penalty, list_of_improvement, list_iteration, count_it = swap_orders_optimization(orders, machines)
random_schedule(orders)
draw_schedule(machine_schedules)
draw_schedule(convert_sched_O_to_sched_M(optimized_schedule))
print(f"Optimized penalty achieved: {optimized_penalty}")
print(penalty1, optimized_penalty, count_it)
# print(min(list_of_improvement))
plt.scatter(list_iteration, list_of_improvement, marker='o')
plt.title('Improvement per iteration')
plt.xlabel('iteration')
plt.ylabel('penalty')
plt.show()