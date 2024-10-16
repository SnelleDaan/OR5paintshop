import pandas as pd
import matplotlib.pyplot as plt
import numpy as np

# Load data from Excel
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
    color_pair2 = (current_color, prev_color)
    setup_times[color_pair] = interval
    setup_times[color_pair2] = interval

# Calculate the paint time based on area and machine speed
def painttime(area, machine, machines):
    return area / machines[machine]['speed']

# Calculate the color switch time between previous and current color
def switchtime(prev_color, current_color):
    if prev_color is None:  # No switch time if the machine has not painted yet
        return 0
    if (prev_color, current_color) in setup_times:
        return setup_times[(prev_color, current_color)]
    return 0

# Function to schedule orders
def schedule_orders(orders, machines):
    schedule_O = []

    # Track each machine's state: current color and available time
    machine_states = {machine: {'current_color': None, 'available_time': 0} for machine in machines}

    for order in orders:
        best_machine = None
        best_time = float('inf')
        best_start_time = None
        
        # Iterate over each machine and evaluate the schedule
        for machine_name, machine_data in machines.items():
            current_machine = machine_states[machine_name]
            switch_time = switchtime(current_machine['current_color'], order['Colour']) 
            paint_time = painttime(order['Surface'], machine_name, machines)
            start_time = current_machine['available_time']  # Directly use the machine's available time
            completion_time = start_time + switch_time + paint_time
            
            # Choose the best machine based on the earliest completion time
            if completion_time < best_time:
                best_time = completion_time
                best_machine = machine_name
                best_start_time = start_time
                best_paint_time = paint_time
                best_switch_time = switch_time

        # Append the scheduled order details
        schedule_O.append([order['Order'], best_machine, best_time, order['Colour'], best_paint_time, best_start_time + best_switch_time])

        # Update the chosen machine's state (availability and current color)
        machine_states[best_machine]['available_time'] = best_time
        machine_states[best_machine]['current_color'] = order['Colour']

    return schedule_O

# Convert the schedule of orders to a schedule for machines
def convert_sched_O_to_sched_M(schedule_O):
    schedule_M = {machine: [] for machine in machines}
    for entry in schedule_O:
        order_index, machine_name, end_time, color, duration, start_time = entry
        schedule_M[machine_name].append((order_index, end_time, color.lower(), duration, start_time))
    return schedule_M

# Calculate penalties for late orders
def calculate_penalty(orders, schedule):
    penalty = 0
    for order in orders:
        for entry in schedule:
            time_finish = entry[2]
            if order['Deadline'] < time_finish and entry[0] == order['Order']:
                penalty += order['Penalty'] * (time_finish - order['Deadline'])
    return penalty

# Draw a Gantt chart of the schedule
def draw_schedule(machine_schedules):
    fig, ax = plt.subplots(figsize=(12, 6))

    # Loop through each machine's schedule
    for i, (machine_name, machine_schedule) in enumerate(machine_schedules.items()):
        for order_num, completion_time, color, duration, start_time in machine_schedule:
            width = duration  # The completion time is the total time spent on the order
            ax.barh(i + 1, width, left=start_time, color=color, edgecolor='black', label=color if i == 0 else "")
            ax.text(start_time + width / 2, i + 1, order_num, va='center', ha='center', color='white')

    ax.set_yticks(range(1, len(machine_schedules) + 1))
    ax.set_yticklabels([f'Machine {i + 1}' for i in range(len(machine_schedules))])
    ax.set_xlabel('Time')
    ax.set_title('Gantt Chart of Orders for Each Machine')

    plt.grid(axis='x', linestyle='--', alpha=0.7)
    plt.tight_layout()
    plt.show()

# Run the initial scheduling process
schedule1_O = schedule_orders(orders, machines)
machine_schedules = convert_sched_O_to_sched_M(schedule1_O)
draw_schedule(machine_schedules)

# Calculate and print the penalty for the initial schedule
penalty1 = calculate_penalty(orders, schedule1_O)
print(f"Initial penalty: {penalty1}")

# Function for 2-opt improving search heuristic
def swap_orders_optimization(orders, machines, max_iterations=1000):
    current_orders = orders.copy()
    current_schedule = schedule_orders(current_orders, machines)
    current_penalty = calculate_penalty(current_orders, current_schedule)

    for iteration in range(max_iterations):
        improved = False

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
                    improved = True
                else:
                    # Swap back if not improving
                    current_orders[i], current_orders[j] = current_orders[j], current_orders[i]

        # If no improvements found, stop
        if not improved:
            break

    return current_schedule, current_penalty

# Run the 2-opt optimization
optimized_schedule_2opt, optimized_penalty_2opt = swap_orders_optimization(orders, machines)

# Visualize and print the optimized schedule from 2-opt
machine_schedules_2opt = convert_sched_O_to_sched_M(optimized_schedule_2opt)
draw_schedule(machine_schedules_2opt)
print(f"Optimized penalty (2-opt): {optimized_penalty_2opt}")