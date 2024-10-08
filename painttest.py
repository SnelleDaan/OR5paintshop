import pandas as pd
import matplotlib.pyplot as plt

file = 'PaintShop - September 2024.xlsx'
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
    if Mx == 'M1':
        return 0
    elif Mx == 'M2':
        return 1
    elif Mx == 'M3':
        return 2
    elif Mx == 'Mx':
        return 3
        

def schedule_orders(orders, machines):
    schedule_O = []

    machine_states = {machine: {'current_color': None, 'available_time': 0} for machine in machines}

    machine_time = [0,0,0]
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


                
        # Schudule format = 'Order index,   machine,      end time,       colour,          duration,    start time
        schedule_O.append([order['Order'], best_machine, best_time, order['Colour'], best_paint_time, best_start_time])
        # Update the chosen machine state
        machine_time[machine_to_index(best_machine)] = best_time
        machine_states[best_machine]['available_time'] = best_time
        machine_states[best_machine]['current_color'] = order['Colour']
    return schedule_O

schedule1_O = schedule_orders(orders, machines)

def convert_sched_O_to_sched_M(schedule_O):
    schedule_M = [ [], [], []]
    for entry in schedule_O:
        order_index = entry[0]
        end_time = entry[2]
        color = entry[3].lower()
        duration = entry[4]
        start_time = entry[5]
        if entry[1] == 'M1':
            schedule_M[0].append((order_index, end_time, color, duration, start_time))
        elif entry[1] == 'M2':
            schedule_M[1].append((order_index, end_time, color, duration, start_time))
        elif entry[1] == 'M3':
            schedule_M[2].append((order_index, end_time, color, duration, start_time))
    return schedule_M

def calculate_penalty(orders, schedule):
    penalty = 0
    for order in orders:
        for entry in schedule:
            time_finish = entry[2]
            if order['Deadline'] < time_finish and entry[0]==order['Order']:
                penalty = penalty + order['Penalty'] * (time_finish - order['Deadline'])
    return penalty
calculate_penalty1 = calculate_penalty(orders, schedule1_O)

machine_schedules = convert_sched_O_to_sched_M(schedule1_O)

print(machine_schedules)

def draw_schedule(schedule):
    # schedule has to be of machince type
    # Create a figure and axis for plotting
    fig, ax = plt.subplots(figsize=(12, 6))

    # Loop through each machine's schedule
    for i, machine_schedule in enumerate(machine_schedules):
        for order_num, completion_time, color, duration, start_time in machine_schedule:
            # Calculate the width (duration) of each order
            width =  duration # The completion time is the total time spent on the order
            ax.barh(i + 1, width, left=start_time, color=color, edgecolor='black', label=color if i == 0 else "")
            ax.text(start_time + width / 2, i + 1, order_num, va='center', ha='center', color='white')  # Add order number to the bar

    # Add labels and title
    ax.set_yticks(range(1, len(machine_schedules) + 1))
    ax.set_yticklabels([f'Machine {i+1}' for i in range(len(machine_schedules))])
    ax.set_xlabel('Time')
    ax.set_title('Gantt Chart of Orders for Each Machine')

    # Show the plot
    plt.grid(axis='x', linestyle='--', alpha=0.7)
    plt.tight_layout()
    return plt.show()

draw_schedule(machine_schedules)