import pandas as pd

df = pd.read_excel('PaintShop - September 2024.xlsx', sheet_name=0)
orders  = df.to_dict(orient='records')

df2 = pd.read_excel('PaintShop - September 2024.xlsx', sheet_name=1)
machines = {}
for index, row in df2.iterrows():
    machines[row['Machine']] = {'speed': row['Speed']}

df3 = pd.read_excel('PaintShop - September 2024.xlsx', sheet_name=2)
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


def schedule_orders(orders, machines):
    schedule_O = []

    machine_states = {machine: {'current_color': None, 'available_time': 0} for machine in machines}
    
    for order in orders:
        best_machine = None
        best_time = float('inf')
        
        for machine_name, machine_data in machines.items():
            current_machine = machine_states[machine_name]
            switch_time = switchtime(current_machine['current_color'], order['Colour']) 
            paint_time = painttime(order['Surface'], machine_name, machines)
            time_to_complete = switch_time + paint_time
            completion_time = current_machine['available_time'] + time_to_complete
            
            if completion_time < best_time:
                best_time = completion_time
                best_machine = machine_name
        # Schudule format = 'Order index,   machine,    end time
        schedule_O.append([order['Order'], best_machine, completion_time])
        # Update the chosen machine state
        machine_states[best_machine]['available_time'] = best_time
        machine_states[best_machine]['current_color'] = order['Colour']
    return schedule_O

schedule1 = schedule_orders(orders, machines)

def convert_sched_O_to_sched_M(schedule_O):
    schedule_M = [ [], [], []]
    for entry in schedule_O:
        if entry[1] == 'M1':
            schedule_M[0].append()
        elif entry[1] == 'M2':
            schedule_M[1].append()
        elif entry[1] == 'M3':
            schedule_M[2].append()
    return schedule_M

def calculate_penalty(orders, schedule):
    penalty = 0
    for order in orders:
        for entry in schedule:
            time_finish = entry[2]
            if order['Deadline'] < time_finish and entry[0]==order['Order']:
                penalty = penalty + order['Penalty']
    return penalty
calculate_penalty1 = calculate_penalty(orders, schedule1)
print(schedule1, calculate_penalty1)