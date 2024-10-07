'''
define example variables
'''

colors = ['red','yellow','blue']

kleurwissels = [
    ("red", "yellow", 2),
    ("yellow", "blue", 3),
    ("blue", "red", 1),
]
setup_times = {}
for prev_color, current_color, interval in kleurwissels:
    color_pair = (prev_color, current_color)
    color_pair2 = (current_color, prev_color)
    if color_pair in setup_times:
        setup_times[color_pair].append(interval)
        setup_times[color_pair2].append(interval)
    else:
        setup_times[color_pair] = interval
        setup_times[color_pair2] = interval


# Order, area, Color, Deadline, Penalty
orders = [
    {'order': 1,'area': 100, 'color': 'red'},   # Order 1: 100 m² of red paint
    {'order': 2,'area': 150, 'color': 'blue'},  # Order 2: 150 m² of blue paint
    {'order': 3,'area': 80, 'color': 'red'}     # Order 3: 80 m² of red paint
]


machines = {
    'Machine1': {'speed': 10},  # m²/hour
    'Machine2': {'speed': 15},
    'Machine3': {'speed': 20}
}


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
'''
visualisize the planning
'''


def schedule_orders(orders, machines):
    schedule = []
    machine_states = {machine: {'current_color': None, 'available_time': 0} for machine in machines}
    
    for order in orders:
        best_machine = None
        best_time = float('inf')
        
        for machine_name, machine_data in machines.items():
            current_machine = machine_states[machine_name]
            switch_time = switchtime(current_machine['current_color'], order['color']) 
            paint_time = painttime(order['area'], machine_name, machines)
            time_to_complete = switch_time + paint_time
            completion_time = current_machine['available_time'] + time_to_complete
            
            if completion_time < best_time:
                best_time = completion_time
                best_machine = machine_name
        # Schudule format = 'Order index,   machine,    start time
        schedule.append([order['order'], best_machine, machine_states[best_machine]['available_time']])
        # Update the chosen machine state
        machine_states[best_machine]['available_time'] = best_time
        machine_states[best_machine]['current_color'] = order['color']
    
    return schedule

print(schedule_orders(orders, machines))
