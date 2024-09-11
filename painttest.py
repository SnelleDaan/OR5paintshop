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
        setup_times[color_pair] = [interval]
        setup_times[color_pair2] = [interval]

# Order, Surface, Color, Deadline, Penalty
orders = [(1, 10, 'yellow', 18, 10), (2, 5, 'red', 10, 2), (3, 15, 'blue', 20, 8)]

machines = ['M1','M2']

machine_speed = {
    'M1': 10,
    'M2': 15
}

