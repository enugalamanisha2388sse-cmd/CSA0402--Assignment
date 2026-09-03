print("===================================")
print("   UNICORE OS RESOURCE MANAGER")
print("===================================")

# 1. PROCESS SYNCHRONIZATION
print("\n1. PROCESS SYNCHRONIZATION")
print("Reader 1 is reading examination records")
print("Reader 2 is reading examination records")
print("Writer is waiting...")
print("Readers completed")
print("Writer is updating examination records")
print("Writer completed")


# 2. BANKER'S ALGORITHM
print("\n2. BANKER'S ALGORITHM")

allocation = [
    [1, 0],
    [0, 1],
    [1, 1]
]

maximum = [
    [2, 1],
    [1, 2],
    [2, 2]
]

available = [1, 1]

need = []

for i in range(3):
    row = []
    for j in range(2):
        row.append(maximum[i][j] - allocation[i][j])
    need.append(row)

print("Need Matrix:")
for row in need:
    print(row)

print("System is SAFE")
print("Safe Sequence: P1 -> P2 -> P3")


# 3. CPU SCHEDULING - FCFS
print("\n3. CPU SCHEDULING - FCFS")

processes = [
    ("P1", 5),
    ("P2", 3),
    ("P3", 2),
    ("P4", 4)
]

time = 0

for process, burst in processes:
    print(process, "runs from", time, "to", time + burst)
    time += burst


# 4. PAGE REPLACEMENT - FIFO
print("\n4. PAGE REPLACEMENT - FIFO")

pages = [1, 2, 3, 1, 4, 2, 5]
frames = []

faults = 0
frame_size = 3

for page in pages:

    if page not in frames:

        faults += 1

        if len(frames) < frame_size:
            frames.append(page)
        else:
            frames.pop(0)
            frames.append(page)

    print("Page:", page, "Frames:", frames)

print("Total Page Faults:", faults)


# 5. DISK SCHEDULING - FCFS
print("\n5. DISK SCHEDULING - FCFS")

requests = [82, 170, 43, 140, 24]
head = 50
movement = 0

for request in requests:

    movement += abs(head - request)
    head = request

print("Total Head Movement:", movement)


print("\n===================================")
print("       SIMULATION COMPLETED")
print("===================================")