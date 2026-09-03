# UniCore: OS Resource Orchestration System
### Design, Mathematical Analysis, and Simulated Implementation for a Smart University
**Course:** CSA04 – Operating Systems (Slot C)  
**Target Outcomes:** CO2 (Concurrency & Deadlocks), CO3 (Virtual Memory), CO4 (File Systems & I/O)  
**Faculty In-charge:** Dr. J. Priskilla Angel Rani  
**Submission Document:** [CSA04_OS_Assignment_SLOT_C_Complete_Submission.pdf](CSA04_OS_Assignment_SLOT_C_Complete_Submission.pdf)

---

## 📌 Executive Summary

**UniCore** is an enterprise-grade academic computing platform modeled to coordinate shared hardware resources across the heterogeneous digital workloads of a modern smart university. Under baseline operating conditions, the system supports ~500 concurrent sessions, scaling beyond 800 during peak examination, registration, and result-disclosure windows.

This repository contains the complete algorithmic implementations, mathematical proofs, simulation engines, graphical trajectory visualizations, and the full comprehensive project report in PDF format.

---

## 📁 Repository Structure

```
.
├── CSA04_OS_Assignment_SLOT_C_Complete_Submission.pdf  # Complete Final Assignment Report
├── full_simulator.py                                   # Master OS Subsystem Simulator Engine
├── simulator.py                                        # Core simulation logic & benchmarking
├── draw_architecture.py                                # Architectural diagram generator
├── build_full_submission.py                            # Report builder & documentation compiler
├── generate_docx.py                                    # Word document generation utility
├── screenshots/                                        # Visual artifacts, charts, and terminal output captures
│   ├── chart_architecture_flow.png                    # End-to-end subsystem architectural flow
│   ├── chart_cpu_gantt.png                             # Execution Gantt charts across schedulers
│   ├── chart_disk_trajectories.png                     # Multi-algorithm disk head travel curves
│   ├── terminal_cpu.png                                # CPU scheduling execution log
│   ├── terminal_bankers.png                            # Banker's safety state & matrix verification
│   ├── terminal_memory.png                             # Page replacement trace & fault rate
│   ├── terminal_filealloc.png                          # Storage layout & fragmentation benchmark
│   └── terminal_disk.png                               # Disk scheduling comparison metrics
├── cpu_output.png                                      # Root visual output for CPU scheduler
├── bankers_output.png                                  # Root visual output for Deadlock avoidance
├── memory_output.png                                   # Root visual output for Memory management
├── disk_output.png                                     # Root visual output for Disk scheduling
├── .gitignore                                          # Git ignore rules
└── README.md                                           # Repository overview and documentation
```

---

## ⚙️ Implemented Subsystems & Algorithms

### 1. CPU Scheduling (CO2)
- **Algorithms:** First-Come First-Served (FCFS), Shortest Job First (SJF Preemptive / Non-Preemptive), Priority Scheduling (Preemptive / Non-Preemptive), and Round Robin (RR with configurable time quantum $q$).
- **Metrics Evaluated:** Completion Time, Turnaround Time ($TAT = CT - AT$), Waiting Time ($WT = TAT - BT$), and Response Time ($RT$).
- **Artifacts:** Execution logs and multi-queue Gantt timelines.

### 2. Process Synchronization & Deadlock Avoidance (CO2)
- **Algorithm:** Banker's Algorithm for Multi-Resource Safety and Request Allocation.
- **Resource Classes Modeled:**
  - $R_0$: CPU Cores (5 units)
  - $R_1$: Database Locks (17 units)
  - $R_2$: Shared Memory Buffers (12 units)
  - $R_3$: Network Sockets (9 units)
- **Safety State Verification:** Validates $\text{Need}[i][j] \le \text{Available}[j]$ iteratively to establish a starvation-free safe execution sequence: `<P1, P3, P4, P0, P2>`.

### 3. Virtual Memory & Page Replacement (CO3)
- **Architectural Specs:** 16 GB Physical RAM, 4 KB Page Size, 64 MB Logical Address Space, 4-byte Page Table Entry (PTE).
- **Replacement Algorithms:** First-In First-Out (FIFO), Least Recently Used (LRU), and Belady's Optimal (OPT).
- **Performance Evaluation:** Page hit/fault rates and effective memory access time ($EMAT$) under Translation Lookaside Buffer (TLB) hits and misses.

### 4. Storage & File Allocation (CO4)
- **Allocation Techniques:**
  - Contiguous Allocation
  - Linked List Allocation
  - Indexed Allocation (single-level and multi-level index blocks)
- **Evaluation:** Internal vs. external fragmentation, seek overhead, and random vs. sequential file access efficiency.

### 5. Secondary Storage & Disk Arm Scheduling (CO4)
- **Algorithms:** FCFS, Shortest Seek Time First (SSTF), SCAN (Elevator), Circular SCAN (C-SCAN), LOOK, and Circular LOOK (C-LOOK).
- **Benchmark Geometry:** 200-cylinder disk ($0 - 199$), starting position at cylinder $53$.
- **Pending Cylinders:** `[98, 183, 37, 122, 14, 124, 65, 67, 190, 55]`.
- **Output:** Total head movement, average seek distance, and head movement trajectories.

---

## 🚀 Execution & Usage

### Prerequisites
Ensure Python 3.9+ is installed with necessary libraries:
```bash
pip install matplotlib numpy python-docx pypdf
```

### Running the Full Simulator
To run all modules and display tabular performance statistics:
```bash
python full_simulator.py
```

### Generating System Architecture Diagrams
```bash
python draw_architecture.py
```

### Building Submission Documentation
```bash
python build_full_submission.py
```

---


