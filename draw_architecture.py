import matplotlib.pyplot as plt
import matplotlib.patches as patches

def draw_architecture_diagram():
    fig, ax = plt.subplots(figsize=(12, 6.5))
    fig.patch.set_facecolor('#ffffff')
    ax.set_facecolor('#ffffff')
    ax.axis('off')
    
    # Title
    ax.text(0.5, 0.96, "UniCore Smart University OS Resource Orchestration Architecture",
            fontsize=14, fontweight='bold', ha='center', color='#1a1a2e', fontfamily='sans-serif')
    ax.text(0.5, 0.92, "End-to-End Request Flow: Process Scheduling -> Virtual Memory -> VFS / Allocation -> Disk I/O",
            fontsize=10, fontstyle='italic', ha='center', color='#555566', fontfamily='sans-serif')

    boxes = [
        {"title": "1. Client & Ingestion Layer", "desc": "500-800 Users: Exams, LMS,\nLibrary, Research, IoT", "xy": (0.05, 0.55), "wh": (0.16, 0.28), "color": "#e8f0fe", "edge": "#4285f4"},
        {"title": "2. Process & Concurrency Tier", "desc": "Round Robin (Interactive q=4ms)\nBanker's Deadlock Avoidance\nFair Readers-Writers Semaphores", "xy": (0.28, 0.55), "wh": (0.20, 0.28), "color": "#e6f4ea", "edge": "#34a853"},
        {"title": "3. Memory Management Tier", "desc": "16 GB Physical RAM (4M Frames)\nDemand Paging + Working Set\nClock/LRU Replacement Policy", "xy": (0.54, 0.55), "wh": (0.20, 0.28), "color": "#fef7e0", "edge": "#fbbc05"},
        {"title": "4. Storage & I/O Subsystem", "desc": "Indexed Allocation (ext4 Extents)\nPage Cache / Unified Buffer\nC-SCAN Disk Arm Scheduler", "xy": (0.79, 0.55), "wh": (0.18, 0.28), "color": "#fce8e6", "edge": "#ea4335"},
    ]
    
    for b in boxes:
        rect = patches.FancyBboxPatch(b["xy"], b["wh"][0], b["wh"][1], boxstyle="round,pad=0.02,rounding_size=0.03",
                                      facecolor=b["color"], edgecolor=b["edge"], linewidth=2)
        ax.add_patch(rect)
        cx = b["xy"][0] + b["wh"][0] / 2
        cy = b["xy"][1] + b["wh"][1] - 0.05
        ax.text(cx, cy, b["title"], fontsize=10, fontweight='bold', ha='center', color='#1a1a2e')
        ax.text(cx, cy - 0.08, b["desc"], fontsize=8.5, ha='center', color='#333333', va='top')

    # Connecting Arrows between main tiers
    arrow_props = dict(facecolor='#5f6368', edgecolor='#5f6368', width=2, headwidth=8, shrink=0.05)
    ax.annotate('', xy=(0.28, 0.69), xytext=(0.21, 0.69), arrowprops=arrow_props)
    ax.annotate('', xy=(0.54, 0.69), xytext=(0.48, 0.69), arrowprops=arrow_props)
    ax.annotate('', xy=(0.79, 0.69), xytext=(0.74, 0.69), arrowprops=arrow_props)

    # Detailed Flow Steps below
    flow_steps = [
        "Step 1: Exam Submission / Batch Job enters Ready Queue",
        "Step 2: CPU Scheduler invokes RR; Banker's Algorithm evaluates Resource Claims",
        "Step 3: MMU translates 64MB Virtual Addresses; Working Set monitor detects frame pressure",
        "Step 4: Page Fault initiates VFS Inode lookup; Page Cache checks for dirty buffers",
        "Step 5: Indexed Allocation resolves block addresses; C-SCAN queues cylinder seeks (0-199)"
    ]
    
    flow_box = patches.FancyBboxPatch((0.05, 0.08), 0.90, 0.38, boxstyle="round,pad=0.02,rounding_size=0.02",
                                      facecolor='#f1f3f4', edgecolor='#9aa0a6', linewidth=1.5)
    ax.add_patch(flow_box)
    ax.text(0.5, 0.41, "UniCore Request Lifecycle & Feedback Flow", fontsize=11, fontweight='bold', ha='center', color='#202124')
    
    for idx, text in enumerate(flow_steps):
        ax.text(0.08, 0.34 - (idx * 0.055), f"▶  {text}", fontsize=9, color='#3c4043', va='center')

    os.makedirs("screenshots", exist_ok=True)
    plt.savefig("screenshots/chart_architecture_flow.png", dpi=200)
    plt.close()
    print("Rendered: screenshots/chart_architecture_flow.png")

if __name__ == '__main__':
    draw_architecture_diagram()
