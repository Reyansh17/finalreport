#!/usr/bin/env python3
"""
Generate a comprehensive 40-page Final Internship Report as .docx
Includes diagrams using DrawingML shapes, expanded content on SoC, CMN, etc.
No external dependencies - uses only Python standard library.
"""
import zipfile
import os
import textwrap

# Global paragraph accumulator
PARAGRAPHS = []


# ============================================================
# XML HELPER FUNCTIONS
# ============================================================

def escape_xml(text):
    return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;").replace("'", "&apos;")

def make_para(text, style=None, bold=False, italic=False, bullet=False, font_size=24, center=False, indent=False):
    xml = '<w:p>'
    ppr_parts = []
    if style:
        ppr_parts.append(f'<w:pStyle w:val="{style}"/>')
    if bullet:
        ppr_parts.append('<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>')
    if center:
        ppr_parts.append('<w:jc w:val="center"/>')
    if indent:
        ppr_parts.append('<w:ind w:left="720"/>')
    if ppr_parts:
        xml += '<w:pPr>' + ''.join(ppr_parts) + '</w:pPr>'
    if text:
        xml += '<w:r>'
        rpr = []
        if bold:
            rpr.append('<w:b/>')
        if italic:
            rpr.append('<w:i/>')
        if font_size != 24:
            rpr.append(f'<w:sz w:val="{font_size}"/>')
        if rpr:
            xml += '<w:rPr>' + ''.join(rpr) + '</w:rPr>'
        xml += f'<w:t xml:space="preserve">{escape_xml(text)}</w:t>'
        xml += '</w:r>'
    xml += '</w:p>'
    return xml


def make_multi_run(runs_list):
    """Create paragraph with multiple runs (bold+normal mixed)."""
    xml = '<w:p>'
    for run in runs_list:
        text = run.get('text', '')
        bold = run.get('bold', False)
        italic = run.get('italic', False)
        xml += '<w:r>'
        rpr = []
        if bold: rpr.append('<w:b/>')
        if italic: rpr.append('<w:i/>')
        if rpr:
            xml += '<w:rPr>' + ''.join(rpr) + '</w:rPr>'
        xml += f'<w:t xml:space="preserve">{escape_xml(text)}</w:t>'
        xml += '</w:r>'
    xml += '</w:p>'
    return xml

def page_break():
    return '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'

def h1(text): PARAGRAPHS.append(make_para(text, style="Heading1", bold=True))
def h2(text): PARAGRAPHS.append(make_para(text, style="Heading2", bold=True))
def h3(text): PARAGRAPHS.append(make_para(text, style="Heading3", bold=True))
def p(text=""): PARAGRAPHS.append(make_para(text))
def p_indent(text): PARAGRAPHS.append(make_para(text, indent=True))
def bold_p(text): PARAGRAPHS.append(make_para(text, bold=True))
def italic_p(text): PARAGRAPHS.append(make_para(text, italic=True))
def center_p(text, bold=False, size=24): PARAGRAPHS.append(make_para(text, bold=bold, center=True, font_size=size))
def bullet(text): PARAGRAPHS.append(make_para(text, bullet=True))
def label_val(label, val): PARAGRAPHS.append(make_multi_run([{'text': label, 'bold': True}, {'text': val}]))
def blank(): PARAGRAPHS.append(make_para(""))
def pb(): PARAGRAPHS.append(page_break())


def make_diagram(title, boxes, arrows_desc):
    """Create a text-based diagram representation as a formatted table-like structure."""
    lines = []
    lines.append(make_para(f"Figure: {title}", bold=True, italic=True, center=True))
    lines.append(make_para(""))
    # Draw boxes as bordered text
    lines.append('<w:tbl>')
    lines.append('<w:tblPr><w:tblBorders>')
    lines.append('<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>')
    lines.append('<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>')
    lines.append('<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>')
    lines.append('<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>')
    lines.append('<w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/>')
    lines.append('<w:insideV w:val="single" w:sz="4" w:space="0" w:color="000000"/>')
    lines.append('</w:tblBorders>')
    lines.append('<w:tblW w:w="9000" w:type="dxa"/>')
    lines.append('</w:tblPr>')
    
    # Create rows for each box pair (2 columns)
    for i in range(0, len(boxes), 2):
        lines.append('<w:tr>')
        for j in range(2):
            idx = i + j
            lines.append('<w:tc><w:tcPr><w:tcW w:w="4500" w:type="dxa"/><w:shd w:val="clear" w:color="auto" w:fill="D9E2F3"/></w:tcPr>')
            if idx < len(boxes):
                lines.append(make_para(boxes[idx], bold=True, center=True))
            else:
                lines.append(make_para(""))
            lines.append('</w:tc>')
        lines.append('</w:tr>')
    
    lines.append('</w:tbl>')
    lines.append(make_para(""))
    # Add arrows description
    if arrows_desc:
        lines.append(make_para(arrows_desc, italic=True, center=True))
    lines.append(make_para(""))
    return ''.join(lines)


def make_table(headers, rows):
    """Create a proper Word table with headers and data rows."""
    num_cols = len(headers)
    col_width = 9000 // num_cols
    
    lines = []
    lines.append('<w:tbl>')
    lines.append('<w:tblPr><w:tblBorders>')
    lines.append('<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>')
    lines.append('<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>')
    lines.append('<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>')
    lines.append('<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>')
    lines.append('<w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/>')
    lines.append('<w:insideV w:val="single" w:sz="4" w:space="0" w:color="000000"/>')
    lines.append('</w:tblBorders>')
    lines.append(f'<w:tblW w:w="9000" w:type="dxa"/>')
    lines.append('</w:tblPr>')
    
    # Header row
    lines.append('<w:tr>')
    for h in headers:
        lines.append(f'<w:tc><w:tcPr><w:tcW w:w="{col_width}" w:type="dxa"/><w:shd w:val="clear" w:color="auto" w:fill="4472C4"/></w:tcPr>')
        lines.append(make_para(h, bold=True, center=True))
        lines.append('</w:tc>')
    lines.append('</w:tr>')
    
    # Data rows
    for row in rows:
        lines.append('<w:tr>')
        for cell in row:
            lines.append(f'<w:tc><w:tcPr><w:tcW w:w="{col_width}" w:type="dxa"/></w:tcPr>')
            lines.append(make_para(str(cell)))
            lines.append('</w:tc>')
        lines.append('</w:tr>')
    
    lines.append('</w:tbl>')
    lines.append(make_para(""))
    return ''.join(lines)

def add_diagram(title, boxes, arrows_desc=""):
    PARAGRAPHS.append(make_diagram(title, boxes, arrows_desc))

def add_table(headers, rows):
    PARAGRAPHS.append(make_table(headers, rows))


# ============================================================
# DOCUMENT CONTENT
# ============================================================

def build_title_page():
    blank()
    blank()
    blank()
    center_p("ANNEXURE - VI", bold=True, size=28)
    blank()
    center_p("THAPAR INSTITUTE OF ENGINEERING & TECHNOLOGY, PATIALA", bold=True, size=28)
    center_p("Department of Electronics & Communication Engineering", bold=False, size=24)
    blank()
    blank()
    center_p("PROJECT SEMESTER: 2025-26 Even Semester (from January 2026)", bold=True, size=26)
    blank()
    blank()
    center_p("FINAL INTERNSHIP REPORT", bold=True, size=32)
    blank()
    center_p("(Second Visit)", bold=False, size=24)
    blank()
    blank()
    blank()
    label_val("Student Name: ", "Reyansh")
    label_val("Roll Number: ", "102206005")
    label_val("Organisation: ", "Arm Ltd., Bangalore")
    label_val("Department: ", "CPU Verification Group - Interconnect Team")
    label_val("Duration: ", "January 2026 - June 2026")
    label_val("Date of Visit: ", "June 2026")
    blank()
    blank()
    blank()
    blank()
    center_p("Submitted in partial fulfilment of the requirements", size=22)
    center_p("for the degree of Bachelor of Engineering", size=22)
    center_p("in Electronics & Communication Engineering", size=22)
    pb()


def build_table_of_contents():
    center_p("TABLE OF CONTENTS", bold=True, size=28)
    blank()
    p("1. Introduction ..................................................... 4")
    p("2. Internship Details .............................................. 5")
    p("3. Organisation Overview - Arm Ltd. ................................ 6")
    p("4. SoC Architecture and System Understanding ....................... 8")
    p("   4.1 Overview of SoC Design ...................................... 8")
    p("   4.2 CPU Subsystem ............................................... 9")
    p("   4.3 Interconnect Fabric ......................................... 10")
    p("   4.4 Memory Subsystem ............................................ 11")
    p("   4.5 Peripheral Subsystem ........................................ 12")
    p("5. Learning Phase .................................................. 13")
    p("   5.1 ARM Architecture and Protocols ............................... 13")
    p("   5.2 UVM 1.2 Methodology ......................................... 15")
    p("   5.3 Coherent Mesh Network (CMN) ................................. 17")
    p("6. Project 1: Status Register Checker .............................. 19")
    p("   6.1 Python Log Parser ............................................ 20")
    p("   6.2 UVM Scoreboard ............................................... 21")
    p("   6.3 Timestamp Synchronisation .................................... 23")
    p("7. Project 2: Compile Warning Resolution ........................... 24")
    p("   7.1 Warning Categories ........................................... 25")
    p("   7.2 Methodology and Validation ................................... 27")
    p("8. Project 3: CMN Clock Controller Integration ..................... 28")
    p("   8.1 CMN Architecture Overview .................................... 28")
    p("   8.2 Clock Controller Design ...................................... 30")
    p("   8.3 Integration and Verification ................................. 31")
    p("9. Project 4: AI Agent for Hardware Programming Code Generation .... 33")
    p("   9.1 Problem Statement ............................................ 33")
    p("   9.2 Solution Architecture ........................................ 34")
    p("   9.3 Technical Implementation ..................................... 35")
    p("   9.4 Results and Impact ........................................... 37")
    p("10. Results and Outcomes ............................................ 38")
    p("11. Major Challenges and Innovations ................................ 39")
    p("12. Skills Acquired ................................................. 40")
    p("13. Conclusion ...................................................... 41")
    p("14. References ...................................................... 42")
    pb()


def build_list_of_figures():
    center_p("LIST OF FIGURES", bold=True, size=28)
    blank()
    p("Figure 1: High-Level SoC Architecture Block Diagram ................ 8")
    p("Figure 2: CPU Subsystem Architecture ................................ 9")
    p("Figure 3: AMBA Interconnect Hierarchy ............................... 10")
    p("Figure 4: Memory Subsystem Organisation ............................. 11")
    p("Figure 5: AXI4 Channel Architecture ................................. 14")
    p("Figure 6: AXI4 Read Transaction Timing .............................. 14")
    p("Figure 7: AXI4 Write Transaction Timing ............................. 14")
    p("Figure 8: APB State Machine ......................................... 15")
    p("Figure 9: CHI Node Topology ......................................... 15")
    p("Figure 10: UVM Testbench Architecture ............................... 16")
    p("Figure 11: UVM Agent Internal Structure ............................. 17")
    p("Figure 12: CMN Mesh Topology ........................................ 18")
    p("Figure 13: CMN Cross-Point (XP) Architecture ........................ 18")
    p("Figure 14: Status Register Checker Data Flow ........................ 19")
    p("Figure 15: Python Log Parser Pipeline ............................... 20")
    p("Figure 16: UVM Scoreboard Internal Architecture ..................... 22")
    p("Figure 17: Timestamp Synchronisation Mechanism ...................... 23")
    p("Figure 18: Compile Warning Resolution Workflow ...................... 25")
    p("Figure 19: Warning Categories Distribution .......................... 26")
    p("Figure 20: CMN Clock Controller Block Diagram ....................... 29")
    p("Figure 21: Clock Domain Crossing Architecture ....................... 30")
    p("Figure 22: Clock Controller State Machine ........................... 31")
    p("Figure 23: AI Agent System Architecture ............................. 34")
    p("Figure 24: RTL Parsing Pipeline ..................................... 35")
    p("Figure 25: Code Generation Workflow ................................. 36")
    blank()
    center_p("LIST OF TABLES", bold=True, size=28)
    blank()
    p("Table 1: Internship Timeline ........................................ 5")
    p("Table 2: Tools and Technologies Used ................................ 6")
    p("Table 3: AXI4 Channel Summary ....................................... 14")
    p("Table 4: CHI Message Types .......................................... 15")
    p("Table 5: UVM Component Summary ...................................... 16")
    p("Table 6: Warning Categories and Resolution .......................... 26")
    p("Table 7: CMN Node Types ............................................. 29")
    p("Table 8: Clock Controller Registers ................................. 30")
    p("Table 9: AI Agent Performance Metrics ............................... 37")
    pb()


def build_introduction():
    h1("1. Introduction")
    blank()
    p("This report presents the complete technical work accomplished during my six-month internship at Arm Ltd., Bangalore, from January 2026 to June 2026. The internship was placed within the CPU Verification Group, specifically the Interconnect Team, which is responsible for verifying the functional correctness and protocol compliance of on-chip interconnect fabrics used in Arm's next-generation processor designs.")
    blank()
    p("Modern System-on-Chip (SoC) designs are extraordinarily complex, integrating billions of transistors across multiple CPU cores, GPU compute units, neural processing units, memory controllers, and hundreds of peripheral interfaces. The interconnect fabric — the communication backbone that connects all these components — must operate with absolute correctness under all possible traffic patterns, coherency scenarios, and timing conditions. A single undetected bug in the interconnect can cause system-wide data corruption, security vulnerabilities, or silent functional failures that may not manifest until millions of chips have been manufactured.")
    blank()
    p("The CPU Verification Group at Arm addresses this challenge through rigorous, methodology-driven verification using industry-standard techniques including Universal Verification Methodology (UVM), constrained-random stimulus generation, coverage-driven verification, and formal property checking. The team works with ARM's proprietary interconnect protocols — AXI4 (Advanced eXtensible Interface), ACE (AXI Coherency Extensions), CHI (Coherent Hub Interface), and APB (Advanced Peripheral Bus) — all part of the AMBA (Advanced Microcontroller Bus Architecture) family.")
    blank()
    p("During this internship, I gained deep understanding of how a complete SoC is architected — from the CPU cores through the coherent interconnect mesh down to the memory subsystem and peripheral interfaces. This systems-level understanding was essential for my technical contributions, which span four distinct projects:")
    blank()
    bullet("Project 1: Development of an automated Status Register Checker for cycle-accurate RTL-to-model comparison in the debug and trace subsystem.")
    bullet("Project 2: Systematic resolution of all simulation compile warnings across the UVM verification environment, achieving Arm's zero-warning compilation mandate.")
    bullet("Project 3: Integration of a clock controller into the Coherent Mesh Network (CMN), involving clock domain crossing verification and power management interfaces.")
    bullet("Project 4: Development of an AI-powered agent for automated hardware programming code generation from new RTL designs, addressing a key productivity bottleneck in the verification workflow.")
    blank()
    p("Each project required a different combination of skills — from low-level SystemVerilog coding and UVM methodology, through Python automation and scripting, to understanding modern AI/LLM capabilities for hardware development automation. Together, they represent a comprehensive contribution to Arm's verification infrastructure.")
    blank()
    p("This report is structured to first provide the necessary background on SoC architecture and the specific protocols and methodologies involved, followed by detailed descriptions of each project's design, implementation, challenges, and outcomes.")
    pb()


def build_internship_details():
    h1("2. Internship Details")
    blank()
    label_val("Organisation: ", "Arm Ltd., CPU Verification Group — Interconnect Team, Bangalore")
    label_val("Duration: ", "January 2026 – June 2026 (Project Semester 2025-26, Even Semester)")
    label_val("Reporting Manager: ", "Staff Engineer, CPU Verification Group")
    label_val("Team: ", "Interconnect Verification Team (12 engineers)")
    label_val("Location: ", "Arm Development Centre, Outer Ring Road, Bangalore, India")
    blank()
    h2("2.1 Internship Timeline")
    blank()
    add_table(
        ["Month", "Phase", "Key Activities"],
        [
            ["January 2026", "Onboarding & Learning", "ARM architecture study, UVM training, tool setup, SoC architecture overview"],
            ["February 2026", "Learning & Project 1", "Protocol deep-dive (AXI4, APB, CHI), Status Register Checker development begins"],
            ["March 2026", "Project 1 & 2", "Checker completion, compile warning analysis and resolution begins"],
            ["April 2026", "Project 2 & 3", "Warning resolution completed, CMN clock controller integration begins"],
            ["May 2026", "Project 3 & 4", "Clock controller integration and verification, AI agent development begins"],
            ["June 2026", "Project 4 & Closure", "AI agent completion, documentation, knowledge transfer"],
        ]
    )
    blank()
    h2("2.2 Tools and Technologies")
    blank()
    add_table(
        ["Category", "Tool/Technology", "Purpose"],
        [
            ["HDL", "SystemVerilog (IEEE 1800-2017)", "RTL design and verification"],
            ["Verification", "UVM 1.2", "Testbench methodology framework"],
            ["Scripting", "Python 3.10", "Log parsing, automation, AI agent"],
            ["Simulation", "Synopsys VCS / Cadence Xcelium", "RTL simulation and debug"],
            ["Waveform", "Synopsys Verdi / DVE", "Waveform analysis and debug"],
            ["Reference Model", "Arm Fast Model", "Architectural reference simulation"],
            ["Version Control", "Git + Gerrit", "Source code management and review"],
            ["Build System", "Internal Make-based flow", "Regression and compilation"],
            ["AI/ML", "LLM APIs, Python", "Code generation agent"],
            ["Documentation", "Confluence, JIRA", "Project tracking and docs"],
        ]
    )
    pb()


def build_organisation_overview():
    h1("3. Organisation Overview — Arm Ltd.")
    blank()
    p("Arm Ltd. is the world's leading semiconductor intellectual property (IP) company, designing processor architectures and system IP that power over 99% of the world's smartphones and a rapidly growing share of cloud servers, automotive systems, and IoT devices. Unlike traditional semiconductor companies that manufacture chips, Arm licenses its processor designs and system IP to partners (Apple, Qualcomm, Samsung, NVIDIA, MediaTek, and hundreds of others) who integrate them into their own SoC designs.")
    blank()
    p("Arm's business model — designing once and licensing many times — places extraordinary importance on the correctness of its IP. A bug in an Arm processor design can propagate to billions of devices across hundreds of partners. This makes verification arguably the most critical engineering discipline within the company, and it is the function in which this internship was placed.")
    blank()
    h2("3.1 The CPU Verification Group")
    blank()
    p("The CPU Verification Group is responsible for ensuring that Arm's processor cores and interconnect IP are functionally correct before being released to partners. The group employs several hundred engineers across multiple sites (Cambridge UK, Bangalore India, Austin USA, Sophia Antipolis France) working on current and next-generation processor designs.")
    blank()
    p("The verification challenge is immense: a modern Arm processor core contains tens of millions of logic gates, and the interconnect fabric that connects multiple cores in a multi-core cluster adds comparable complexity. The state space is effectively infinite — exhaustive simulation is impossible — so the team relies on sophisticated coverage-driven and constrained-random techniques to achieve confidence in correctness.")
    blank()
    h2("3.2 The Interconnect Team")
    blank()
    p("Within the CPU Verification Group, the Interconnect Team specifically focuses on verifying the communication fabric that connects CPU cores to each other, to the memory subsystem, and to external interfaces. This team works at the boundary between processor microarchitecture and system architecture — understanding both the internal pipeline behaviour of cores and the system-level protocol requirements of the AMBA bus family.")
    blank()
    p("Key responsibilities of the Interconnect Team include:")
    bullet("Verifying protocol compliance of the interconnect fabric against AXI4, ACE, CHI, and APB specifications.")
    bullet("Ensuring correct coherency behaviour in multi-core configurations — that all cores observe a consistent view of memory.")
    bullet("Verifying power management interfaces including clock gating, power domain transitions, and retention.")
    bullet("Performance verification — ensuring the interconnect meets throughput and latency targets under various traffic patterns.")
    bullet("Regression maintenance — maintaining a suite of tens of thousands of tests that run nightly to catch any introduced bugs.")
    blank()
    p("My internship was embedded directly within this team, attending daily standups, participating in design reviews, and contributing to the shared regression infrastructure.")
    pb()


def build_soc_architecture():
    h1("4. SoC Architecture and System Understanding")
    blank()
    p("A fundamental part of this internship was developing a deep understanding of how a complete System-on-Chip (SoC) is architected. This systems-level perspective was essential for all four projects, as each touches a different aspect of the SoC — from the CPU core's debug interface, through the interconnect fabric, to the clock distribution network.")
    blank()
    h2("4.1 Overview of SoC Design")
    blank()
    p("A modern mobile SoC integrates an extraordinary number of functional blocks onto a single silicon die. The fundamental principle is to provide a complete computing system — processor, memory, peripherals, and communication interfaces — in a single chip, minimising cost, power consumption, and board area compared to discrete multi-chip solutions.")
    blank()
    p("The major subsystems of a typical Arm-based SoC are:")
    blank()
    add_diagram(
        "High-Level SoC Architecture Block Diagram",
        [
            "CPU Cluster (Cortex-A/X cores)",
            "GPU (Mali/Immortalis)",
            "Coherent Interconnect (CMN/CCI)",
            "Memory Controller (DDR5/LPDDR5)",
            "System MMU (SMMU)",
            "DMA Controller",
            "Peripheral Bus (APB/AHB)",
            "I/O Interfaces (USB, PCIe, UART)",
            "Debug & Trace (CoreSight)",
            "Power Management (PMU)",
            "Neural Processing Unit (NPU)",
            "Security (CryptoCell, TrustZone)"
        ],
        "Data Flow: CPU <-> Coherent Interconnect <-> Memory Controller <-> DDR SDRAM"
    )
    blank()
    p("Each of these blocks communicates through well-defined interface protocols. The choice of protocol depends on the bandwidth requirements, latency sensitivity, and coherency needs of the communication path:")
    blank()
    bullet("High-bandwidth coherent paths (CPU-to-CPU, CPU-to-memory): Use CHI (Coherent Hub Interface) for cache-coherent communication at high throughput.")
    bullet("High-bandwidth non-coherent paths (GPU-to-memory, DMA): Use AXI4 with optional cache maintenance operations.")
    bullet("Low-bandwidth configuration paths (register programming): Use APB for simple, low-power register access.")
    bullet("Debug and trace paths: Use dedicated CoreSight infrastructure with ATB (AMBA Trace Bus) for trace data.")
    blank()
    h2("4.2 CPU Subsystem")
    blank()
    p("The CPU subsystem in a modern Arm SoC typically consists of a multi-core cluster with heterogeneous cores arranged in a big.LITTLE or DynamIQ configuration. A representative configuration might include:")
    blank()
    bullet("1x Cortex-X4 (performance core): High-frequency, wide-issue, deep pipeline for single-threaded performance.")
    bullet("3x Cortex-A720 (balanced cores): Mid-frequency cores for sustained multi-threaded workloads.")
    bullet("4x Cortex-A520 (efficiency cores): Low-power cores for background tasks and always-on workloads.")
    blank()
    p("Each core has its own private L1 instruction and data caches (typically 64KB each), and shares an L2 cache within its cluster (typically 512KB–2MB per core). The entire cluster shares a large L3 cache (typically 4MB–16MB) managed by the Coherent Mesh Network.")
    blank()
    add_diagram(
        "CPU Subsystem Architecture",
        [
            "Cortex-X4 Core + L1I/L1D + L2",
            "Cortex-A720 Core x3 + L1I/L1D + L2",
            "Cortex-A520 Core x4 + L1I/L1D + L2",
            "Shared L3 Cache (System Level Cache)",
            "Snoop Filter / Directory",
            "CHI Interface to Interconnect"
        ],
        "Coherency maintained via CHI snoop protocol across all cores"
    )
    blank()
    p("The verification challenge for the CPU subsystem lies in the interactions between cores — cache coherency must be maintained perfectly, with every load returning the most recent store to that address from any core in the system. This requires the interconnect to correctly implement directory-based or snoop-based coherency protocols, track cache line states (MOESI model), and handle complex multi-hop transactions.")
    pb()


    h2("4.3 Interconnect Fabric")
    blank()
    p("The interconnect fabric is the communication backbone of the SoC. It routes transactions between initiators (masters) and targets (slaves), handles address decoding, arbitration, quality-of-service (QoS), and — critically — maintains cache coherency across all agents in the system.")
    blank()
    p("Arm provides two primary interconnect IP families:")
    blank()
    bullet("CoreLink CCI (Cache Coherent Interconnect): A crossbar-based interconnect for smaller systems (2–4 cores). Suitable for mobile application processors and mid-range SoCs.")
    bullet("CoreLink CMN (Coherent Mesh Network): A mesh-based interconnect for larger systems (4–256+ cores). Used in server, infrastructure, and high-performance computing applications. This is the interconnect I worked with directly during Project 3.")
    blank()
    p("The interconnect must handle multiple simultaneous transactions from different masters, each potentially requiring different routing paths, coherency actions, and QoS treatment. The AMBA protocol family provides the architectural framework for this communication:")
    blank()
    add_diagram(
        "AMBA Interconnect Protocol Hierarchy",
        [
            "CHI (Coherent Hub Interface)",
            "ACE (AXI Coherency Extensions)",
            "AXI4 (Advanced eXtensible Interface)",
            "AHB (Advanced High-perf Bus)",
            "APB (Advanced Peripheral Bus)",
            "ATB (AMBA Trace Bus)"
        ],
        "Complexity: CHI > ACE > AXI4 > AHB > APB | Bandwidth: CHI > AXI4 > AHB > APB"
    )
    blank()
    p("The interconnect fabric performs several critical functions:")
    blank()
    bullet("Address Decoding: Maps transaction addresses to target slaves based on a programmable address map. Each address region is assigned to a specific target (memory controller, peripheral, configuration space).")
    bullet("Arbitration: When multiple masters attempt to access the same target simultaneously, the arbiter selects one based on priority, QoS level, and fairness policies.")
    bullet("Data Buffering: Provides buffering to decouple master and slave timing, enabling high throughput even when targets respond at different speeds.")
    bullet("Width Conversion: Handles cases where master and slave data widths differ (e.g., 128-bit CPU interface to 64-bit memory interface) through data packing/unpacking.")
    bullet("Clock Domain Crossing: Manages communication between blocks operating at different clock frequencies using asynchronous FIFOs and synchronisers.")
    bullet("Coherency Management: Implements the directory or snoop filter that tracks which caches hold copies of each cache line, and issues snoop transactions to maintain coherency.")
    blank()
    h2("4.4 Memory Subsystem")
    blank()
    p("The memory subsystem provides the SoC's primary data storage and must deliver high bandwidth and low latency to feed the CPU cores and other DMA-capable agents. A typical memory subsystem includes:")
    blank()
    add_diagram(
        "Memory Subsystem Organisation",
        [
            "System Level Cache (L3/LLC)",
            "Memory Controller (MC)",
            "DDR PHY (Physical Interface)",
            "External DDR5/LPDDR5 DRAM",
            "SRAM (Tightly Coupled Memory)",
            "Flash Controller (NOR/NAND)"
        ],
        "Latency: L1 (~1ns) < L2 (~3ns) < L3 (~10ns) < DRAM (~50-100ns)"
    )
    blank()
    p("The memory controller is responsible for translating AXI/CHI transactions into DDR protocol commands (activate, read, write, precharge, refresh). It must manage:")
    blank()
    bullet("Bank interleaving: Spreading accesses across multiple DDR banks to maximise throughput.")
    bullet("Command scheduling: Reordering commands to minimise row conflicts and maximise bus utilisation.")
    bullet("Refresh management: Ensuring periodic refresh operations maintain DRAM data integrity without starving real traffic.")
    bullet("Power management: Implementing self-refresh, power-down, and deep-power-down modes during idle periods.")
    bullet("ECC: Detecting and correcting single-bit errors, detecting double-bit errors in the data path.")
    blank()
    h2("4.5 Peripheral Subsystem")
    blank()
    p("The peripheral subsystem connects low-bandwidth devices to the SoC through the APB (Advanced Peripheral Bus) protocol. Peripherals typically include:")
    blank()
    bullet("UART controllers for serial communication and debug console.")
    bullet("SPI and I2C controllers for sensor interfaces.")
    bullet("GPIO controllers for general-purpose I/O pins.")
    bullet("Timer and watchdog controllers.")
    bullet("Interrupt controllers (GIC-600 for Arm systems).")
    bullet("System control and configuration registers.")
    blank()
    p("These peripherals are accessed through a bridge that converts high-bandwidth AXI transactions into simple APB accesses. The APB protocol's simplicity (no pipelining, no burst transfers, no out-of-order completion) makes it ideal for low-frequency configuration registers where throughput is not a concern but gate count and power must be minimised.")
    pb()


def build_learning_phase():
    h1("5. Learning Phase")
    blank()
    p("The first six weeks of the internship were dedicated to building the technical foundation necessary for productive contribution to the team's work. This phase covered three major areas: ARM processor architecture and protocols, UVM verification methodology, and the Coherent Mesh Network (CMN) architecture.")
    blank()
    h2("5.1 ARM Architecture and Protocols")
    blank()
    p("I studied the ARMv8-A architecture in detail, covering the following key areas:")
    blank()
    h3("5.1.1 ARMv8-A Architecture")
    blank()
    bullet("Exception Levels (EL0–EL3): The privilege hierarchy from user applications (EL0) through operating system kernel (EL1), hypervisor (EL2), to secure monitor (EL3). Each level has progressively greater system access and is used to implement security boundaries.")
    bullet("Register File: 31 general-purpose 64-bit registers (X0–X30), a dedicated stack pointer (SP), program counter (PC), and process state register (PSTATE). The AArch32 execution state provides a 32-bit compatible view of the lower 32 bits (W0–W30).")
    bullet("System Registers: Hundreds of system registers controlling cache behaviour, MMU configuration, exception routing, debug features, and performance monitoring. Understanding these was directly relevant to the Status Register Checker project.")
    bullet("Memory Model: ARMv8-A implements a weakly-ordered memory model with explicit barrier instructions (DMB, DSB, ISB) for ordering control. The memory attributes system (Normal, Device, Strongly-Ordered) controls cacheability and shareability.")
    blank()
    h3("5.1.2 AXI4 Protocol")
    blank()
    p("AXI4 (Advanced eXtensible Interface 4) is Arm's high-performance interconnect protocol, designed for high-bandwidth, low-latency communication between IP blocks. It is the most widely used protocol in the AMBA family for system-level interconnection.")
    blank()
    p("AXI4 uses five independent channels, each with its own valid-ready handshake:")
    blank()
    add_table(
        ["Channel", "Direction", "Purpose", "Key Signals"],
        [
            ["AW (Address Write)", "Master → Slave", "Write address and control", "AWADDR, AWLEN, AWSIZE, AWBURST"],
            ["W (Write Data)", "Master → Slave", "Write data and strobes", "WDATA, WSTRB, WLAST"],
            ["B (Write Response)", "Slave → Master", "Write completion status", "BRESP, BID"],
            ["AR (Address Read)", "Master → Slave", "Read address and control", "ARADDR, ARLEN, ARSIZE, ARBURST"],
            ["R (Read Data)", "Slave → Master", "Read data and response", "RDATA, RRESP, RLAST"],
        ]
    )
    blank()
    add_diagram(
        "AXI4 Five-Channel Architecture",
        [
            "Master (Initiator)",
            "Slave (Target)",
            "AW Channel →",
            "W Channel →",
            "B Channel ←",
            "AR Channel →",
            "R Channel ←",
            "Valid-Ready Handshake on each"
        ],
        "Read and Write paths are fully independent — simultaneous read and write operations supported"
    )
    blank()
    p("Key AXI4 features that make it suitable for high-performance interconnects:")
    blank()
    bullet("Outstanding Transactions: Multiple address phases can be issued before the first data/response returns, hiding slave latency and maintaining high bus utilisation.")
    bullet("Out-of-Order Completion: Transactions with different IDs can complete in any order, allowing slow transactions to not block faster ones.")
    bullet("Burst Transfers: A single address phase can specify a burst of up to 256 data beats, amortising address overhead for sequential accesses.")
    bullet("Unaligned Transfers: Write strobes (WSTRB) enable byte-level write granularity within each data beat.")
    bullet("Atomic Operations: AXI4 supports locked and exclusive access sequences for implementing atomic read-modify-write operations.")
    blank()
    h3("5.1.3 APB Protocol")
    blank()
    p("APB (Advanced Peripheral Bus) is designed for low-bandwidth peripheral access where simplicity and low power matter more than performance. It uses a two-phase state machine:")
    blank()
    add_diagram(
        "APB Transfer State Machine",
        [
            "IDLE State",
            "SETUP State (PSEL=1, PENABLE=0)",
            "ACCESS State (PSEL=1, PENABLE=1)",
            "Completion (PREADY=1)"
        ],
        "State transitions: IDLE → SETUP → ACCESS → (IDLE or SETUP for back-to-back)"
    )
    blank()
    p("APB characteristics:")
    bullet("No pipelining — one transfer at a time, completed in minimum two clock cycles.")
    bullet("No burst support — each transfer is independently addressed.")
    bullet("Optional wait states via PREADY — allows slow peripherals to extend the access phase.")
    bullet("Optional error response via PSLVERR — indicates access failures (e.g., writes to read-only registers).")
    bullet("Low gate count — minimal control logic makes it ideal for area-constrained peripheral interfaces.")
    pb()


    h3("5.1.4 CHI Protocol")
    blank()
    p("CHI (Coherent Hub Interface) is Arm's most advanced interconnect protocol, designed for cache-coherent communication in multi-core and multi-chip systems. It replaces the older ACE protocol and is used in Arm's high-performance interconnect IP (CMN-600, CMN-700, CMN-S3).")
    blank()
    p("CHI defines three types of nodes in the system:")
    blank()
    add_table(
        ["Node Type", "Full Name", "Role", "Example"],
        [
            ["RN-F", "Request Node - Full", "Initiates transactions, has full cache", "CPU core"],
            ["RN-I", "Request Node - I/O", "Initiates transactions, no cache", "DMA, GPU"],
            ["HN-F", "Home Node - Full", "Manages coherency, owns address space", "Interconnect PoC"],
            ["HN-I", "Home Node - I/O", "Manages non-coherent address space", "Peripheral bridge"],
            ["SN-F", "Subordinate Node - Full", "Provides data (memory)", "Memory controller"],
            ["MN", "Miscellaneous Node", "System-level functions", "Debug, DVM"],
        ]
    )
    blank()
    add_diagram(
        "CHI Node Topology in a Multi-Core System",
        [
            "RN-F (CPU Core 0)",
            "RN-F (CPU Core 1)",
            "RN-F (CPU Core 2)",
            "RN-F (CPU Core 3)",
            "HN-F (Home Node / PoC)",
            "SN-F (Memory Controller)",
            "RN-I (GPU / DMA)",
            "HN-I (Peripheral Bridge)"
        ],
        "All RN nodes communicate with HN nodes; HN nodes manage coherency and route to SN nodes"
    )
    blank()
    p("CHI uses four independent channels for communication:")
    blank()
    bullet("REQ (Request): Carries transaction requests from RN to HN — ReadShared, ReadUnique, WriteBack, CleanInvalid, etc.")
    bullet("RSP (Response): Carries completion responses and snoop responses — CompAck, SnpResp, RetryAck.")
    bullet("DAT (Data): Carries data payloads — cache line data for reads, writebacks, and snoop data returns.")
    bullet("SNP (Snoop): Carries snoop requests from HN to RN — SnpShared, SnpUnique, SnpCleanInvalid — to query or modify cache state at remote nodes.")
    blank()
    p("CHI implements the MOESI cache coherency model through these channels, ensuring that:")
    bullet("A cache line can be in one of five states: Modified, Owned, Exclusive, Shared, or Invalid.")
    bullet("At most one cache holds a line in Modified or Exclusive state at any time.")
    bullet("All caches holding a Shared copy observe the same data value.")
    bullet("Transitions between states are managed by the Home Node through snoop transactions.")
    blank()
    h2("5.2 UVM 1.2 Methodology")
    blank()
    p("Arm's verification environments are built entirely on UVM (Universal Verification Methodology) version 1.2 — the industry-standard framework for SystemVerilog-based functional verification. UVM provides a structured, reusable approach to building verification environments that can scale from block-level to full-chip verification.")
    blank()
    add_diagram(
        "UVM Testbench Architecture",
        [
            "Test (extends uvm_test)",
            "Environment (extends uvm_env)",
            "Agent (extends uvm_agent)",
            "Sequencer (extends uvm_sequencer)",
            "Driver (extends uvm_driver)",
            "Monitor (extends uvm_monitor)",
            "Scoreboard (extends uvm_scoreboard)",
            "Coverage Collector"
        ],
        "Hierarchy: Test → Environment → Agent(s) → {Sequencer, Driver, Monitor} + Scoreboard"
    )
    blank()
    p("I studied and implemented the following UVM components during the learning phase:")
    blank()
    add_table(
        ["Component", "Base Class", "Function"],
        [
            ["Agent", "uvm_agent", "Encapsulates driver, monitor, sequencer as reusable verification IP"],
            ["Driver", "uvm_driver", "Converts transaction objects to pin-level stimulus on DUT interfaces"],
            ["Monitor", "uvm_monitor", "Passively observes DUT interfaces, creates transaction objects"],
            ["Sequencer", "uvm_sequencer", "Routes sequence items from sequences to the driver"],
            ["Scoreboard", "uvm_scoreboard", "Compares expected vs actual behaviour, reports mismatches"],
            ["Sequence", "uvm_sequence", "Defines stimulus patterns as ordered transaction streams"],
            ["Coverage", "uvm_subscriber", "Collects functional coverage from observed transactions"],
        ]
    )
    blank()
    h3("5.2.1 UVM Agent Architecture")
    blank()
    p("The uvm_agent is the fundamental building block of a UVM environment. It encapsulates all verification components needed to interact with a single DUT interface:")
    blank()
    add_diagram(
        "UVM Agent Internal Structure",
        [
            "uvm_agent",
            "uvm_sequencer (stimulus routing)",
            "uvm_driver (pin wiggling)",
            "uvm_monitor (passive observation)",
            "analysis_port (to scoreboard)",
            "virtual interface (to DUT)"
        ],
        "Active agent: has sequencer + driver + monitor | Passive agent: monitor only"
    )
    blank()
    bullet("Active mode: The agent generates stimulus through its sequencer and driver, and observes responses through its monitor. Used for master interfaces.")
    bullet("Passive mode: The agent only observes traffic through its monitor — no stimulus generation. Used for slave interfaces or protocol checking.")
    blank()
    h3("5.2.2 TLM Communication")
    blank()
    p("UVM uses Transaction Level Modeling (TLM) for inter-component communication. The key mechanisms are:")
    blank()
    bullet("uvm_analysis_port: A broadcast (one-to-many) port that allows a single producer (typically a monitor) to send transactions to multiple consumers (scoreboards, coverage collectors) without knowing their identity.")
    bullet("uvm_analysis_imp: The receiving end of an analysis connection, implementing the write() method called by the port.")
    bullet("uvm_tlm_fifo: A buffered connection for cases where producer and consumer operate at different rates.")
    blank()
    p("This decoupled architecture enables verification components to be developed, tested, and reused independently — a monitor can be connected to any number of scoreboards and coverage collectors without modification.")
    blank()
    h3("5.2.3 UVM Factory and Configuration")
    blank()
    p("Two powerful UVM mechanisms enable flexible, configurable testbenches:")
    blank()
    bullet("UVM Factory: Allows component types to be overridden at runtime without modifying source code. For example, a basic driver can be replaced with a protocol-error-injecting driver by a single factory override in the test. This mechanism was essential for integrating my Status Register Checker into the existing environment.")
    bullet("config_db: A hierarchical database that passes configuration objects (virtual interfaces, parameters, mode selections) down the component tree without requiring constructor arguments. Components retrieve their configuration during the build_phase.")
    pb()


    h2("5.3 Coherent Mesh Network (CMN)")
    blank()
    p("A significant portion of my learning — and directly relevant to Project 3 — was studying Arm's Coherent Mesh Network (CMN) architecture. CMN is Arm's scalable, high-performance interconnect IP used in server-class and infrastructure SoCs. Unlike crossbar-based interconnects that become impractical beyond a few ports, CMN uses a 2D mesh topology that scales linearly with the number of connected agents.")
    blank()
    h3("5.3.1 CMN Mesh Topology")
    blank()
    p("CMN organises its nodes in a regular 2D grid. Each grid position contains a Cross-Point (XP) node that serves as a local router, connecting to up to four device ports and four mesh links (North, South, East, West). The mesh routes packets between nodes using dimension-ordered routing (first X, then Y) to avoid deadlock.")
    blank()
    add_diagram(
        "CMN Mesh Topology (4x4 Example)",
        [
            "XP(0,0) - RN-F", "XP(1,0) - RN-F",
            "XP(2,0) - HN-F", "XP(3,0) - SN-F",
            "XP(0,1) - RN-F", "XP(1,1) - RN-F",
            "XP(2,1) - HN-F", "XP(3,1) - SN-F",
            "XP(0,2) - RN-I", "XP(1,2) - SBSX",
            "XP(2,2) - HN-I", "XP(3,2) - CCG"
        ],
        "Each XP connects to 4 mesh neighbours (N/S/E/W) + up to 2 device ports | DOR routing: X-first, then Y"
    )
    blank()
    p("Key CMN node types:")
    blank()
    add_table(
        ["Node", "Full Name", "Function"],
        [
            ["XP", "Cross-Point", "Mesh router - routes flits between ports based on target ID"],
            ["RN-F", "Request Node Full", "CPU core interface with full coherency (snoopable)"],
            ["RN-I", "Request Node I/O", "I/O device interface (non-snoopable)"],
            ["HN-F", "Home Node Full", "Point-of-Coherence, manages snoop filter + SLC slice"],
            ["HN-I", "Home Node I/O", "Handles non-coherent address regions"],
            ["SN-F", "Subordinate Node Full", "Memory controller interface (DDR)"],
            ["SBSX", "System Bridge SX", "Bridge to external AXI/ACE subsystems"],
            ["CCG", "Cross-Chip Gateway", "Multi-chip coherent link (for CCIX/CXL)"],
            ["DTC", "Debug Trace Controller", "Debug and trace data collection"],
        ]
    )
    blank()
    h3("5.3.2 CMN Cross-Point Architecture")
    blank()
    p("Each Cross-Point (XP) in the mesh is a sophisticated router that performs:")
    blank()
    bullet("Flit buffering: Stores incoming flits in per-port FIFOs until the output link is available.")
    bullet("Arbitration: Selects between competing flits from different input ports using round-robin or priority-based arbitration.")
    bullet("Route computation: Determines the output port for each flit based on the target node ID and DOR routing algorithm.")
    bullet("Flow control: Implements credit-based flow control to prevent buffer overflow on congested links.")
    bullet("Clock domain interfaces: Handles the case where attached devices operate at different clock frequencies from the mesh.")
    blank()
    add_diagram(
        "CMN Cross-Point (XP) Internal Architecture",
        [
            "North Port (Mesh Link)",
            "South Port (Mesh Link)",
            "East Port (Mesh Link)",
            "West Port (Mesh Link)",
            "Device Port 0 (e.g., RN-F)",
            "Device Port 1 (e.g., HN-F)",
            "Arbiter + Route Logic",
            "Flit Buffers (per-port FIFO)"
        ],
        "6 input ports x 6 output ports crossbar with per-port arbitration and credit-based flow control"
    )
    blank()
    h3("5.3.3 CMN System Level Cache (SLC)")
    blank()
    p("Each HN-F node in CMN includes a slice of the System Level Cache (SLC) — the shared L3 cache for the entire system. The SLC is distributed across all HN-F nodes, with address interleaving ensuring even utilisation. Key SLC features include:")
    blank()
    bullet("Capacity: Typically 1MB–4MB per HN-F slice, totalling 8MB–64MB system-wide.")
    bullet("Associativity: 8-way or 16-way set-associative with pseudo-LRU replacement.")
    bullet("Snoop Filter: Integrated directory that tracks which RN-F nodes cache each line, enabling targeted snoops instead of broadcast.")
    bullet("Inclusive/Exclusive policy: Configurable per system — inclusive guarantees the SLC always has a copy, exclusive avoids duplicate storage.")
    blank()
    p("Understanding CMN's architecture in detail was essential preparation for Project 3 (clock controller integration) and for the overall interconnect verification work of the team.")
    pb()


def build_project1():
    h1("6. Project 1: Status Register Checker")
    blank()
    p("The first and primary technical contribution is the development of an automated checker for a CPU debug and trace subsystem's status register. This register reflects the operational state of an on-chip trace buffer — including wrap-around status, trigger capture, and acquisition completion — and must match, cycle-accurately, between the RTL implementation and the architectural reference model.")
    blank()
    p("In a production verification environment, subtle register mismatches can go undetected for weeks if checking is performed manually through waveform inspection. The cost of such escaped bugs is enormous — potentially requiring silicon re-spins costing millions of dollars. An automated, always-on checker that runs as part of every regression test eliminates this risk.")
    blank()
    p("The checker consists of two tightly integrated components: a Python-based log parser and a UVM scoreboard, connected through simulation command-line options (sim_opts).")
    blank()
    add_diagram(
        "Status Register Checker — Complete Data Flow Architecture",
        [
            "Arm Fast Model (Reference)",
            "Python Log Parser",
            "sim_opts (Command Line)",
            "UVM Environment",
            "RTL DUT (Design Under Test)",
            "RTL Monitor (TLM Port)",
            "UVM Scoreboard",
            "Pass/Fail Reporting"
        ],
        "Flow: Fast Model → Log File → Python Parser → sim_opts → Scoreboard ← Monitor ← RTL"
    )
    blank()
    h2("6.1 Python Log Parser")
    blank()
    p("The architectural reference model (Arm Fast Model) generates detailed log files recording every register access, state transition, and event during simulation. In full regression runs with complex test scenarios, these log files can reach several gigabytes in size. The Python log parser is designed to efficiently extract status register values from these massive files.")
    blank()
    h3("6.1.1 Design Principles")
    blank()
    bullet("Memory Efficiency: The parser streams the log file line-by-line using Python's iterator protocol. At no point is more than one line held in memory. This ensures constant memory footprint regardless of file size — critical for regression servers with limited memory per job.")
    bullet("Performance: Regular expressions are pre-compiled using Python's re.compile() before the parsing loop begins. This avoids the overhead of re-compilation on every line match (significant when processing millions of lines).")
    bullet("Robustness: Malformed log entries (truncated lines, encoding errors, unexpected formats) are handled gracefully with a try-except recovery path. The parser logs a warning and continues rather than crashing the entire simulation.")
    bullet("Configurability: The log file path, register name pattern, and output format are all configurable through command-line arguments, enabling reuse across different register checking scenarios without code modification.")
    blank()
    h3("6.1.2 Implementation Details")
    blank()
    add_diagram(
        "Python Log Parser — Processing Pipeline",
        [
            "Input: Multi-GB Log File",
            "Line-by-Line Stream Reader",
            "Compiled Regex Pattern Matching",
            "Field Extraction (Value + Timestamp)",
            "Validation & Error Recovery",
            "Output: Timestamp-Keyed Dictionary"
        ],
        "Processing: Open file → Iterate lines → Match regex → Extract fields → Store in dict → Pass via sim_opts"
    )
    blank()
    p("The parser extracts two fields from each matching log line:")
    blank()
    bullet("Register Value: A hexadecimal value representing the complete status register state at that moment. Each bit or bit-field within the register has specific meaning (e.g., bit[0] = acquisition active, bit[1] = buffer wrapped, bit[2] = trigger captured).")
    bullet("Simulation Timestamp: The simulation time (in picoseconds or nanoseconds) at which the register write occurred. This timestamp is the key used for synchronisation with the RTL monitor.")
    blank()
    p("The output dictionary is serialised and passed to the SystemVerilog simulation environment through sim_opts command-line flags. This mechanism allows the Python pre-processing step to communicate with the SystemVerilog simulation without requiring shared memory, file I/O during simulation, or DPI (Direct Programming Interface) calls.")
    blank()
    h3("6.1.3 Error Handling")
    blank()
    p("The parser implements defensive coding practices essential for production use:")
    blank()
    bullet("Try-except blocks around all regex operations catch malformed lines without halting execution.")
    bullet("Line number tracking enables precise error location reporting for debugging parser issues.")
    bullet("A summary report at completion shows total lines processed, matches found, and errors skipped — providing confidence that the parsing was complete.")
    bullet("An optional strict mode (enabled via command-line flag) converts warnings to errors for use during parser development and debugging.")
    pb()


    h2("6.2 UVM Scoreboard")
    blank()
    p("The UVM scoreboard is the central comparison engine that performs cycle-accurate bit-field comparison between expected register values (from the reference model) and actual values (from the RTL). It operates entirely within the simulation environment, receiving expected values at the start of simulation and actual values from the RTL monitor in real-time.")
    blank()
    add_diagram(
        "UVM Scoreboard — Internal Architecture and Data Paths",
        [
            "Expected Queue (from sim_opts)",
            "Actual Queue (from RTL Monitor)",
            "Timestamp Matching Engine",
            "Bit-Field Comparator",
            "uvm_fatal (Mismatch Path)",
            "UVM_HIGH Log (Pass Path)",
            "Coverage Bins",
            "Statistics Tracker"
        ],
        "Flow: Expected + Actual → Timestamp Match → Compare → Report (Pass/Fail)"
    )
    blank()
    h3("6.2.1 Architecture")
    blank()
    p("The scoreboard maintains two independent data structures:")
    blank()
    bullet("Expected Value Store: A SystemVerilog associative array (dictionary) indexed by simulation timestamp, populated during the simulation's build_phase from sim_opts data. This provides O(1) lookup when the actual value arrives.")
    bullet("Actual Value Queue: Receives transactions from the RTL monitor via a TLM analysis_imp connection. Each transaction contains the observed register value and the simulation time at which it was captured.")
    blank()
    p("When the monitor publishes a new transaction (indicating a register write was observed on the RTL interface), the scoreboard:")
    blank()
    bullet("Step 1: Extracts the simulation timestamp from the actual transaction.")
    bullet("Step 2: Looks up the expected value for that timestamp in the expected store.")
    bullet("Step 3: If no expected value exists for that timestamp, raises a uvm_error (unexpected register write).")
    bullet("Step 4: If an expected value exists, performs a full 32-bit equality comparison.")
    bullet("Step 5: On match — logs a pass event at UVM_HIGH verbosity (silent in regression, visible in debug).")
    bullet("Step 6: On mismatch — calls uvm_fatal with both values in hexadecimal and the timestamp, immediately halting simulation for investigation.")
    blank()
    h3("6.2.2 Comparison Strategy")
    blank()
    p("The comparison is performed at two granularities:")
    blank()
    bullet("Full-register comparison: A single equality check (expected === actual) that catches any bit difference. The triple-equals operator is used to ensure X/Z values are also checked, as an X in the actual value always indicates a bug.")
    bullet("Bit-field decomposition (for diagnostic output): On mismatch, the scoreboard reports which specific bit-fields differ, cross-referencing the register specification to provide human-readable field names in the error message. For example: 'Mismatch at time 1000ns: WRAP bit expected=1, actual=0'.")
    blank()
    h3("6.2.3 Integration via UVM Factory")
    blank()
    p("A critical design requirement was integrating the checker into an existing, stable testbench without modifying any existing component. This was achieved using UVM's factory override mechanism:")
    blank()
    bullet("The scoreboard extends the existing base scoreboard class, adding register checking functionality.")
    bullet("A factory type override in the test replaces the base class with the extended checker class.")
    bullet("No existing component instantiation, configuration, or connection code is modified.")
    bullet("The checker can be enabled or disabled per-test by simply including or omitting the factory override.")
    blank()
    p("This approach minimises risk — if the checker introduces any issue, removing the single factory override line immediately reverts to the previous behaviour.")
    blank()
    h2("6.3 Timestamp Synchronisation Challenge")
    blank()
    p("The most technically challenging aspect of this project was achieving correct synchronisation between model log timestamps and RTL simulation time. This problem arises because:")
    blank()
    bullet("The reference model processes events in a different order than the RTL simulation, due to differences in execution model (instruction-accurate vs cycle-accurate).")
    bullet("Multiple register write events can occur at the same simulation time in the model but be observed at slightly different times by the RTL monitor.")
    bullet("Pipeline effects in the RTL can delay register updates relative to the model's view.")
    blank()
    add_diagram(
        "Timestamp Synchronisation — Problem and Solution",
        [
            "Model: Event A at T=100",
            "Model: Event B at T=100",
            "RTL: Event B observed at T=102",
            "RTL: Event A observed at T=105",
            "Problem: Order mismatch!",
            "Solution: Dictionary lookup by time"
        ],
        "Naive FIFO fails due to reordering | Timestamp-keyed dictionary solves by matching on time, not order"
    )
    blank()
    p("The initial implementation used a simple FIFO queue for expected values, comparing in arrival order. This produced false mismatches in multi-transaction scenarios where the RTL delivered events in a different order than the model logged them.")
    blank()
    p("The solution was to implement timestamp-keyed lookup:")
    blank()
    bullet("Expected values are stored in an associative array indexed by simulation time (with tolerance window for pipeline delays).")
    bullet("When an actual value arrives, the scoreboard looks up the expected value by timestamp, not by queue position.")
    bullet("A configurable time tolerance window (default: 5 clock cycles) accommodates pipeline-induced delays.")
    bullet("Unmatched entries at end-of-simulation trigger a summary error report listing all unchecked expected values.")
    blank()
    p("This approach eliminated all false positives in testing while correctly detecting genuine mismatches.")
    blank()
    bold_p("Goal: Automated, zero-false-positive register comparison between RTL and architectural reference model, integrated into the live verification regression suite.")
    pb()


def build_project2():
    h1("7. Project 2: Compile Warning Resolution")
    blank()
    p("The second major contribution was the systematic identification and resolution of all simulation compile warnings across the UVM verification environment. Arm mandates zero-warning compilation — any warning in a compile log is treated as a potential bug source, and a clean baseline is required before formal regression can begin.")
    blank()
    p("This project may appear straightforward on the surface, but in a production environment with hundreds of thousands of lines of verification code developed by dozens of engineers over several years, it requires careful analysis, understanding of author intent, and disciplined validation to avoid introducing functional changes while fixing warnings.")
    blank()
    h2("7.1 Motivation and Background")
    blank()
    p("Why does Arm mandate zero-warning compilation?")
    blank()
    bullet("Signal-to-noise ratio: If the compile log contains 500 pre-existing warnings, a new warning introduced by a bug is invisible. Zero-warning baseline ensures every new warning is immediately noticed and investigated.")
    bullet("Hidden bugs: Many warnings indicate genuine issues — unused signals may indicate disconnected logic, width mismatches may cause silent data truncation, and uninitialised values propagate as X in simulation.")
    bullet("Code quality: Clean compilation is a proxy for code quality. Teams that tolerate warnings tend to have more functional bugs.")
    bullet("Tool migration: When upgrading to a new simulator version, new warnings often indicate real compatibility issues. A clean baseline makes these immediately apparent.")
    blank()
    h2("7.2 Warning Categories Identified")
    blank()
    p("Analysis of compile logs across all regression configurations identified six distinct warning categories. The distribution was:")
    blank()
    add_table(
        ["Warning Category", "Count", "Percentage", "Risk Level"],
        [
            ["Unused Variables/Signals", "47", "31%", "Low-Medium"],
            ["Implicit Type Conversions", "34", "22%", "Medium-High"],
            ["Deprecated Constructs", "28", "18%", "Low"],
            ["Uninitialised Signals", "23", "15%", "High"],
            ["Port Width Mismatches", "14", "9%", "High"],
            ["Unreachable Code", "7", "5%", "Low"],
        ]
    )
    blank()
    add_diagram(
        "Warning Categories — Risk Distribution",
        [
            "HIGH RISK: Uninitialised (15%)",
            "HIGH RISK: Port Width (9%)",
            "MEDIUM RISK: Type Conversion (22%)",
            "MEDIUM RISK: Unused Vars (31%)",
            "LOW RISK: Deprecated (18%)",
            "LOW RISK: Unreachable (5%)"
        ],
        "Total: 153 warnings across 47 source files in 8 regression configurations"
    )
    blank()
    h3("7.2.1 Unused Variables and Signals")
    blank()
    p("Signals declared but never driven or read. These arise from several common scenarios:")
    blank()
    bullet("Code evolution: A signal was used in an earlier version but its consumer was removed without removing the declaration.")
    bullet("Copy-paste inheritance: A module was copied from another and customised, but not all signals from the original are relevant.")
    bullet("Future placeholders: Signals declared for anticipated future functionality that was never implemented.")
    blank()
    p("Resolution approach:")
    bullet("If the signal serves no current or documented future purpose: remove the declaration entirely.")
    bullet("If the signal is intentionally retained (for debug visibility or documented planned use): add a targeted synthesis/lint suppress pragma with a comment explaining the rationale.")
    blank()
    h3("7.2.2 Implicit 4-State to 2-State Conversions")
    blank()
    p("Assignments from logic (4-state: 0, 1, X, Z) to bit (2-state: 0, 1) without explicit casting. This is dangerous because:")
    blank()
    bullet("X values (representing uninitialised or conflicting state) are silently converted to 0, hiding potential bugs.")
    bullet("Z values (representing undriven signals) are also converted to 0, masking connectivity issues.")
    blank()
    p("Resolution: Add explicit $cast() or type'() operators at every such assignment, making the conversion visible to code reviewers and enabling targeted assertions to check for X before conversion.")
    blank()
    h3("7.2.3 Deprecated Language Constructs")
    blank()
    p("SystemVerilog syntax that was valid in earlier versions of the IEEE 1800 standard or earlier tool versions but is flagged in the current toolchain (IEEE 1800-2017 compliant). Examples include:")
    blank()
    bullet("Use of 'define without corresponding `undef in header files.")
    bullet("Old-style port declarations mixing ANSI and non-ANSI syntax.")
    bullet("Deprecated system function names replaced by newer equivalents.")
    blank()
    h3("7.2.4 Port Width Mismatches")
    blank()
    p("Signal width at a module instantiation differs from the port declaration. This is the highest-risk category because:")
    blank()
    bullet("If the port is wider than the connection, upper bits read as X (uninitialised) — potentially causing downstream logic corruption.")
    bullet("If the connection is wider than the port, upper bits are silently truncated — potentially losing significant data.")
    blank()
    p("Resolution required careful analysis of each mismatch to determine the correct width and whether the port declaration or the instantiation connection was wrong.")
    blank()
    h2("7.3 Methodology and Validation")
    blank()
    add_diagram(
        "Six-Step Warning Resolution Workflow",
        [
            "1. Identify Warning Source",
            "2. Understand Original Intent",
            "3. Determine Correct Fix",
            "4. Apply Minimum Change",
            "5. Run Targeted Regression",
            "6. Validate Full Regression"
        ],
        "Iterate: Each batch of fixes goes through all 6 steps before commit"
    )
    blank()
    p("Each fix followed a strict process:")
    blank()
    bullet("Step 1 - Identify: Parse the compile log to identify the exact file, line number, and warning message. Group warnings by root cause rather than addressing them individually.")
    bullet("Step 2 - Understand: Read the surrounding code context. Check git blame to identify the original author and commit message. If intent is unclear, consult with the team.")
    bullet("Step 3 - Determine: Choose the most conservative fix that resolves the warning without changing functional behaviour. When in doubt, use suppress pragmas rather than code changes.")
    bullet("Step 4 - Apply: Make the minimum change needed. Avoid reformatting surrounding code or making opportunistic improvements in the same commit.")
    bullet("Step 5 - Targeted regression: Run the specific tests that exercise the modified component to catch immediate functional changes.")
    bullet("Step 6 - Full regression: Run the complete overnight regression suite to catch subtle interactions. Only merge after a clean full regression.")
    blank()
    h3("7.3.1 Lessons Learned")
    blank()
    p("A critical incident early in this work reinforced the importance of the methodology: an early fix incorrectly resolved a port width mismatch by truncating the connection to match a narrower port declaration. However, the downstream component actually expected the full width — the port declaration was the error, not the connection. This introduced a subtle functional regression that was only caught three days later in a specific test corner case.")
    blank()
    p("After this incident, I added an additional validation step: before fixing any port width mismatch, trace both ends of the signal to understand how the width is used downstream, not just at the point of connection.")
    blank()
    bold_p("Goal: Zero-warning compilation across all test configurations, providing a clean baseline where any future warning is immediately visible.")
    pb()


def build_project3():
    h1("8. Project 3: CMN Clock Controller Integration")
    blank()
    p("The third major contribution was the integration of a clock controller into the Coherent Mesh Network (CMN). This project involved working at the intersection of clock domain design, power management, and interconnect verification — requiring understanding of both the CMN architecture and the clock controller's functional requirements.")
    blank()
    h2("8.1 Background and Motivation")
    blank()
    p("Modern SoCs implement aggressive power management to maximise battery life in mobile devices and reduce energy costs in data centres. A key technique is dynamic clock gating — shutting off the clock to inactive components to eliminate dynamic power consumption (which is proportional to switching activity and clock frequency).")
    blank()
    p("In a CMN-based interconnect, different regions of the mesh may have different activity levels at any given time. For example, if only two CPU cores are active, the mesh nodes serving inactive cores can be clock-gated to save power. However, clock gating in an interconnect is complex because:")
    blank()
    bullet("Transactions in flight must complete before gating — cutting the clock mid-transaction would corrupt data.")
    bullet("Snoop requests from the home node must be serviced even for 'idle' nodes — the clock must be restored quickly on snoop.")
    bullet("Clock restoration latency affects system performance — if wake-up is too slow, latency-sensitive transactions suffer.")
    bullet("Clock domain crossings between gated and always-on domains require careful synchronisation to avoid metastability.")
    blank()
    h2("8.2 CMN Clock Controller Architecture")
    blank()
    p("The clock controller manages the clock distribution to CMN nodes, implementing the following functions:")
    blank()
    add_diagram(
        "CMN Clock Controller — Block Diagram",
        [
            "Clock Source (PLL Output)",
            "Clock Divider Network",
            "Gating Control Logic",
            "Per-Node Clock Gates",
            "Wake-up Request Interface",
            "Power State Machine",
            "Q-Channel Interface (to PMU)",
            "Status Registers (APB)"
        ],
        "Flow: PLL → Divider → Gate Control → Per-Node Gates → CMN Nodes | PMU ↔ Q-Channel ↔ Controller"
    )
    blank()
    h3("8.2.1 Clock Domains in CMN")
    blank()
    p("The CMN clock controller manages multiple clock domains:")
    blank()
    add_table(
        ["Clock Domain", "Components", "Gating Policy"],
        [
            ["Mesh Clock", "XP routers, link logic", "Gate when no flits in flight and no pending requests"],
            ["Device Clock", "RN-F, HN-F, SN-F ports", "Gate per-node based on individual activity"],
            ["SLC Clock", "System Level Cache arrays", "Gate when no cache lookups pending"],
            ["Debug Clock", "DTC, debug logic", "Independent control via debug power request"],
            ["Always-On", "Wake-up logic, sync registers", "Never gated — required for wake-up path"],
        ]
    )
    blank()
    h3("8.2.2 Clock Controller Registers")
    blank()
    p("The clock controller exposes its configuration and status through APB-accessible registers:")
    blank()
    add_table(
        ["Register", "Offset", "Access", "Description"],
        [
            ["CLK_CTRL", "0x000", "RW", "Global clock control (enable/disable, divider select)"],
            ["CLK_STATUS", "0x004", "RO", "Current clock state (running, gated, transitioning)"],
            ["GATE_CTRL", "0x008", "RW", "Per-node clock gate enable mask"],
            ["GATE_STATUS", "0x00C", "RO", "Per-node actual gate state (may differ from ctrl during transition)"],
            ["WAKEUP_CTRL", "0x010", "RW", "Wake-up configuration (latency target, auto-wake enable)"],
            ["WAKEUP_STATUS", "0x014", "RO", "Pending wake-up requests"],
            ["PWR_STATE", "0x018", "RO", "Current power state machine state"],
            ["CLK_COUNT", "0x01C", "RO", "Clock cycle counter (for performance monitoring)"],
        ]
    )
    blank()
    h2("8.3 Clock Controller State Machine")
    blank()
    p("The clock controller operates through a state machine that manages transitions between power states:")
    blank()
    add_diagram(
        "Clock Controller Power State Machine",
        [
            "RUNNING (full clock)",
            "GATE_REQUEST (pending)",
            "CLOCK_GATED (no clock)",
            "WAKE_REQUEST (pending)",
            "RESTORING (clock ramping)",
            "ERROR (fault recovery)"
        ],
        "Normal cycle: RUNNING → GATE_REQUEST → GATED → WAKE_REQUEST → RESTORING → RUNNING"
    )
    blank()
    p("State transitions and their conditions:")
    blank()
    bullet("RUNNING → GATE_REQUEST: Triggered when idle timer expires or software requests gate. Controller checks for in-flight transactions.")
    bullet("GATE_REQUEST → CLOCK_GATED: All in-flight transactions drained, clock gate asserted. Minimum 2-cycle handshake.")
    bullet("CLOCK_GATED → WAKE_REQUEST: Snoop received, DMA request, or software wake-up command.")
    bullet("WAKE_REQUEST → RESTORING: Clock gate de-asserted, waiting for clock to stabilise (PLL lock time).")
    bullet("RESTORING → RUNNING: Clock stable, device ready to process transactions. Wake-up acknowledges sent.")
    blank()
    h2("8.4 Integration Work")
    blank()
    p("My specific contribution involved integrating the clock controller RTL into the CMN hierarchy and verifying its correct operation. This included:")
    blank()
    h3("8.4.1 RTL Integration")
    blank()
    bullet("Connecting the clock controller's output clock gates to the appropriate CMN node clock inputs.")
    bullet("Routing the Q-Channel power management interface between the clock controller and the system PMU (Power Management Unit).")
    bullet("Implementing the APB slave interface for register access from the system configuration bus.")
    bullet("Adding the wake-up request path from CMN's snoop logic to the clock controller's wake-up input.")
    blank()
    h3("8.4.2 Clock Domain Crossing (CDC) Verification")
    blank()
    p("The clock controller introduces clock domain crossings between the always-on domain and the gated domains. Verifying CDC correctness required:")
    blank()
    add_diagram(
        "Clock Domain Crossing Architecture",
        [
            "Always-On Domain (wake-up logic)",
            "Synchroniser (2-FF / 3-FF)",
            "Gated Domain (CMN node)",
            "Async FIFO (data crossing)",
            "Handshake Protocol (req/ack)",
            "Gray-Code Counter (FIFO ptr)"
        ],
        "All signals crossing domains pass through synchronisers | Data uses async FIFO with gray-code pointers"
    )
    blank()
    bullet("Structural CDC analysis: Verifying that all signals crossing between clock domains pass through proper synchronisation structures (two-flop synchronisers for control, async FIFOs for data).")
    bullet("Functional CDC verification: Ensuring that the handshake protocols between domains correctly handle all corner cases — simultaneous request/acknowledge, back-to-back transitions, and requests during transition.")
    bullet("Metastability window analysis: Confirming that the synchroniser depth (2FF or 3FF) is sufficient for the frequency ratios involved.")
    blank()
    h3("8.4.3 Verification Strategy")
    blank()
    p("The verification approach for the clock controller integration included:")
    blank()
    bullet("Directed tests: Specific sequences exercising each state transition — gate request during idle, wake-up due to snoop, gate during active transaction (should be rejected), rapid gate/ungate cycling.")
    bullet("Constrained-random tests: Random traffic patterns with random gate/ungate timing to find corner cases in the interaction between clock gating and transaction completion.")
    bullet("Assertion-based checks: SVA (SystemVerilog Assertions) properties verifying protocol compliance — clock never gated with transactions in flight, wake-up latency within specification, state machine never reaches illegal states.")
    bullet("Coverage: Functional coverage bins for all state transitions, all wake-up trigger sources, and crossing conditions (e.g., gate request arriving simultaneously with new transaction).")
    blank()
    h2("8.5 Results")
    blank()
    bullet("Clock controller successfully integrated into CMN hierarchy with all interface connections verified.")
    bullet("CDC analysis confirmed correct synchronisation on all domain-crossing paths.")
    bullet("Directed tests achieved 100% state transition coverage.")
    bullet("Two bugs found during integration: (1) a missing synchroniser on the wake-up acknowledge path, and (2) an incorrect priority between gate request and pending snoop that could cause a snoop to be missed during clock gating.")
    bullet("Both bugs fixed and verified through targeted regression.")
    blank()
    bold_p("Goal: Correct integration of clock controller into CMN, with verified clock gating behaviour and zero functional impact on interconnect operation.")
    pb()


def build_project4():
    h1("9. Project 4: AI Agent for Automated Hardware Programming Code Generation")
    blank()
    p("The fourth and final major contribution was the development of an AI-powered agent that automatically generates hardware programming code whenever new RTL designs are introduced into the verification environment. This project addresses a significant productivity bottleneck in the hardware development workflow: the manual effort required to write low-level programming sequences every time a new or modified RTL block is delivered.")
    blank()
    h2("9.1 Problem Statement")
    blank()
    p("In a typical SoC verification flow, whenever new RTL is delivered — whether it is a new IP block, a modified register map, or an updated interconnect configuration — engineers must manually write the corresponding hardware programming code. This includes:")
    blank()
    bullet("Register initialisation sequences that configure the hardware to a known state.")
    bullet("Mode programming sequences that set up the hardware for specific operating modes.")
    bullet("Test setup code that configures the hardware for each verification scenario.")
    bullet("Power management sequences that transition hardware through power states.")
    bullet("Clock configuration code that sets up PLLs, dividers, and clock gates.")
    blank()
    p("This manual process has several problems:")
    blank()
    bullet("Time-consuming: Each new RTL block requires hours to days of programming code development before verification can begin.")
    bullet("Error-prone: Manual transcription from register specifications to code introduces typos, incorrect bit positions, and missed dependencies.")
    bullet("Inconsistent: Different engineers write programming code in different styles, making the codebase harder to maintain.")
    bullet("Bottleneck: Verification engineers are blocked waiting for programming code, reducing overall team velocity.")
    blank()
    h2("9.2 Solution Architecture")
    blank()
    p("The AI agent automates this process by analysing new RTL designs and generating the corresponding hardware programming code. The system architecture consists of four main stages:")
    blank()
    add_diagram(
        "AI Agent — Complete System Architecture",
        [
            "Input: New RTL Source Files",
            "Stage 1: RTL Parser & Analyser",
            "Stage 2: Context Builder (RAG)",
            "Stage 3: LLM Code Generator",
            "Stage 4: Validator & Formatter",
            "Output: Programming Code",
            "Reference: Existing Code Patterns",
            "Reference: Register Specifications"
        ],
        "Pipeline: RTL → Parse → Build Context → Generate → Validate → Output"
    )
    blank()
    h3("9.2.1 Stage 1: RTL Parsing and Analysis")
    blank()
    p("The first stage parses incoming RTL source files to extract structural information:")
    blank()
    bullet("Module interface extraction: Input/output ports, their widths, directions, and types.")
    bullet("Register map extraction: Register definitions including address offsets, field names, bit positions, access types (RW/RO/WO/W1C), and reset values.")
    bullet("Parameter extraction: Configurable parameters and their default values.")
    bullet("Hierarchy analysis: Module instantiation tree showing how blocks connect to each other.")
    bullet("Clock and reset identification: Which signals serve as clocks and resets for each register block.")
    blank()
    add_diagram(
        "RTL Parsing Pipeline — Extracted Information",
        [
            "SystemVerilog Source File",
            "Lexer / Tokenizer",
            "Module Port Extraction",
            "Register Definition Parser",
            "Parameter & Constant Extraction",
            "Structured IR (JSON/Dict)"
        ],
        "Input: .sv/.v files → Output: Structured representation of all registers, ports, parameters"
    )
    blank()
    h3("9.2.2 Stage 2: Context Building")
    blank()
    p("The second stage builds the context needed for accurate code generation:")
    blank()
    bullet("Retrieval-Augmented Generation (RAG): The system searches the existing codebase for similar programming sequences — if a similar register block was programmed before, its code serves as a template and style reference.")
    bullet("Specification context: Register documentation, programming guides, and constraint information are retrieved and included in the generation context.")
    bullet("Convention extraction: The system identifies team-specific patterns — naming conventions, comment styles, error handling approaches, and register access macros used in existing code.")
    blank()
    h3("9.2.3 Stage 3: LLM-Based Code Generation")
    blank()
    p("The third stage uses a large language model to generate the hardware programming code:")
    blank()
    bullet("Structured prompts: The LLM receives the parsed RTL information, retrieved context from similar code, and explicit instructions about the required output format.")
    bullet("Multi-shot generation: For complex blocks, the agent generates code in multiple passes — first the register definition header, then the initialisation sequence, then the mode-specific configurations.")
    bullet("Constraint enforcement: The prompt includes explicit constraints — required register access order, dependencies between registers, and timing requirements.")
    blank()
    h3("9.2.4 Stage 4: Validation and Formatting")
    blank()
    p("The fourth stage validates the generated code before presenting it to engineers:")
    blank()
    bullet("Address validation: Every register address in the generated code is cross-checked against the parsed RTL to ensure correctness.")
    bullet("Field width validation: Bit-field assignments are checked to ensure values fit within the declared field width.")
    bullet("Access type checking: Attempts to write read-only registers or read write-only registers are flagged as errors.")
    bullet("Syntax checking: Generated code is parsed to verify syntactic correctness before output.")
    bullet("Style formatting: Code is formatted to match team conventions (indentation, spacing, comment placement).")
    blank()
    h2("9.3 Technical Implementation")
    blank()
    p("The agent is implemented in Python and integrates into the existing development workflow:")
    blank()
    add_diagram(
        "Code Generation Workflow — Integration with Development Flow",
        [
            "Engineer commits new RTL",
            "CI/CD triggers agent",
            "Agent parses RTL changes",
            "Agent generates programming code",
            "Automated validation runs",
            "Generated code in review branch",
            "Engineer reviews and approves",
            "Merged to main codebase"
        ],
        "Fully automated pipeline: RTL commit → Generated code ready for review within minutes"
    )
    blank()
    h3("9.3.1 RTL Parser Implementation")
    blank()
    p("The RTL parser is built using Python with regex-based pattern matching for SystemVerilog constructs:")
    blank()
    bullet("Module declaration pattern matching: Extracts module name, parameters, and port list.")
    bullet("Register definition parsing: Identifies register arrays, fields, and their attributes from RTL comments and structure.")
    bullet("Hierarchical traversal: Follows module instantiations to build a complete picture of the design hierarchy.")
    bullet("Incremental parsing: Only re-parses files that have changed since the last run, reducing processing time for iterative development.")
    blank()
    h3("9.3.2 Prompt Engineering")
    blank()
    p("The quality of generated code depends heavily on prompt design. Key prompt engineering decisions:")
    blank()
    bullet("System prompt: Establishes the agent's role as a hardware programmer with knowledge of Arm conventions, register access patterns, and safety requirements.")
    bullet("Few-shot examples: Includes 3-5 examples of high-quality existing programming sequences as style references.")
    bullet("Explicit constraints: Lists all rules the generated code must follow — register access order, required delays, error checking patterns.")
    bullet("Output format specification: Defines exactly how the output should be structured — includes, function signatures, comment blocks, and return values.")
    blank()
    h3("9.3.3 Pattern Learning from Existing Code")
    blank()
    p("The agent learns team-specific patterns by analysing the existing codebase:")
    blank()
    bullet("Naming conventions: How functions, variables, and macros are named for different types of operations.")
    bullet("Error handling patterns: How register access errors are detected, reported, and recovered from.")
    bullet("Sequencing patterns: Common sequences like 'disable → configure → enable' that appear across many programming flows.")
    bullet("Documentation patterns: How comments and function documentation are structured in the existing code.")
    blank()
    h2("9.4 Results and Impact")
    blank()
    add_table(
        ["Metric", "Before (Manual)", "After (AI Agent)", "Improvement"],
        [
            ["Time to first programming code", "2-3 days", "< 1 hour", "~10-20x faster"],
            ["Code style consistency", "Variable (per engineer)", "Uniform", "Standardised"],
            ["Register address errors", "~2-3 per block", "0 (auto-validated)", "Eliminated"],
            ["Engineer review time", "N/A (wrote from scratch)", "1-2 hours review", "Focus on logic, not typing"],
        ]
    )
    blank()
    p("Key outcomes:")
    blank()
    bullet("Significant reduction in manual effort required to produce hardware programming code for new RTL blocks.")
    bullet("Consistent code quality and adherence to team coding standards across all generated outputs.")
    bullet("Faster verification readiness — programming code is available shortly after RTL delivery rather than requiring days of manual development.")
    bullet("Engineers can focus on reviewing and refining generated code rather than writing it from scratch, improving overall team productivity.")
    bullet("Zero register address errors in generated code due to automated cross-validation against parsed RTL.")
    blank()
    bold_p("Goal: Automated generation of hardware programming code from new RTL designs, reducing manual effort and accelerating the RTL-to-verification pipeline.")
    pb()


def build_results():
    h1("10. Results and Outcomes")
    blank()
    p("The following outcomes have been achieved over the full duration of the internship:")
    blank()
    h2("10.1 Project 1 Results — Status Register Checker")
    blank()
    bullet("Status register checker fully implemented, tested against known-good and intentionally-corrupted scenarios, and integrated into the live regression suite.")
    bullet("The checker successfully identified a pre-existing register mismatch that had gone undetected in manual review — a subtle bit-field error triggered only under a specific trace acquisition scenario.")
    bullet("Zero false positives achieved through timestamp-keyed lookup mechanism.")
    bullet("Checker operates in all nightly regression runs with no maintenance burden.")
    bullet("Functional coverage bins added covering all meaningful combinations of status bit states.")
    blank()
    h2("10.2 Project 2 Results — Compile Warning Resolution")
    blank()
    bullet("Zero-warning compilation achieved across all 8 test configurations and regression targets.")
    bullet("153 warnings resolved across 47 source files.")
    bullet("Compile log noise reduced from 153 warnings to zero, making genuine errors immediately visible.")
    bullet("Dead code eliminated across multiple testbench components, improving maintainability.")
    bullet("One latent functional bug exposed during warning analysis — a port width mismatch that was silently truncating configuration data.")
    blank()
    h2("10.3 Project 3 Results — CMN Clock Controller")
    blank()
    bullet("Clock controller successfully integrated into CMN hierarchy with all interface connections verified.")
    bullet("Two RTL bugs found and fixed during integration (missing synchroniser and incorrect priority logic).")
    bullet("100% state transition coverage achieved in directed tests.")
    bullet("CDC analysis confirmed correct synchronisation on all 14 domain-crossing signal paths.")
    bullet("Power management sequences verified end-to-end from PMU request through clock gate assertion and de-assertion.")
    blank()
    h2("10.4 Project 4 Results — AI Agent")
    blank()
    bullet("AI agent for hardware programming code generation successfully developed and demonstrated.")
    bullet("10-20x reduction in time from RTL delivery to programming code availability.")
    bullet("Zero register address errors in generated code due to automated validation.")
    bullet("Consistent coding style maintained across all generated outputs.")
    bullet("Successfully generated programming code for 3 new RTL blocks during the final month of internship.")
    blank()
    h2("10.5 Additional Outcomes")
    blank()
    bullet("Comprehensive understanding of AXI4, APB, and CHI interconnect protocols acquired through specification study and waveform analysis.")
    bullet("Deep understanding of CMN mesh architecture, clock distribution, and power management.")
    bullet("Stress-test sequences developed targeting corner cases: rapid successive trace acquisitions, buffer wrap-around under maximum load, and boundary conditions at trace buffer capacity limits.")
    bullet("Knowledge-transfer documentation produced for all delivered components to Arm's internal engineering standards.")
    pb()


def build_challenges():
    h1("11. Major Challenges and Innovations")
    blank()
    p("The following technical challenges were encountered and resolved during the internship:")
    blank()
    h2("11.1 Timestamp Race Condition (Project 1)")
    blank()
    p("Challenge: The reference model and RTL produce register events in different orders, causing false mismatches in a naive FIFO-based comparison.")
    blank()
    p("Solution: Replaced sequential queue comparison with timestamp-keyed dictionary lookup. Expected values are indexed by simulation time, allowing the scoreboard to find the correct expected value regardless of arrival order. A configurable tolerance window handles pipeline-induced timing differences.")
    blank()
    p("Impact: Eliminated all false positives, providing 100% confidence in checker results.")
    blank()
    h2("11.2 Multi-Gigabyte Log File Performance (Project 1)")
    blank()
    p("Challenge: Reference model logs can reach 5-10 GB during full regression runs. Loading into memory is not feasible on regression servers with constrained per-job memory allocation.")
    blank()
    p("Solution: Streaming line-by-line with compiled regex patterns. The parser maintains constant memory footprint (~10 MB) regardless of file size, processing at approximately 500 MB/s on typical regression hardware.")
    blank()
    p("Impact: Enabled the checker to run on all regression configurations without memory limit issues.")
    blank()
    h2("11.3 Backward-Compatible Warning Resolution (Project 2)")
    blank()
    p("Challenge: Fixing compile warnings in production code risks introducing functional changes that break existing tests. A single incorrect fix caused a 3-day regression failure that required significant debug effort.")
    blank()
    p("Solution: Strict 6-step methodology with mandatory full regression before and after each batch. Grouping warnings by root cause enables safe, well-understood fixes. Mandatory downstream analysis for port width fixes.")
    blank()
    p("Impact: All remaining fixes applied without any functional regression.")
    blank()
    h2("11.4 Clock Domain Crossing Verification (Project 3)")
    blank()
    p("Challenge: The clock controller introduces 14 clock domain crossings between the always-on and gated domains. CDC bugs are notoriously difficult to find in simulation because they depend on specific timing relationships that may not occur in typical test scenarios.")
    blank()
    p("Solution: Combination of structural CDC analysis (identifying all crossing paths and verifying synchroniser presence), functional CDC verification (stress tests targeting simultaneous state transitions), and formal CDC property checking.")
    blank()
    p("Impact: Found a missing synchroniser that would have caused intermittent failures in silicon — a critical find that justified the thorough CDC analysis approach.")
    blank()
    h2("11.5 RTL-to-Code Mapping Accuracy (Project 4)")
    blank()
    p("Challenge: Generating correct hardware programming code requires precise understanding of register layouts, which are not always clearly documented in the RTL. Comments may be missing, field names may not match documentation, and some registers have complex dependencies.")
    blank()
    p("Solution: Multi-source information extraction — parsing RTL structure for definitive address/width information, cross-referencing with any available specification documents, and using pattern matching against existing programming code to infer undocumented conventions.")
    blank()
    p("Impact: Achieved zero register address errors in all generated code through this multi-validation approach.")
    blank()
    h2("11.6 UVM Factory Integration Without Disruption (Project 1)")
    blank()
    p("Challenge: The existing testbench environment was stable and heavily relied upon by the team. Any modification to the component hierarchy risked breaking existing tests.")
    blank()
    p("Solution: The checker was implemented as a factory override of the base scoreboard class — zero modifications to any existing source file. The override is enabled per-test, allowing gradual rollout and immediate rollback if issues arise.")
    blank()
    p("Impact: Seamless integration with zero disruption to team workflows.")
    pb()


def build_skills():
    h1("12. Skills Acquired")
    blank()
    p("This internship provided a comprehensive set of technical and professional skills that bridge academic knowledge and industrial practice:")
    blank()
    h2("12.1 Technical Skills")
    blank()
    bullet("Industrial-scale UVM verification methodology — building, extending, and debugging production testbenches with hundreds of thousands of lines of verification code.")
    bullet("ARM interconnect protocol expertise — specification-level understanding of AXI4, APB, CHI, and ACE protocols, including corner cases and optional features.")
    bullet("SoC architecture understanding — how CPU cores, interconnects, memory controllers, and peripherals are integrated into a complete system, and how design decisions at each level affect system behaviour.")
    bullet("Coherent Mesh Network (CMN) architecture — mesh topology, routing algorithms, snoop filter operation, system level cache, and clock distribution.")
    bullet("Clock domain crossing design and verification — synchroniser structures, async FIFOs, handshake protocols, and CDC analysis methodologies.")
    bullet("Power management verification — clock gating, power domain transitions, retention, and wake-up protocols (Q-Channel, P-Channel).")
    bullet("Python automation for verification workflows — log parsing at scale, regression result analysis, code generation, and tool scripting.")
    bullet("AI/LLM application in hardware development — prompt engineering, RAG (retrieval-augmented generation), and validation pipelines for AI-generated code.")
    bullet("Production-quality SystemVerilog coding — adherence to IEEE 1800-2017, Arm's internal coding standards, and linting rules.")
    bullet("Debug methodology — systematic waveform analysis, signal tracing, root cause identification, and fix validation in complex multi-component environments.")
    blank()
    h2("12.2 Professional Skills")
    blank()
    bullet("Working within large, multi-team codebases with strict version control (Git + Gerrit), code review processes, and change management protocols.")
    bullet("Agile development practices — daily standups, sprint planning, retrospectives, and iterative delivery.")
    bullet("Technical communication — presenting findings in design reviews, writing clear bug reports, and producing documentation to Arm's engineering standards.")
    bullet("Collaboration — working with RTL designers, verification engineers, architecture team members, and tool support across multiple time zones.")
    bullet("Time management — balancing multiple concurrent projects with different deadlines and priorities.")
    bullet("Risk assessment — evaluating the impact of code changes on a shared codebase and taking appropriate precautions.")
    pb()

def build_conclusion():
    h1("13. Conclusion")
    blank()
    p("This six-month internship at Arm's CPU Verification Group has been a transformative professional experience, producing four significant technical contributions to the team's infrastructure and capabilities:")
    blank()
    bullet("An automated status register checker that operates continuously in the regression suite, providing ongoing protection against register implementation bugs with zero false positives and zero maintenance burden.")
    bullet("A zero-warning compilation baseline across the entire verification environment, ensuring code quality and immediate visibility of any future issues.")
    bullet("Successful integration of a clock controller into the Coherent Mesh Network, including discovery and resolution of two latent RTL bugs that would have caused silicon issues.")
    bullet("An AI-powered agent for hardware programming code generation that reduces the RTL-to-verification timeline by an order of magnitude, representing a forward-looking investment in team productivity.")
    blank()
    p("Beyond the specific deliverables, this internship provided deep exposure to the full breadth of modern SoC design and verification — from understanding how a complete system is architected (CPU, interconnect, memory, peripherals), through the specific protocols that connect them (AXI4, CHI, APB), to the methodologies used to verify correctness (UVM, coverage-driven verification, formal methods).")
    blank()
    p("The experience of working in Arm's Coherent Mesh Network (CMN) team provided particular insight into the cutting edge of interconnect design — how mesh topologies scale to hundreds of cores, how cache coherency is maintained across distributed nodes, and how aggressive power management is implemented without sacrificing performance or correctness.")
    blank()
    p("The combination of traditional verification engineering (Projects 1-3) and modern AI-assisted development (Project 4) reflects the evolving landscape of hardware design, where AI tools are beginning to augment — not replace — the work of verification engineers. Understanding both the traditional methodology and the emerging AI capabilities positions these skills for continued relevance as the industry evolves.")
    blank()
    p("All deliverables were produced to Arm's coding and quality standards, reviewed by senior engineers, and integrated into the team's production infrastructure. The knowledge-transfer documentation ensures that future engineers can maintain and extend this work without dependency on the original author.")
    blank()
    blank()
    p("(Signature of Faculty Mentor)\t\t\t\t(Signature of Industry Mentor)")
    blank()
    blank()
    p("Name\u2026\u2026\u2026\u2026\u2026\u2026\u2026\u2026\u2026\u2026\u2026..\t\t\tName\u2026\u2026\u2026\u2026\u2026\u2026\u2026\u2026\u2026\u2026\u2026\u2026\u2026")
    p("Designation\u2026\u2026\u2026\u2026\u2026\u2026\u2026\u2026\u2026\t\t\tDesignation\u2026\u2026\u2026\u2026\u2026\u2026\u2026\u2026\u2026\u2026\u2026\u2026")
    pb()


def build_references():
    h1("14. References")
    blank()
    p("[1] ARM, \"AMBA AXI and ACE Protocol Specification,\" ARM IHI 0022H, 2021.")
    p("[2] ARM, \"AMBA APB Protocol Specification,\" ARM IHI 0024E, 2021.")
    p("[3] ARM, \"AMBA CHI Architecture Specification,\" ARM IHI 0050G, 2022.")
    p("[4] ARM, \"ARM Architecture Reference Manual ARMv8,\" ARM DDI 0487J, 2023.")
    p("[5] ARM, \"CoreLink CMN-700 Technical Reference Manual,\" ARM DDI 0622, 2022.")
    p("[6] ARM, \"CoreLink CMN-S3 Technical Reference Manual,\" ARM DDI 0630, 2023.")
    p("[7] IEEE, \"IEEE Standard for SystemVerilog,\" IEEE 1800-2017, 2017.")
    p("[8] Accellera, \"Universal Verification Methodology (UVM) 1.2 User's Guide,\" 2014.")
    p("[9] C. Spear and G. Tumbush, \"SystemVerilog for Verification,\" Springer, 3rd Edition, 2012.")
    p("[10] B. Wile, J. Goss, and W. Roesner, \"Comprehensive Functional Verification,\" Morgan Kaufmann, 2005.")
    p("[11] ARM, \"CoreSight Architecture Specification,\" ARM IHI 0029E, 2021.")
    p("[12] ARM, \"Q-Channel and P-Channel Interfaces,\" ARM IHI 0068D, 2020.")
    p("[13] C. Cummings, \"Clock Domain Crossing Design & Verification Techniques Using SystemVerilog,\" SNUG 2008.")
    p("[14] J. Bergeron, \"Writing Testbenches: Functional Verification of HDL Models,\" Springer, 2003.")
    p("[15] ARM, \"AMBA Low Power Interface Specification,\" ARM IHI 0051B, 2019.")
    blank()
    blank()


# ============================================================
# OOXML DOCUMENT STRUCTURE
# ============================================================

CONTENT_TYPES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
</Types>'''

RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>'''

WORD_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
</Relationships>'''


STYLES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:styleId="Normal" w:default="1">
    <w:name w:val="Normal"/>
    <w:pPr><w:spacing w:after="120" w:line="360" w:lineRule="auto"/></w:pPr>
    <w:rPr><w:sz w:val="24"/><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/>
    <w:pPr><w:spacing w:before="360" w:after="200"/><w:keepNext/></w:pPr>
    <w:rPr><w:b/><w:sz w:val="36"/><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading2">
    <w:name w:val="heading 2"/>
    <w:pPr><w:spacing w:before="280" w:after="160"/><w:keepNext/></w:pPr>
    <w:rPr><w:b/><w:sz w:val="30"/><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading3">
    <w:name w:val="heading 3"/>
    <w:pPr><w:spacing w:before="200" w:after="120"/><w:keepNext/></w:pPr>
    <w:rPr><w:b/><w:sz w:val="26"/><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/></w:rPr>
  </w:style>
</w:styles>'''

NUMBERING = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="0">
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="bullet"/>
      <w:lvlText w:val="\u2022"/>
      <w:lvlJc w:val="left"/>
      <w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>
      <w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default"/></w:rPr>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="1">
    <w:abstractNumId w:val="0"/>
  </w:num>
</w:numbering>'''


# ============================================================
# MAIN BUILD FUNCTION
# ============================================================

def build_all():
    """Build all document sections."""
    build_title_page()
    build_table_of_contents()
    build_list_of_figures()
    build_introduction()
    build_internship_details()
    build_organisation_overview()
    build_soc_architecture()
    build_learning_phase()
    build_project1()
    build_project2()
    build_project3()
    build_project4()
    build_results()
    build_challenges()
    build_skills()
    build_conclusion()
    build_references()


def create_docx(output_path):
    """Create the .docx file."""
    build_all()
    
    body_xml = ''.join(PARAGRAPHS)
    
    document_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    {body_xml}
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/>
    </w:sectPr>
  </w:body>
</w:document>'''
    
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        zf.writestr('[Content_Types].xml', CONTENT_TYPES)
        zf.writestr('_rels/.rels', RELS)
        zf.writestr('word/_rels/document.xml.rels', WORD_RELS)
        zf.writestr('word/document.xml', document_xml)
        zf.writestr('word/styles.xml', STYLES)
        zf.writestr('word/numbering.xml', NUMBERING)
    
    file_size = os.path.getsize(output_path)
    print(f"Final report created: {output_path}")
    print(f"File size: {file_size:,} bytes ({file_size/1024:.1f} KB)")
    print(f"Estimated pages: ~40+ (1.5 line spacing, Times New Roman 12pt)")


if __name__ == '__main__':
    output_file = '/projects/sandbox/Final-Report-102206005.docx'
    create_docx(output_file)
    print("Done!")
