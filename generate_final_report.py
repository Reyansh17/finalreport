#!/usr/bin/env python3
"""
Generate Final Internship Report as a .docx file using raw Open XML.
No external dependencies required - uses only Python standard library.
"""

import zipfile
import os

# ============================================================
# OOXML Templates
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
    <w:rPr><w:sz w:val="24"/><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/>
    <w:pPr><w:spacing w:before="240" w:after="120"/></w:pPr>
    <w:rPr><w:b/><w:sz w:val="32"/><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading2">
    <w:name w:val="heading 2"/>
    <w:pPr><w:spacing w:before="200" w:after="100"/></w:pPr>
    <w:rPr><w:b/><w:sz w:val="28"/><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading3">
    <w:name w:val="heading 3"/>
    <w:pPr><w:spacing w:before="160" w:after="80"/></w:pPr>
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
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="1">
    <w:abstractNumId w:val="0"/>
  </w:num>
</w:numbering>'''


def escape_xml(text):
    """Escape special XML characters."""
    return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;").replace("'", "&apos;")


def make_paragraph(text, style=None, bold=False, bullet=False):
    """Create a paragraph XML element."""
    xml = '<w:p>'
    
    # Paragraph properties
    ppr_parts = []
    if style:
        ppr_parts.append(f'<w:pStyle w:val="{style}"/>')
    if bullet:
        ppr_parts.append('<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>')
    if ppr_parts:
        xml += '<w:pPr>' + ''.join(ppr_parts) + '</w:pPr>'
    
    # Run properties and text
    if text:
        xml += '<w:r>'
        if bold:
            xml += '<w:rPr><w:b/></w:rPr>'
        xml += f'<w:t xml:space="preserve">{escape_xml(text)}</w:t>'
        xml += '</w:r>'
    
    xml += '</w:p>'
    return xml


def make_bold_then_normal(bold_text, normal_text):
    """Create a paragraph with bold prefix followed by normal text."""
    xml = '<w:p>'
    xml += '<w:r><w:rPr><w:b/></w:rPr>'
    xml += f'<w:t xml:space="preserve">{escape_xml(bold_text)}</w:t></w:r>'
    xml += '<w:r>'
    xml += f'<w:t xml:space="preserve">{escape_xml(normal_text)}</w:t></w:r>'
    xml += '</w:p>'
    return xml


def build_document():
    """Build the complete document.xml content."""
    
    paras = []
    
    # Helper functions
    def h1(text): paras.append(make_paragraph(text, style="Heading1", bold=True))
    def h2(text): paras.append(make_paragraph(text, style="Heading2", bold=True))
    def h3(text): paras.append(make_paragraph(text, style="Heading3", bold=True))
    def p(text=""): paras.append(make_paragraph(text))
    def bold_p(text): paras.append(make_paragraph(text, bold=True))
    def bullet(text): paras.append(make_paragraph(text, bullet=True))
    def label_val(label, val): paras.append(make_bold_then_normal(label, val))
    def blank(): paras.append(make_paragraph(""))
    
    # ===== TITLE PAGE =====
    bold_p("ANNEXURE - VI")
    blank()
    bold_p("THAPAR INSTITUTE OF ENGINEERING & TECHNOLOGY, PATIALA")
    p("Department of Electronics & Communication Engineering")
    blank()
    bold_p("PROJECT SEMESTER: 2025-26 Even Semester (from January 2026)")
    blank()
    bold_p("Internship Student FINAL REPORT (during Second Visit)")
    blank()
    label_val("Date of Visit: ", "June 2026")
    blank()
    blank()
    
    # ===== 1. INTRODUCTION =====
    h1("1. Introduction")
    p("This report documents the complete technical work accomplished during my internship at Arm Ltd., Bangalore. The internship was placed within the CPU Verification Group, Interconnect Team, which is responsible for verifying the functional correctness and protocol compliance of on-chip interconnect fabrics. The team works with industry-standard protocols including AXI4 (Advanced eXtensible Interface), APB (Advanced Peripheral Bus), and CHI (Coherent Hub Interface for cache-coherent multi-core systems).")
    blank()
    p("The internship commenced with an intensive learning phase covering ARM's verification ecosystem \u2014 ARMv8-A processor architecture, UVM 1.2 methodology, and the testbench infrastructure used in industrial processor verification. This was followed by three primary technical contributions: (1) a Status Register Checker for automated RTL-to-model comparison, (2) systematic compile warning resolution across the verification environment, and (3) an AI-powered agent for automated hardware programming code generation from new RTL designs.")
    blank()
    
    # ===== 2. INTERNSHIP DETAILS =====
    h1("2. Internship Details")
    label_val("Organisation: ", "Arm Ltd., CPU Verification Group \u2014 Interconnect Team, Bangalore")
    label_val("Duration: ", "January 2026 \u2013 June 2026 (Project Semester 2025-26)")
    label_val("Protocols Worked On: ", "AXI4, APB, CHI (Coherent Hub Interface)")
    label_val("Tools & Technologies: ", "SystemVerilog, UVM 1.2, Python 3, EDA simulation tools, Arm Fast Model, LLM-based AI agents")
    blank()
    
    # ===== 3. LEARNING PHASE =====
    h1("3. Learning Phase")
    p("The first month was dedicated to building a strong technical foundation before undertaking any project contribution.")
    blank()
    
    h2("3.1 ARM Architecture and Protocols")
    p("I studied the ARMv8-A architecture \u2014 exception levels (EL0\u2013EL3), general-purpose registers (X0\u2013X30), system registers, and the privilege model. Understanding how system registers reflect processor state was directly relevant to my project work on register verification.")
    blank()
    p("For interconnect protocols, I worked through:")
    bullet("AXI4: Five independent channels \u2014 AW (address write), W (write data), B (write response), AR (address read), R (read data). The valid-ready handshake decouples address from data, enabling outstanding transactions and high throughput.")
    bullet("APB: A simpler, low-bandwidth protocol used for configuration registers and slow peripherals. Two-phase transfer (Setup \u2192 Access) with no pipelining. Chosen when low latency and simplicity matter over throughput.")
    bullet("CHI: Arm\u2019s cache-coherency protocol for multi-core SoCs. Three node types \u2014 RN-F (Request Node with full cache), HN-F (Home Node handling coherency), SN-F (Subordinate Node/memory). Four channels \u2014 REQ, RSP, DAT, SNP \u2014 each serving a specific role in the coherency state machine.")
    blank()
    
    h2("3.2 UVM 1.2 Methodology")
    p("Arm\u2019s verification environments are built on UVM 1.2. I studied and implemented the following components:")
    bullet("uvm_agent: Encapsulates driver, monitor, and sequencer as a self-contained verification block.")
    bullet("uvm_driver: Converts abstract transaction objects to pin-level stimulus on the DUT interface.")
    bullet("uvm_monitor: Observes DUT interface signals passively and publishes transaction objects via analysis ports.")
    bullet("uvm_scoreboard: Receives transactions from monitors and compares expected versus actual behaviour.")
    bullet("TLM analysis ports (uvm_analysis_port / uvm_analysis_imp): One-to-many broadcast mechanism used to connect monitors to scoreboards without tight coupling.")
    bullet("UVM factory and config_db: Enable component substitution and hierarchical configuration passing without modifying testbench source code.")
    blank()
    
    # ===== 4. PROJECT 1 =====
    h1("4. Project 1: Status Register Checker")
    p("The first technical contribution is the development of an automated checker for a CPU debug and trace subsystem\u2019s status register. This register reflects the operational state of an on-chip trace buffer \u2014 including wrap-around status, trigger capture, and acquisition completion \u2014 and must match, cycle-accurately, between the RTL implementation and the architectural reference model.")
    p("The checker consists of two tightly integrated components: a Python-based log parser and a UVM scoreboard.")
    blank()
    
    h2("4.1 Python Log Parser")
    p("The architectural reference model generates log files that can reach several gigabytes in size during full regression runs. Naively loading such files into memory is not viable. I implemented a class-based Python parser with the following design:")
    bullet("Streams the log file line-by-line \u2014 memory footprint remains flat regardless of file size.")
    bullet("Uses compiled regular expressions for pattern matching \u2014 identifying lines corresponding to register write events.")
    bullet("Extracts two fields per match: the hexadecimal register value and the simulation timestamp.")
    bullet("Stores results in a Python dictionary keyed by timestamp, providing O(1) lookup during scoreboard comparison.")
    bullet("Handles malformed entries with an error-recovery path \u2014 logs the skipped line and continues parsing.")
    bullet("Outputs are passed to the SystemVerilog environment via sim_opts command-line flags \u2014 no recompilation required to change log file path or verbosity.")
    blank()
    
    h2("4.2 UVM Scoreboard")
    p("The scoreboard is the central comparison engine. It receives register values from two sources: expected values (from the Python parser via sim_opts) and actual values (from the RTL monitor via a TLM analysis port).")
    blank()
    p("Key implementation decisions:")
    bullet("Maintains two independent queues \u2014 one for expected transactions, one for actual. Matching is performed on simulation time, not arrival order, to handle the race condition between model log timestamps and RTL simulation time.")
    bullet("Comparison is bit-wise and cycle-accurate. On every register write event, the scoreboard pops the corresponding expected entry and performs a full bit-field equality check.")
    bullet("Mismatch reporting calls uvm_fatal with the expected hexadecimal value, the actual hexadecimal value, and the simulation timestamp \u2014 providing complete diagnostic information for debug.")
    bullet("Pass-path events are logged at UVM_HIGH verbosity \u2014 present in debug runs, silent in regression.")
    bullet("Integration uses the UVM factory override mechanism \u2014 the checker is inserted into the existing testbench hierarchy without modifying any existing component instantiation.")
    blank()
    
    h2("4.3 Timestamp Synchronisation Challenge")
    p("One non-trivial implementation challenge was synchronising model log timestamps with RTL simulation time. Model logs are written in wall-clock order of the reference model, which does not always match the order in which the RTL monitor produces transactions. A naive FIFO comparison produced false mismatches in certain multi-transaction scenarios.")
    blank()
    p("The solution was to implement a timestamp-keyed lookup rather than a sequential queue: the scoreboard holds expected values in a dictionary indexed by simulation time, and retrieves the correct entry when the actual transaction arrives, regardless of relative ordering. This eliminated all false positives in testing.")
    blank()
    bold_p("Goal: Automated, zero-false-positive register comparison between RTL and architectural reference model, integrated into the live verification regression suite.")
    blank()
    
    # ===== 5. PROJECT 2 =====
    h1("5. Project 2: Compile Warning Resolution")
    p("The second major contribution was the systematic identification and resolution of all simulation compile warnings across the UVM verification environment. Arm mandates zero-warning compilation \u2014 any warning in a compile log is treated as a potential bug source, and a clean baseline is required before formal regression can begin.")
    blank()
    
    h2("5.1 Warning Categories Identified")
    p("Analysis of the compile logs across multiple regression configurations identified six distinct warning categories:")
    bullet("Unused Variables and Signals: Signals declared but never driven or read. Resolved by removing the declaration where the signal served no purpose, or adding a targeted pragma-suppress comment where the signal is intentionally retained (e.g., for debug visibility or future use).")
    bullet("Implicit 4-state to 2-state Conversions: Assignments from logic (4-state) to bit (2-state) without explicit casting. These risk silently discarding X or Z values that carry meaningful information about uninitialized state. Resolved by adding explicit cast operators.")
    bullet("Deprecated Language Constructs: SystemVerilog syntax valid in older tool versions but flagged in the current toolchain. Resolved by updating to IEEE 1800-2017 compliant equivalents.")
    bullet("Uninitialised Signals: Signals with no explicit reset or default value, producing X at simulation start. Resolved by adding explicit initial assignments or reset conditions.")
    bullet("Port Width Mismatches: Signal width at a module instantiation differs from the port declaration. Resolved by correcting widths and adding explicit casts at domain boundaries.")
    bullet("Unreachable Code: Branches or statements that the simulator\u2019s static analysis identifies as never reachable. Resolved by removing the dead code after confirming it serves no purpose.")
    blank()
    
    h2("5.2 Methodology and Validation")
    p("Each fix followed a strict process: identify the warning source, understand the original author\u2019s intent, apply the minimum change needed, and run a full regression before and after the patch batch. This discipline was important because verification code is interconnected \u2014 a change that appears local can affect simulation behaviour in distant components.")
    blank()
    p("A key lesson from this work: one early fix incorrectly resolved a port width mismatch without accounting for the downstream component\u2019s expected width, introducing a subtle functional change that surfaced as a regression failure three days later. That experience reinforced the importance of understanding context before touching production code.")
    blank()
    bold_p("Goal: Zero-warning compilation across all test configurations, providing a clean baseline where any future warning is immediately visible.")
    blank()
    
    # ===== 6. PROJECT 3 - AI AGENT =====
    h1("6. Project 3: AI Agent for Automated Hardware Programming Code Generation")
    p("The third and final major contribution was the development of an AI-powered agent that automatically generates hardware programming code whenever new RTL designs are introduced into the verification environment. This project addresses a significant bottleneck in the hardware development workflow: the manual effort required to write low-level programming sequences every time a new or modified RTL block is delivered.")
    blank()
    
    h2("6.1 Problem Statement")
    p("In a typical SoC verification flow, whenever new RTL is delivered \u2014 whether it is a new IP block, a modified register map, or an updated interconnect configuration \u2014 engineers must manually write the corresponding hardware programming code. This includes register configuration sequences, initialization routines, and programming flows that correctly set up the hardware for functional verification. This manual process is time-consuming, error-prone, and creates a bottleneck between RTL delivery and verification readiness.")
    blank()
    
    h2("6.2 Solution Architecture")
    p("The AI agent automates this process by analysing new RTL designs and generating the corresponding hardware programming code. The system operates as follows:")
    bullet("RTL Analysis: The agent parses incoming RTL source files to extract structural information including module interfaces, register definitions, signal hierarchies, and configuration parameters.")
    bullet("Context Understanding: The agent builds an understanding of the hardware block\u2019s functionality, its register map, programming constraints, and interdependencies with other blocks in the system.")
    bullet("Code Generation: Based on the extracted RTL information and learned patterns from existing programming sequences in the codebase, the agent generates hardware programming code that correctly configures and initialises the new hardware.")
    bullet("Validation: The generated code is validated against the RTL specification to ensure correctness of register addresses, field widths, access permissions, and programming order dependencies.")
    blank()
    
    h2("6.3 Technical Implementation")
    p("The agent leverages large language model (LLM) capabilities combined with structured RTL parsing to produce accurate hardware programming sequences. Key implementation aspects include:")
    bullet("RTL Parsing Pipeline: A front-end parser extracts register maps, port definitions, and configuration spaces from SystemVerilog RTL files, converting them into a structured intermediate representation.")
    bullet("Prompt Engineering: Carefully designed prompts provide the LLM with RTL context, existing code patterns, and hardware programming conventions specific to Arm\u2019s codebase to generate contextually appropriate code.")
    bullet("Pattern Learning: The agent references existing, verified programming sequences in the repository to learn team-specific coding patterns, naming conventions, and programming flow structures.")
    bullet("Iterative Refinement: Generated code undergoes automated checks for syntactic correctness, register address validity, and adherence to coding standards before being presented for engineer review.")
    blank()
    
    h2("6.4 Integration and Workflow")
    p("The agent is designed to integrate into the existing development workflow:")
    bullet("Triggered automatically when new RTL is committed or delivered to the verification environment.")
    bullet("Generates a draft of the hardware programming code that engineers can review, modify, and approve.")
    bullet("Reduces the turnaround time between RTL delivery and verification readiness from days to hours.")
    bullet("Maintains consistency across programming sequences by applying uniform patterns learned from the established codebase.")
    blank()
    
    h2("6.5 Results and Impact")
    p("The AI agent demonstrates the following outcomes:")
    bullet("Significant reduction in manual effort required to produce hardware programming code for new RTL blocks.")
    bullet("Consistent code quality and adherence to team coding standards across all generated outputs.")
    bullet("Faster verification readiness \u2014 programming code is available shortly after RTL delivery rather than requiring days of manual development.")
    bullet("Engineers can focus on reviewing and refining generated code rather than writing it from scratch, improving overall team productivity.")
    blank()
    bold_p("Goal: Automated generation of hardware programming code from new RTL designs, reducing manual effort and accelerating the RTL-to-verification pipeline.")
    blank()
    
    # ===== 7. RESULTS AND OUTCOMES =====
    h1("7. Results and Outcomes")
    p("The following outcomes have been achieved over the full duration of the internship:")
    bullet("Status register checker is fully implemented, tested against known-good and intentionally-corrupted scenarios, and integrated into the live regression suite.")
    bullet("The checker successfully identified a pre-existing register mismatch that had gone undetected in manual review \u2014 a subtle bit-field error triggered only under a specific trace acquisition scenario.")
    bullet("Zero-warning compilation achieved across all test configurations and regression targets.")
    bullet("Compile log noise substantially reduced, making genuine errors immediately visible.")
    bullet("Dead code eliminated across multiple testbench components, improving maintainability.")
    bullet("AI agent for hardware programming code generation successfully developed and demonstrated, showing significant reduction in manual coding effort for new RTL integration.")
    bullet("Comprehensive understanding of AXI4, APB, and CHI interconnect protocols acquired through specification study and waveform analysis.")
    bullet("Functional coverage bins added to the status register checker, covering meaningful combinations of status bit states.")
    bullet("Stress-test sequences developed targeting corner cases: rapid successive trace acquisitions, buffer wrap-around under maximum load, and boundary conditions at trace buffer capacity limits.")
    blank()
    
    # ===== 8. MAJOR CHALLENGES AND INNOVATIONS =====
    h1("8. Major Challenges and Innovations")
    p("The following technical challenges were encountered and resolved:")
    bullet("Timestamp race condition: solved by replacing sequential queue comparison with timestamp-keyed dictionary lookup, eliminating all false positives.")
    bullet("Multi-gigabyte log file performance: solved by streaming line-by-line with compiled regex patterns \u2014 memory footprint is constant regardless of file size.")
    bullet("Backward-compatible warning resolution: solved by running full regressions before and after every batch, and grouping warnings by root cause to apply safe, targeted fixes.")
    bullet("UVM factory integration: the checker was inserted into an existing, stable testbench environment using UVM factory overrides \u2014 zero modifications to existing component hierarchy required.")
    bullet("RTL-to-code mapping for AI agent: solved by building a structured intermediate representation from parsed RTL that provides sufficient context for accurate code generation while abstracting away implementation-specific details.")
    bullet("Code generation accuracy: addressed through iterative prompt refinement and pattern learning from existing verified code in the repository, ensuring generated code adheres to team standards.")
    blank()
    
    # ===== 9. SKILLS ACQUIRED =====
    h1("9. Skills Acquired")
    bullet("Industrial-scale UVM verification methodology and testbench architecture.")
    bullet("ARM interconnect protocol expertise \u2014 AXI4, APB, and CHI specification-level understanding.")
    bullet("Python automation for verification workflows including log parsing, regression triage, and code generation.")
    bullet("AI/LLM-based tooling for hardware development automation.")
    bullet("Production-quality SystemVerilog coding to Arm\u2019s coding and quality standards.")
    bullet("Working within large, multi-team codebases with strict version control and review processes.")
    bullet("Debug methodology for industrial processor verification environments.")
    blank()
    
    # ===== 10. CONCLUSION =====
    h1("10. Conclusion")
    p("This internship at Arm\u2019s CPU Verification Group has produced three significant technical contributions: an automated status register checker integrated into the live regression suite, a zero-warning compilation baseline across the verification environment, and an AI-powered agent for automated hardware programming code generation from new RTL designs. All deliverables were produced to Arm\u2019s coding and quality standards.")
    blank()
    p("The status register checker has proven its value by identifying a pre-existing bug and continues to operate as part of the daily regression infrastructure. The zero-warning baseline ensures that any future issue is immediately visible rather than lost in noise. The AI agent represents a forward-looking contribution that addresses a genuine productivity bottleneck in the RTL-to-verification workflow.")
    blank()
    p("Working with production-grade UVM environments, industrial-scale interconnect protocols, and cutting-edge AI tooling has provided practical experience that extends well beyond academic coursework. The combination of traditional verification engineering and modern AI-assisted development reflects the evolving landscape of hardware verification and positions these skills for continued relevance in the industry.")
    blank()
    blank()
    
    # ===== SIGNATURES =====
    p("(Signature of Faculty Mentor)\t\t\t\t(Signature of Industry Mentor)")
    blank()
    p("Name\u2026\u2026\u2026\u2026\u2026\u2026\u2026\u2026\u2026\u2026\u2026..\t\t\tName\u2026\u2026\u2026\u2026\u2026\u2026\u2026\u2026\u2026\u2026\u2026\u2026\u2026")
    p("Designation\u2026\u2026\u2026\u2026\u2026\u2026\u2026\u2026\u2026\t\t\tDesignation\u2026\u2026\u2026\u2026\u2026\u2026\u2026\u2026\u2026\u2026\u2026\u2026")
    
    # Build document XML
    body_xml = ''.join(paras)
    
    document_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    {body_xml}
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
    </w:sectPr>
  </w:body>
</w:document>'''
    
    return document_xml


def create_docx(output_path):
    """Create the .docx file as a ZIP archive with Open XML content."""
    
    document_xml = build_document()
    
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        zf.writestr('[Content_Types].xml', CONTENT_TYPES)
        zf.writestr('_rels/.rels', RELS)
        zf.writestr('word/_rels/document.xml.rels', WORD_RELS)
        zf.writestr('word/document.xml', document_xml)
        zf.writestr('word/styles.xml', STYLES)
        zf.writestr('word/numbering.xml', NUMBERING)
    
    print(f"Final report created successfully: {output_path}")
    print(f"File size: {os.path.getsize(output_path)} bytes")


if __name__ == '__main__':
    output_file = '/projects/sandbox/Final-Report-102206005.docx'
    create_docx(output_file)
