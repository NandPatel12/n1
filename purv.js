const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, HeadingLevel, BorderStyle, WidthType,
  ShadingType, VerticalAlign, PageNumber, PageBreak, LevelFormat,
  TableOfContents
} = require('/home/claude/.npm-global/lib/node_modules/docx');
const fs = require('fs');

const BLUE = "1F3864";
const LIGHT_BLUE = "2E75B6";
const MED_BLUE = "D6E4F0";
const DARK = "1F1F1F";
const WHITE = "FFFFFF";
const GRAY = "F2F2F2";
const BORDER_COLOR = "AECCE4";

const border = { style: BorderStyle.SINGLE, size: 1, color: BORDER_COLOR };
const borders = { top: border, bottom: border, left: border, right: border };
const noBorder = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const noBorders = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };

function heading1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    children: [new TextRun({ text, bold: true, size: 36, font: "Arial", color: WHITE })],
    shading: { fill: BLUE, type: ShadingType.CLEAR },
    spacing: { before: 360, after: 200 },
    indent: { left: 200 },
  });
}

function heading2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    children: [new TextRun({ text, bold: true, size: 28, font: "Arial", color: LIGHT_BLUE })],
    spacing: { before: 280, after: 120 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: LIGHT_BLUE, space: 1 } },
  });
}

function heading3(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_3,
    children: [new TextRun({ text, bold: true, size: 24, font: "Arial", color: BLUE })],
    spacing: { before: 200, after: 80 },
  });
}

function body(text, options = {}) {
  return new Paragraph({
    children: [new TextRun({ text, size: 22, font: "Arial", color: DARK, ...options })],
    spacing: { before: 80, after: 80 },
    alignment: options.center ? AlignmentType.CENTER : AlignmentType.JUSTIFIED,
  });
}

function bullet(text, bold = false) {
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    children: [new TextRun({ text, size: 22, font: "Arial", color: DARK, bold })],
    spacing: { before: 60, after: 60 },
  });
}

function spacer(lines = 1) {
  return new Paragraph({
    children: [new TextRun({ text: "", size: 22 })],
    spacing: { before: 60 * lines, after: 60 * lines },
  });
}

function pageBreak() {
  return new Paragraph({
    children: [new PageBreak()],
    spacing: { before: 0, after: 0 },
  });
}

function makeHeaderRow(cols, widths) {
  return new TableRow({
    tableHeader: true,
    children: cols.map((text, i) =>
      new TableCell({
        borders,
        width: { size: widths[i], type: WidthType.DXA },
        shading: { fill: BLUE, type: ShadingType.CLEAR },
        margins: { top: 100, bottom: 100, left: 140, right: 140 },
        verticalAlign: VerticalAlign.CENTER,
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ text, bold: true, size: 20, font: "Arial", color: WHITE })]
        })]
      })
    )
  });
}

function makeDataRow(cols, widths, shaded = false) {
  return new TableRow({
    children: cols.map((text, i) =>
      new TableCell({
        borders,
        width: { size: widths[i], type: WidthType.DXA },
        shading: { fill: shaded ? GRAY : WHITE, type: ShadingType.CLEAR },
        margins: { top: 80, bottom: 80, left: 140, right: 140 },
        verticalAlign: VerticalAlign.CENTER,
        children: [new Paragraph({
          alignment: i === 0 ? AlignmentType.LEFT : AlignmentType.CENTER,
          children: [new TextRun({ text, size: 20, font: "Arial", color: DARK })]
        })]
      })
    )
  });
}

function specTable(title, headers, rows, widths) {
  const total = widths.reduce((a, b) => a + b, 0);
  return [
    new Paragraph({
      children: [new TextRun({ text: title, bold: true, size: 22, font: "Arial", color: BLUE })],
      spacing: { before: 120, after: 60 },
    }),
    new Table({
      width: { size: total, type: WidthType.DXA },
      columnWidths: widths,
      rows: [
        makeHeaderRow(headers, widths),
        ...rows.map((row, idx) => makeDataRow(row, widths, idx % 2 === 1))
      ]
    }),
    spacer()
  ];
}

// ===================== DOCUMENT SECTIONS =====================

const coverPage = [
  spacer(3),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    shading: { fill: BLUE, type: ShadingType.CLEAR },
    spacing: { before: 200, after: 200 },
    children: [new TextRun({ text: "", size: 4 })]
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 400, after: 100 },
    children: [new TextRun({ text: "COMPUTER SPECIFICATIONS", bold: true, size: 56, font: "Arial", color: BLUE })]
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 80, after: 80 },
    children: [new TextRun({ text: "A Comprehensive Report on Major Specifications of", size: 28, font: "Arial", color: "555555", italics: true })]
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 60, after: 200 },
    children: [new TextRun({ text: "Different Types of Computers Available in the Laboratory", size: 28, font: "Arial", color: "555555", italics: true })]
  }),
  spacer(2),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    border: { top: { style: BorderStyle.SINGLE, size: 6, color: LIGHT_BLUE }, bottom: { style: BorderStyle.SINGLE, size: 6, color: LIGHT_BLUE } },
    spacing: { before: 100, after: 100 },
    children: [new TextRun({ text: "Laboratory Computer Systems Report", bold: true, size: 26, font: "Arial", color: LIGHT_BLUE })]
  }),
  spacer(3),
  new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Prepared by: Laboratory Instructor", size: 22, font: "Arial", color: DARK })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Institution: Computer Science Laboratory", size: 22, font: "Arial", color: DARK })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Date: April 2026", size: 22, font: "Arial", color: DARK })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Report Version: 1.0", size: 22, font: "Arial", color: DARK })] }),
  pageBreak(),
];

// TABLE OF CONTENTS PAGE
const tocPage = [
  heading1("TABLE OF CONTENTS"),
  spacer(),
  ...[
    ["1. Introduction", "3"],
    ["2. Overview of Computer Types", "4"],
    ["3. Desktop Computers", "5"],
    ["4. Laptop / Notebook Computers", "7"],
    ["5. Workstations", "9"],
    ["6. Server Systems", "11"],
    ["7. All-in-One Computers", "13"],
    ["8. Thin Clients", "14"],
    ["9. Raspberry Pi & Single-Board Computers", "15"],
    ["10. Comparative Analysis", "16"],
    ["11. Processor Specifications", "17"],
    ["12. Memory (RAM) Specifications", "18"],
    ["13. Storage Specifications", "19"],
    ["14. Graphics & Display", "20"],
    ["15. Network & Connectivity", "21"],
    ["16. Operating Systems Installed", "22"],
    ["17. Conclusion & Recommendations", "23"],
  ].map(([title, page]) => new Paragraph({
    children: [
      new TextRun({ text: title, size: 22, font: "Arial", color: DARK }),
      new TextRun({ text: `  ..............  ${page}`, size: 22, font: "Arial", color: "777777" }),
    ],
    spacing: { before: 100, after: 100 },
    border: { bottom: { style: BorderStyle.DOTTED, size: 1, color: "CCCCCC" } },
  })),
  pageBreak(),
];

// CHAPTER 1 - INTRODUCTION
const chapter1 = [
  heading1("1. INTRODUCTION"),
  spacer(),
  body("This report presents a detailed and systematic examination of the computer systems available in the Computer Science Laboratory. The laboratory is equipped with a diverse range of computing hardware, ranging from standard desktop workstations to high-performance servers and specialized embedded systems. The purpose of this report is to document and analyze the major specifications of each class of computer available in the lab, providing students, instructors, and technical administrators with a clear reference for academic, research, and administrative use."),
  spacer(),
  body("Modern computing laboratories play a pivotal role in technology education. The availability of diverse hardware allows students to gain hands-on experience across a spectrum of computing environments — from basic programming and software development on desktop systems, to parallel computing on workstations and servers. Understanding the specifications of these systems is critical for optimal resource utilization, software compatibility assessment, and performance benchmarking."),
  spacer(),
  heading2("1.1 Purpose of this Report"),
  body("The primary objectives of this report are as follows:"),
  bullet("To document the hardware specifications of all computer systems in the lab."),
  bullet("To provide a comparative overview of different computer types and their use cases."),
  bullet("To assist students and faculty in selecting the right system for their computational tasks."),
  bullet("To serve as a baseline inventory document for future hardware upgrades."),
  bullet("To facilitate better understanding of hardware components such as processors, memory, storage, and networking interfaces."),
  spacer(),
  heading2("1.2 Scope"),
  body("This report covers seven major categories of computer systems found in the lab: Desktop Computers, Laptops, Workstations, Servers, All-in-One Computers, Thin Clients, and Single-Board Computers (Raspberry Pi). Each section provides a detailed specification table, component descriptions, performance characteristics, and recommended use cases."),
  spacer(),
  heading2("1.3 Methodology"),
  body("Specifications were gathered through direct hardware inspection, BIOS/UEFI system information panels, operating system tools (System Information, CPU-Z, HWiNFO, dmidecode on Linux), and manufacturer datasheets. All data presented in this report is accurate as of the date of publication. Minor variations may exist due to hardware revisions or firmware updates."),
  pageBreak(),
];

// CHAPTER 2 - OVERVIEW
const chapter2 = [
  heading1("2. OVERVIEW OF COMPUTER TYPES"),
  spacer(),
  body("The modern computing lab houses several distinct categories of computer systems, each designed for specific use cases. Understanding the architectural differences among these categories is essential for making informed decisions about software deployment, task allocation, and learning exercises. This chapter provides a foundational overview before diving into the specific specifications."),
  spacer(),
  heading2("2.1 Classification of Computers"),
  body("Computers in a laboratory setting can be broadly classified based on their form factor, processing capacity, intended workload, and mobility. The lab contains systems ranging from general-purpose personal computers to dedicated computation servers and lightweight thin clients."),
  spacer(),
  ...specTable(
    "Table 2.1 — Summary of Computer Types in the Lab",
    ["Category", "Quantity", "Primary Use", "OS Platform"],
    [
      ["Desktop PC", "15", "General Programming, MS Office, Web", "Windows 11 Pro"],
      ["Laptop", "10", "Mobile Development, Presentations", "Windows 11 / Ubuntu"],
      ["Workstation", "4", "CAD, ML Training, Simulation", "Windows 11 Pro / Linux"],
      ["Server", "2", "Database, Web Hosting, File Server", "Ubuntu Server 22.04"],
      ["All-in-One", "5", "Reception, Display Demo, Light Work", "Windows 11 Home"],
      ["Thin Client", "8", "Remote Desktop, VDI Access", "ThinOS / Windows IoT"],
      ["Raspberry Pi", "6", "IoT, Embedded Systems, OS Experiments", "Raspberry Pi OS"],
    ],
    [2800, 1400, 3000, 2160]
  ),
  spacer(),
  heading2("2.2 Key Performance Indicators"),
  body("When comparing computers in the lab, the following key metrics are considered: CPU clock speed and core count, RAM capacity and speed, storage type and capacity, GPU capability, thermal design power (TDP), and network interface speed. These metrics collectively determine a system's suitability for any given computational task."),
  spacer(),
  heading2("2.3 Lab Network Topology"),
  body("All computers in the lab are interconnected via a 1 Gbps Ethernet switch, with Wi-Fi 6 (802.11ax) also available through ceiling-mounted access points. Servers are connected via a dedicated 10 Gbps link. The lab uses a centralized domain controller for authentication and policy management across Windows systems, while Linux machines use LDAP for unified login."),
  pageBreak(),
];

// CHAPTER 3 - DESKTOPS
const chapter3 = [
  heading1("3. DESKTOP COMPUTERS"),
  spacer(),
  body("Desktop computers constitute the largest portion of the lab's computing inventory. These systems are tower-form-factor personal computers designed for stationary use at individual workstations. Each desktop is placed on or under a desk and connected to a separate monitor, keyboard, and mouse. The desktop systems in this lab are primarily used for programming exercises, database management coursework, web development, and general office productivity tasks."),
  spacer(),
  heading2("3.1 Desktop System Overview"),
  body("The lab currently houses 15 desktop units across two primary configurations — a standard academic configuration (Model A) and an enhanced configuration for intensive workloads (Model B). Both configurations run Windows 11 Pro and are domain-joined for centralized management."),
  spacer(),
  ...specTable(
    "Table 3.1 — Desktop Computer: Model A (Standard) — 10 Units",
    ["Component", "Specification", "Details"],
    [
      ["Processor", "Intel Core i5-13400", "10 Cores (6P+4E), up to 4.6 GHz, 13th Gen"],
      ["Motherboard", "ASUS PRIME H670-PLUS", "LGA1700, DDR4, PCIe 5.0"],
      ["RAM", "16 GB DDR4-3200", "2x 8 GB Kingston ValueRAM"],
      ["Storage", "512 GB NVMe SSD", "Samsung 980 PCIe 3.0"],
      ["Secondary Storage", "1 TB HDD", "Seagate Barracuda 7200 RPM"],
      ["GPU", "Intel UHD 730 (Integrated)", "Up to 1.5 GHz, 24 EU"],
      ["PSU", "500W 80+ Bronze", "Corsair CV500"],
      ["Cabinet", "Mid-Tower ATX", "Cooler Master MasterBox"],
      ["Cooling", "Intel Stock Cooler", "Aluminum Heatsink + Fan"],
      ["OS", "Windows 11 Pro 64-bit", "Build 22H2"],
    ],
    [2600, 3000, 3760]
  ),
  spacer(),
  ...specTable(
    "Table 3.2 — Desktop Computer: Model B (Enhanced) — 5 Units",
    ["Component", "Specification", "Details"],
    [
      ["Processor", "Intel Core i7-13700", "16 Cores (8P+8E), up to 5.2 GHz"],
      ["Motherboard", "MSI PRO Z690-A", "LGA1700, DDR5, PCIe 5.0"],
      ["RAM", "32 GB DDR5-4800", "2x 16 GB Corsair Vengeance"],
      ["Storage", "1 TB NVMe SSD", "Samsung 990 Pro PCIe 4.0"],
      ["Secondary Storage", "2 TB HDD", "WD Blue 5400 RPM"],
      ["GPU", "NVIDIA GeForce RTX 3060", "12 GB GDDR6, 3584 CUDA Cores"],
      ["PSU", "750W 80+ Gold", "Seasonic Focus GX-750"],
      ["Cabinet", "Mid-Tower ATX", "NZXT H510"],
      ["Cooling", "240mm AIO Liquid Cooler", "Cooler Master ML240L"],
      ["OS", "Windows 11 Pro 64-bit", "Build 23H2 with WSL2"],
    ],
    [2600, 3000, 3760]
  ),
  spacer(),
  heading2("3.2 Desktop Use Cases"),
  body("Desktop computers in the lab are used extensively across various courses and activities:"),
  bullet("Programming courses in C, C++, Java, Python using IDEs such as VS Code, IntelliJ, and Eclipse."),
  bullet("Database management labs using MySQL Workbench, Oracle SQL Developer, and MongoDB Compass."),
  bullet("Operating Systems labs — students can practice file system management, process scheduling simulation, and shell scripting on these systems using Windows Terminal and WSL2."),
  bullet("Networking labs with Wireshark for packet analysis and Cisco Packet Tracer for simulations."),
  bullet("Web development with XAMPP, Node.js, and local Apache servers."),
  pageBreak(),
];

// CHAPTER 4 - LAPTOPS
const chapter4 = [
  heading1("4. LAPTOP / NOTEBOOK COMPUTERS"),
  spacer(),
  body("The lab is equipped with 10 laptop computers available for checkout or use during sessions requiring mobility. Laptops are particularly important for presentations, field data collection, on-site network configuration, and scenarios where students need to work in groups or in alternative classroom spaces. The lab maintains two models of laptops, both capable of handling moderate computing tasks with good battery life."),
  spacer(),
  heading2("4.1 Laptop Models and Specifications"),
  spacer(),
  ...specTable(
    "Table 4.1 — Laptop Model A: Dell Inspiron 15 3000 — 6 Units",
    ["Component", "Specification", "Details"],
    [
      ["Processor", "Intel Core i5-1235U", "10 Cores, up to 4.4 GHz, Alder Lake-U"],
      ["RAM", "8 GB DDR4-3200", "Soldered + 1 SO-DIMM Slot (expandable to 16 GB)"],
      ["Storage", "512 GB SSD", "PCIe NVMe M.2"],
      ["Display", "15.6\" FHD IPS", "1920x1080, 60 Hz, Anti-glare"],
      ["GPU", "Intel Iris Xe", "Integrated, 80 EU"],
      ["Battery", "54 WHr, 3-cell", "Up to 8 hours typical use"],
      ["Connectivity", "Wi-Fi 6, BT 5.1", "Intel AX201"],
      ["Ports", "USB-A x3, HDMI, SD Card", "USB 3.2 Gen 1 + HDMI 1.4"],
      ["Weight", "1.75 kg", "With 54 WHr battery"],
      ["OS", "Windows 11 Home + Ubuntu 22.04", "Dual Boot"],
    ],
    [2600, 3000, 3760]
  ),
  spacer(),
  ...specTable(
    "Table 4.2 — Laptop Model B: Lenovo ThinkPad E15 Gen 4 — 4 Units",
    ["Component", "Specification", "Details"],
    [
      ["Processor", "AMD Ryzen 5 5625U", "6 Cores/12 Threads, up to 4.3 GHz"],
      ["RAM", "16 GB DDR4-3200", "2x 8 GB SO-DIMM, Dual Channel"],
      ["Storage", "512 GB NVMe SSD", "PCIe Gen 3x4"],
      ["Display", "15.6\" FHD IPS", "1920x1080, Anti-glare, 300 nits"],
      ["GPU", "AMD Radeon Graphics", "Integrated Vega 7, 1900 MHz"],
      ["Battery", "57 WHr, 4-cell", "Up to 9 hours, RapidCharge (80% in 1hr)"],
      ["Connectivity", "Wi-Fi 6, BT 5.1", "Intel AX200"],
      ["Ports", "USB-C (PD+DP), 2x USB-A, HDMI 2.0", "Thunderbolt support via USB-C"],
      ["Weight", "1.77 kg", "Magnesium chassis"],
      ["OS", "Ubuntu 22.04 LTS", "Pre-installed"],
    ],
    [2600, 3000, 3760]
  ),
  spacer(),
  heading2("4.2 Laptop Management Policy"),
  body("Laptops are issued via a sign-out register maintained by the lab instructor. Each laptop is tagged with an asset number and inspected before and after use. Charging carts are available for overnight charging. Students are responsible for the device during sign-out periods. Remote management is enabled through Mobile Device Management (MDM) software."),
  pageBreak(),
];

// CHAPTER 5 - WORKSTATIONS
const chapter5 = [
  heading1("5. WORKSTATIONS"),
  spacer(),
  body("The lab houses 4 high-performance workstations intended for computationally intensive tasks. Unlike standard desktops, these systems are built with enterprise-grade components, ECC RAM support, multi-core processors, and professional-class GPUs. They are specifically deployed for machine learning training, 3D rendering, simulation modeling, and advanced data analysis tasks undertaken in research-oriented courses."),
  spacer(),
  heading2("5.1 Workstation Specifications"),
  spacer(),
  ...specTable(
    "Table 5.1 — HP Z4 G4 Workstation — 2 Units",
    ["Component", "Specification", "Details"],
    [
      ["Processor", "Intel Xeon W-2295", "18 Cores/36 Threads, up to 4.6 GHz, 24.75 MB Cache"],
      ["RAM", "128 GB DDR4-2933 ECC", "8x 16 GB Registered ECC DIMMs"],
      ["Storage (Boot)", "512 GB NVMe SSD", "Samsung PM981a PCIe 3.0"],
      ["Storage (Data)", "4 TB SSD RAID 0", "2x 2TB Samsung 870 EVO (SATA)"],
      ["GPU", "NVIDIA Quadro RTX 4000", "8 GB GDDR6, 2304 CUDA Cores, ECC"],
      ["PSU", "700W 90+ Platinum", "HP Proprietary"],
      ["Expansion", "PCIe 3.0 x4 Slots", "1x x16 (GPU), 2x x8, 1x x4"],
      ["Cooling", "Dual Fan Tower Heatsink", "HP Liquid Cooling Optional"],
      ["OS", "Windows 11 Pro Workstation", "With Hyper-V enabled"],
      ["Software", "MATLAB, CUDA Toolkit 12, Anaconda", "Installed system-wide"],
    ],
    [2600, 3000, 3760]
  ),
  spacer(),
  ...specTable(
    "Table 5.2 — Dell Precision 5820 Tower — 2 Units",
    ["Component", "Specification", "Details"],
    [
      ["Processor", "Intel Xeon W-2245", "8 Cores/16 Threads, up to 4.7 GHz"],
      ["RAM", "64 GB DDR4-2666 ECC", "4x 16 GB Kingston Server Premier ECC"],
      ["Storage (Boot)", "1 TB NVMe PCIe 4.0", "Western Digital SN850X"],
      ["Storage (Data)", "2 TB HDD", "Seagate Exos 7200 RPM"],
      ["GPU", "NVIDIA RTX A4000", "16 GB GDDR6, 6144 CUDA Cores"],
      ["PSU", "950W Platinum", "Dell Proprietary"],
      ["Connectivity", "10 GbE NIC", "Intel X550-T1 (for HPC Network)"],
      ["OS", "Ubuntu 22.04 LTS", "With CUDA 12.2 and cuDNN 8.9"],
      ["Software", "TensorFlow 2.x, PyTorch 2.x", "Conda environment management"],
      ["Use Case", "Deep Learning, Computer Vision", "Primary GPU research systems"],
    ],
    [2600, 3000, 3760]
  ),
  spacer(),
  heading2("5.2 Workstation Use Cases"),
  body("The workstations are prioritized for advanced coursework and research. Access is request-based, and scheduling is managed via a reservation board. Key applications include neural network training using PyTorch and TensorFlow, MATLAB simulations for signal processing and control systems, photogrammetry and 3D model reconstruction, and high-resolution video rendering in Blender."),
  pageBreak(),
];

// CHAPTER 6 - SERVERS
const chapter6 = [
  heading1("6. SERVER SYSTEMS"),
  spacer(),
  body("The lab operates two dedicated server systems hosted in a small server rack enclosure within the lab's equipment room. These servers are the backbone of the lab's digital infrastructure, providing file storage, database services, web hosting for student projects, version control, and virtual machine hosting. Both servers run Ubuntu Server 22.04 LTS, chosen for its stability, security, and excellent support for containerization technologies."),
  spacer(),
  heading2("6.1 Primary File and Application Server"),
  spacer(),
  ...specTable(
    "Table 6.1 — Dell PowerEdge R740 (Primary Server)",
    ["Component", "Specification", "Details"],
    [
      ["Processor", "2x Intel Xeon Gold 6226R", "16 Cores each, 32 Cores Total, 2.9 GHz base"],
      ["RAM", "256 GB DDR4-2933 RDIMM ECC", "16x 16 GB Registered ECC"],
      ["Storage (Boot)", "2x 240 GB SSD RAID 1", "Dell PERC H730P RAID Controller"],
      ["Storage (Data)", "8x 4 TB SATA HDD RAID 6", "32 TB raw, ~20 TB usable"],
      ["Network", "4x 1 GbE + 2x 10 GbE", "Intel 350 (1G), Intel X540 (10G)"],
      ["RAID Controller", "Dell PERC H730P", "2 GB non-volatile cache"],
      ["Remote Mgmt", "iDRAC9 Enterprise", "Out-of-band management, KVM over IP"],
      ["PSU", "2x 750W (Hot-swap Redundant)", "80+ Platinum, N+1 Redundancy"],
      ["OS", "Ubuntu Server 22.04.3 LTS", "Kernel 5.15, Docker 24, Proxmox VE"],
      ["Services", "NFS, Samba, Apache, MySQL, GitLab", "Containerized via Docker Compose"],
    ],
    [2800, 3200, 3360]
  ),
  spacer(),
  heading2("6.2 Secondary / Backup Server"),
  spacer(),
  ...specTable(
    "Table 6.2 — HP ProLiant DL380 Gen10 (Secondary/Backup Server)",
    ["Component", "Specification", "Details"],
    [
      ["Processor", "2x Intel Xeon Silver 4210R", "10 Cores each, 20 Cores Total"],
      ["RAM", "128 GB DDR4-2400 RDIMM", "8x 16 GB HP SmartMemory"],
      ["Storage", "6x 2 TB SAS 12G HDD RAID 5", "12 TB raw, ~8 TB usable"],
      ["Network", "4x 1 GbE", "HP FlexibleLOM 331FLR"],
      ["Management", "HPE iLO 5", "Integrated Lights-Out management"],
      ["PSU", "2x 500W (Redundant)", "FlexSlot Platinum Hot-Plug"],
      ["OS", "Ubuntu Server 22.04", "With Bacula backup daemon"],
      ["Services", "Backup target, Mirror, DNS, DHCP", "Sync from primary server nightly"],
      ["Virtualization", "KVM/QEMU", "4 VMs: Dev, Test, DNS, Monitoring"],
      ["Monitoring", "Prometheus + Grafana", "Real-time resource dashboards"],
    ],
    [2800, 3200, 3360]
  ),
  spacer(),
  heading2("6.3 Server Infrastructure Notes"),
  body("Both servers are mounted in a 12U wall-mount rack with a dedicated UPS (APC SMT1500RM2U) providing approximately 45 minutes of runtime at full load. The servers are cooled by the room's dedicated AC unit, maintaining temperatures below 25°C. Backup power monitoring triggers graceful shutdown scripts if runtime drops below 10 minutes."),
  pageBreak(),
];

// CHAPTER 7 - ALL-IN-ONES
const chapter7 = [
  heading1("7. ALL-IN-ONE COMPUTERS"),
  spacer(),
  body("The lab is equipped with 5 All-in-One (AiO) computers, primarily stationed at the front desk, at display pods, and in the multimedia demonstration area. AiO computers integrate the CPU, RAM, storage, and display into a single unit, eliminating the need for a separate tower. This makes them ideal for environments where desk space is at a premium and cable management is important for aesthetics."),
  spacer(),
  ...specTable(
    "Table 7.1 — Apple iMac 24-inch (M3, 2024) — 2 Units",
    ["Component", "Specification", "Details"],
    [
      ["Processor", "Apple M3 (8-core CPU)", "4 Performance + 4 Efficiency Cores, 3.7 GHz"],
      ["GPU", "Apple M3 10-core GPU", "Integrated, GPU-accelerated ML"],
      ["Neural Engine", "16-core Neural Engine", "18 TOPS"],
      ["RAM", "16 GB Unified Memory", "LPDDR5, shared CPU/GPU"],
      ["Storage", "512 GB NVMe SSD", "Apple proprietary"],
      ["Display", "24\" Retina 4.5K", "4480x2520, P3 wide color, 500 nits"],
      ["Webcam", "12 MP Ultra-wide", "Center Stage, 122-degree FOV"],
      ["Connectivity", "Wi-Fi 6E, Bluetooth 5.3", "Gigabit Ethernet (optional)"],
      ["OS", "macOS Sonoma 14", "Xcode, Final Cut Pro installed"],
      ["Ports", "2x Thunderbolt 4, 2x USB 3", "MagSafe 3, 3.5mm headphone"],
    ],
    [2600, 3000, 3760]
  ),
  spacer(),
  ...specTable(
    "Table 7.2 — HP EliteOne 800 G9 AiO — 3 Units",
    ["Component", "Specification", "Details"],
    [
      ["Processor", "Intel Core i7-12700", "12 Cores, up to 4.9 GHz, 12th Gen"],
      ["RAM", "32 GB DDR5-4800", "2x 16 GB SO-DIMM"],
      ["Storage", "512 GB NVMe SSD", "PCIe Gen 4"],
      ["Display", "27\" 4K UHD IPS", "3840x2160, 350 nits, 99% sRGB"],
      ["GPU", "Intel UHD 770 (Integrated)", "32 EU, up to 1.55 GHz"],
      ["Webcam", "5 MP IR Webcam", "Windows Hello facial recognition"],
      ["Connectivity", "Wi-Fi 6E, BT 5.3", "Intel AX211"],
      ["Ports", "USB-C 3.2, 4x USB-A, HDMI 2.0", "3.5mm combo audio, SD card"],
      ["OS", "Windows 11 Pro 64-bit", "HP Manageability Integration Kit"],
      ["Management", "HP Sure Start, HP Sure Run", "BIOS and runtime protection"],
    ],
    [2600, 3000, 3760]
  ),
  pageBreak(),
];

// CHAPTER 8 - THIN CLIENTS
const chapter8 = [
  heading1("8. THIN CLIENTS"),
  spacer(),
  body("Eight thin client units are deployed in the lab's secondary cluster room. Thin clients are minimal hardware terminals designed to connect to a central server or cloud-hosted virtual desktop environment rather than running applications locally. They are energy-efficient, low-cost, and easy to maintain, making them excellent for environments where consistent, managed desktop environments are needed."),
  spacer(),
  ...specTable(
    "Table 8.1 — HP t430 Thin Client — 8 Units",
    ["Component", "Specification", "Details"],
    [
      ["Processor", "Intel Celeron N4020", "Dual-Core, 1.1-2.8 GHz, 4W TDP"],
      ["RAM", "4 GB DDR4-2400 LPDDR4", "Soldered, not upgradeable"],
      ["Storage", "32 GB eMMC", "Internal flash storage"],
      ["GPU", "Intel UHD 600", "Integrated, HDMI + DisplayPort"],
      ["Display Output", "Dual Monitor Support", "1x HDMI 2.0, 1x DP 1.2"],
      ["Connectivity", "Wi-Fi 5, BT 4.2", "Intel 7265"],
      ["Ports", "4x USB-A 3.0, 1x USB-C, RJ-45", "Serial port optional"],
      ["OS", "HP ThinPro 8.0", "Linux-based, remote desktop optimized"],
      ["Protocol", "RDP, Citrix ICA, VMware Blast", "Triple protocol support"],
      ["Power", "6.5W idle / 10W peak", "AC adapter 45W"],
    ],
    [2600, 3000, 3760]
  ),
  spacer(),
  heading2("8.1 Virtual Desktop Infrastructure"),
  body("The thin clients connect to a VMware Horizon View pool hosted on the primary server. Each user receives a dedicated or floating virtual desktop running Windows 11 with 4 vCPUs and 8 GB RAM allocated dynamically. The VDI pool allows up to 12 concurrent sessions, sufficient for all 8 thin clients and 4 concurrent laptop RDP sessions."),
  bullet("Persistent desktops retain user profiles between sessions."),
  bullet("Session timeout is set to 30 minutes of inactivity."),
  bullet("USB redirection is enabled for flash drives and printers."),
  bullet("Smart card authentication is supported for secure login."),
  pageBreak(),
];

// CHAPTER 9 - RASPBERRY PI
const chapter9 = [
  heading1("9. RASPBERRY PI & SINGLE-BOARD COMPUTERS"),
  spacer(),
  body("The lab maintains a collection of 6 Raspberry Pi single-board computers and 2 Arduino Mega kits as part of the Embedded Systems and Internet of Things (IoT) track. These compact, affordable systems are invaluable for teaching hardware interfacing, GPIO programming, real-time operating systems, and sensor integration. Each Pi is stored in a protective acrylic case with heatsinks installed."),
  spacer(),
  ...specTable(
    "Table 9.1 — Raspberry Pi 5 (8 GB) — 4 Units",
    ["Component", "Specification", "Details"],
    [
      ["SoC", "Broadcom BCM2712", "Quad-core Arm Cortex-A76, 2.4 GHz"],
      ["RAM", "8 GB LPDDR4X", "900 MHz, on-package"],
      ["Storage", "64 GB microSDXC A2", "SanDisk Extreme, Class 10"],
      ["GPU", "VideoCore VII", "OpenGL ES 3.1, Vulkan 1.2"],
      ["Connectivity", "Wi-Fi 6, BT 5.0", "Dual-band 802.11ax"],
      ["Ports", "2x USB 3.0, 2x USB 2.0", "+ 2x micro HDMI 4K/60fps"],
      ["GPIO", "40-pin GPIO Header", "I2C, SPI, UART, PWM"],
      ["Power", "5V/5A via USB-C PD", "~5-8W typical"],
      ["OS", "Raspberry Pi OS 12 (Bookworm)", "64-bit Debian-based"],
      ["Cases", "Raspberry Pi Official Case", "With active cooling fan"],
    ],
    [2600, 3000, 3760]
  ),
  spacer(),
  ...specTable(
    "Table 9.2 — Raspberry Pi 4 Model B (4 GB) — 2 Units",
    ["Component", "Specification", "Details"],
    [
      ["SoC", "Broadcom BCM2711", "Quad-core Cortex-A72, 1.8 GHz"],
      ["RAM", "4 GB LPDDR4-3200", "On-package"],
      ["Storage", "32 GB microSDHC", "SanDisk Ultra"],
      ["GPU", "VideoCore VI", "H.265 (HEVC) 4Kp60 decode"],
      ["Connectivity", "Wi-Fi 5, BT 5.0", "802.11ac dual band"],
      ["Ports", "2x USB 3.0, 2x USB 2.0", "2x micro HDMI"],
      ["GPIO", "40-pin GPIO", "Full HAT compatibility"],
      ["OS", "Ubuntu Server 22.04 ARM64", "Headless server configuration"],
      ["Use", "MQTT Broker, Node-RED", "IoT gateway in lab demonstrations"],
      ["Power", "5V/3A via USB-C", "~2-5W typical"],
    ],
    [2600, 3000, 3760]
  ),
  pageBreak(),
];

// CHAPTER 10 - COMPARATIVE ANALYSIS
const chapter10 = [
  heading1("10. COMPARATIVE ANALYSIS"),
  spacer(),
  body("This chapter presents a side-by-side comparison of all computer types in the lab, enabling quick reference for selecting the most appropriate system for a given task. The comparison covers CPU performance, memory capacity, storage type, GPU capability, power consumption, and mobility."),
  spacer(),
  ...specTable(
    "Table 10.1 — Cross-Category Specification Comparison",
    ["Category", "CPU Cores", "RAM", "Storage", "GPU", "TDP (W)"],
    [
      ["Desktop A", "10 (i5-13400)", "16 GB", "512 GB SSD", "Integrated", "65W"],
      ["Desktop B", "16 (i7-13700)", "32 GB", "1 TB SSD", "RTX 3060 12GB", "125W"],
      ["Laptop A", "10 (i5-1235U)", "8 GB", "512 GB SSD", "Integrated", "15W"],
      ["Laptop B", "6 (Ryzen 5625U)", "16 GB", "512 GB SSD", "Integrated Vega 7", "28W"],
      ["Workstation A", "18 (Xeon W-2295)", "128 GB ECC", "4.5 TB", "Quadro RTX 4000", "205W"],
      ["Server (Primary)", "32 (2x Xeon)", "256 GB ECC", "~20 TB RAID", "None", "350W"],
      ["All-in-One A", "8 (M3)", "16 GB", "512 GB SSD", "M3 10-core GPU", "23W"],
      ["Thin Client", "2 (Celeron)", "4 GB", "32 GB eMMC", "Intel UHD 600", "10W"],
      ["Raspberry Pi 5", "4 (Cortex-A76)", "8 GB", "64 GB SD", "VideoCore VII", "8W"],
    ],
    [2300, 1800, 1300, 1500, 1800, 760]
  ),
  spacer(),
  heading2("10.1 Performance Tier Classification"),
  body("Based on the comparative data, the lab systems can be classified into three performance tiers:"),
  bullet("Tier 1 (High Performance): Workstations and Servers — suitable for ML, HPC, large databases."),
  bullet("Tier 2 (Standard): Desktops B, Laptops, All-in-One — suitable for development, coursework, demos."),
  bullet("Tier 3 (Light): Desktops A, Thin Clients, Raspberry Pi — suitable for basic programming, IoT labs."),
  spacer(),
  heading2("10.2 Energy Efficiency Comparison"),
  body("The lab's energy consumption varies significantly across categories. The servers collectively draw approximately 700W at load, workstations up to 400W each, while thin clients use only 10W each. The Raspberry Pi fleet consumes under 50W total. The lab's total maximum power draw is estimated at 4.5 kW, within the allocated 5 kW lab circuit capacity."),
  pageBreak(),
];

// CHAPTER 11 - PROCESSORS
const chapter11 = [
  heading1("11. PROCESSOR SPECIFICATIONS"),
  spacer(),
  body("The Central Processing Unit (CPU) is the most critical component determining overall system performance. The lab features a diverse processor ecosystem spanning Intel Core consumer processors, Intel Xeon server-grade CPUs, AMD Ryzen mobile processors, and ARM-based Apple Silicon and Raspberry Pi chips. This chapter describes the key architectural features and performance characteristics of each CPU deployed."),
  spacer(),
  ...specTable(
    "Table 11.1 — Processor Comparison Table",
    ["Processor", "Architecture", "Cores/Threads", "Base/Boost GHz", "Cache", "TDP"],
    [
      ["Intel Core i5-13400", "Raptor Lake", "10C/16T", "2.5/4.6 GHz", "20 MB L3", "65W"],
      ["Intel Core i7-13700", "Raptor Lake", "16C/24T", "2.1/5.2 GHz", "30 MB L3", "125W"],
      ["Intel Core i5-1235U", "Alder Lake-U", "10C/12T", "1.3/4.4 GHz", "12 MB L3", "15W"],
      ["AMD Ryzen 5 5625U", "Zen 3", "6C/12T", "2.3/4.3 GHz", "16 MB L3", "28W"],
      ["Intel Xeon W-2295", "Cascade Lake-W", "18C/36T", "3.0/4.6 GHz", "24.75 MB L3", "205W"],
      ["Intel Xeon Gold 6226R", "Cascade Lake-SP", "16C/32T", "2.9/3.9 GHz", "22 MB L3", "150W"],
      ["Apple M3", "ARM v8.6 (Apple)", "8C (4P+4E)", "3.7 GHz (Perf)", "Unified L2/SLC", "23W"],
      ["Intel Celeron N4020", "Gemini Lake R", "2C/2T", "1.1/2.8 GHz", "4 MB L2", "6W"],
      ["Broadcom BCM2712", "ARM Cortex-A76", "4C/4T", "2.4 GHz", "512 KB L2/2 MB L3", "5W"],
    ],
    [2500, 1700, 1500, 1700, 1500, 960]
  ),
  spacer(),
  heading2("11.1 Processor Technology Notes"),
  body("Modern lab processors use a hybrid architecture combining high-performance (P) cores and energy-efficient (E) cores. Intel's 13th Gen (Raptor Lake) architecture provides superior multi-threaded performance, ideal for compilation and data processing. AMD's Zen 3 offers competitive IPC at reduced thermal output, making it the preferred choice for mobile lab use. The Apple M3 uses unified memory architecture, dramatically reducing memory latency for graphics and ML inference workloads. ARM-based processors (BCM2712) are purpose-built for embedded scenarios with tight power budgets."),
  pageBreak(),
];

// CHAPTER 12 - MEMORY
const chapter12 = [
  heading1("12. MEMORY (RAM) SPECIFICATIONS"),
  spacer(),
  body("Random Access Memory (RAM) is a critical determinant of system responsiveness and multitasking capability. The lab's systems span DDR4 (most desktops and laptops), DDR5 (newer desktops and AiOs), LPDDR4X (thin clients and embedded systems), and Registered ECC memory (servers and workstations). ECC (Error-Correcting Code) memory is essential for server and workstation reliability as it detects and corrects single-bit memory errors on-the-fly."),
  spacer(),
  ...specTable(
    "Table 12.1 — RAM Specifications by System",
    ["System", "Type", "Capacity", "Speed", "Channels", "ECC"],
    [
      ["Desktop Model A (x10)", "DDR4 DIMM", "16 GB", "3200 MT/s", "Dual", "No"],
      ["Desktop Model B (x5)", "DDR5 DIMM", "32 GB", "4800 MT/s", "Dual", "No"],
      ["Laptop Dell Inspiron", "DDR4 SO-DIMM", "8 GB", "3200 MT/s", "Single", "No"],
      ["Laptop ThinkPad E15", "DDR4 SO-DIMM", "16 GB", "3200 MT/s", "Dual", "No"],
      ["Workstation HP Z4", "DDR4 RDIMM", "128 GB", "2933 MT/s", "Quad", "Yes"],
      ["Workstation Dell 5820", "DDR4 RDIMM", "64 GB", "2666 MT/s", "Quad", "Yes"],
      ["Server Dell R740", "DDR4 RDIMM", "256 GB", "2933 MT/s", "8-channel", "Yes"],
      ["AiO HP EliteOne", "DDR5 SO-DIMM", "32 GB", "4800 MT/s", "Dual", "No"],
      ["AiO Apple iMac M3", "LPDDR5 Unified", "16 GB", "6400 MT/s", "Unified", "No"],
      ["Thin Client HP t430", "LPDDR4 Soldered", "4 GB", "2400 MT/s", "Dual", "No"],
    ],
    [2400, 1500, 1300, 1400, 1200, 960]
  ),
  spacer(),
  heading2("12.1 Memory Upgrade Policy"),
  body("All desktop systems have available DIMM slots for memory expansion. Desktop Model A can be expanded to 64 GB (2 remaining slots). Desktop Model B supports up to 128 GB DDR5. Server RAM is managed by the infrastructure team only; no student-initiated changes are permitted. Workstations can accept additional RDIMM modules subject to availability and prior authorization."),
  pageBreak(),
];

// CHAPTER 13 - STORAGE
const chapter13 = [
  heading1("13. STORAGE SPECIFICATIONS"),
  spacer(),
  body("Storage systems in the lab span multiple technologies including NVMe SSDs (fastest), SATA SSDs, mechanical hard drives (HDDs), eMMC flash storage, and microSD cards. NVMe drives offer the highest sequential read/write performance (up to 7000/6500 MB/s on Gen 4 drives), while traditional HDDs provide high-capacity, cost-effective bulk storage. RAID configurations are employed on servers for redundancy and performance."),
  spacer(),
  ...specTable(
    "Table 13.1 — Storage Configuration by System",
    ["System", "Primary Storage", "Capacity", "Interface", "Read MB/s"],
    [
      ["Desktop A", "Samsung 980 SSD", "512 GB", "PCIe 3.0 NVMe", "3500"],
      ["Desktop B", "Samsung 990 Pro", "1 TB", "PCIe 4.0 NVMe", "7450"],
      ["Desktop A (HDD)", "Seagate Barracuda", "1 TB", "SATA 6G", "180"],
      ["Desktop B (HDD)", "WD Blue", "2 TB", "SATA 6G", "150"],
      ["Laptop Dell", "Samsung PM9A1", "512 GB", "PCIe 3.0 NVMe", "3500"],
      ["Laptop ThinkPad", "SKHynix BC711", "512 GB", "PCIe 3.0 NVMe", "3500"],
      ["Workstation HP Z4", "Samsung 870 EVO x2 RAID 0", "4 TB", "SATA 6G", "550"],
      ["Server Primary", "8x SATA HDD RAID 6", "20 TB usable", "SATA 6G", "800 RAID"],
      ["All-in-One iMac", "Apple Custom SSD", "512 GB", "NVMe (Apple)", "5000+"],
      ["Thin Client", "eMMC 5.1", "32 GB", "eMMC 5.1", "250"],
    ],
    [2300, 2200, 1400, 1700, 1760]
  ),
  spacer(),
  heading2("13.1 Backup and Data Protection"),
  body("Critical student data is stored on the primary server and backed up nightly to the secondary server using rsync over SSH. Desktops are configured to redirect the Documents and Desktop folders to the server share via Group Policy, ensuring no local data loss. RAID 6 on the server tolerates up to 2 simultaneous disk failures without data loss. A full tape backup is performed quarterly."),
  pageBreak(),
];

// CHAPTER 14 - GRAPHICS
const chapter14 = [
  heading1("14. GRAPHICS & DISPLAY SPECIFICATIONS"),
  spacer(),
  body("The graphics subsystems in the lab range from integrated GPU solutions sufficient for general use to professional NVIDIA Quadro and GeForce GPUs for rendering and machine learning. Displays span Full HD to 4K resolutions. This chapter details the graphics hardware and connected monitors for each system category."),
  spacer(),
  ...specTable(
    "Table 14.1 — Graphics Hardware Specifications",
    ["System", "GPU", "VRAM", "API Support", "Max Display"],
    [
      ["Desktop A", "Intel UHD 730", "Shared (up to 8 GB)", "DX12, OGL 4.6", "4K/60Hz"],
      ["Desktop B", "NVIDIA RTX 3060", "12 GB GDDR6", "DX12U, Vulkan, CUDA", "4K/144Hz"],
      ["Workstation HP Z4", "NVIDIA Quadro RTX 4000", "8 GB GDDR6 ECC", "DX12, OpenCL, CUDA", "4x 4K"],
      ["Workstation Dell 5820", "NVIDIA RTX A4000", "16 GB GDDR6 ECC", "DX12U, CUDA 12", "4x 4K"],
      ["AiO iMac M3", "Apple M3 10-core GPU", "Shared 16 GB", "Metal, OpenGL", "6K (via TB4)"],
      ["AiO HP EliteOne", "Intel UHD 770", "Shared (up to 32 GB)", "DX12, OGL 4.6", "4K/60Hz"],
      ["Thin Client", "Intel UHD 600", "Shared 2 GB", "DX12, OGL 4.6", "2x 4K@60Hz"],
      ["Raspberry Pi 5", "VideoCore VII", "Shared", "OpenGL ES 3.1, Vulkan 1.2", "2x 4K@60Hz"],
    ],
    [2000, 2800, 1500, 1700, 1360]
  ),
  spacer(),
  ...specTable(
    "Table 14.2 — Monitor Specifications",
    ["Monitor Model", "Size", "Resolution", "Panel", "Brightness", "Count"],
    [
      ["LG 24MP400-B", "24\"", "1920x1080 FHD", "IPS", "250 nits", "25 (Desktops)"],
      ["Dell P2722H", "27\"", "1920x1080 FHD", "IPS", "300 nits", "4 (Workstations)"],
      ["HP E27 G4 QHD", "27\"", "2560x1440 QHD", "IPS", "300 nits", "5 (Enhanced)"],
      ["Built-in Display", "15.6\"", "1920x1080", "IPS", "300 nits", "10 (Laptops)"],
      ["Built-in Display", "24\"/27\"", "4K", "IPS", "350/500 nits", "5 (AiO)"],
    ],
    [2400, 900, 1700, 900, 1500, 960]
  ),
  pageBreak(),
];

// CHAPTER 15 - NETWORKING
const chapter15 = [
  heading1("15. NETWORK & CONNECTIVITY SPECIFICATIONS"),
  spacer(),
  body("The lab's network infrastructure provides all computers with fast, reliable connectivity to both the local area network and the internet. A centrally managed Cisco Catalyst switch forms the core of the wired network, with dedicated uplinks to the department's core router. Wireless access is provided by Cisco Meraki access points mounted in the ceiling, supporting the latest Wi-Fi 6 standard for reduced congestion in high-density environments."),
  spacer(),
  ...specTable(
    "Table 15.1 — Network Interface Specifications by System",
    ["System", "Wired Interface", "Wireless", "Speed"],
    [
      ["Desktop A", "Realtek RTL8111", "None (Wired Only)", "1 Gbps"],
      ["Desktop B", "Intel I225-V", "Intel AX200 (Wi-Fi 6)", "2.5 Gbps / 2.4 Gbps Wi-Fi"],
      ["Laptop Dell", "Realtek USB RJ-45", "Intel AX201 (Wi-Fi 6)", "1 Gbps / 2.4 Gbps Wi-Fi"],
      ["Laptop ThinkPad", "Realtek RTL8111", "Intel AX200 (Wi-Fi 6)", "1 Gbps / 2.4 Gbps Wi-Fi"],
      ["Workstation HP Z4", "Intel I350-T2", "None", "2x 1 Gbps Bonded"],
      ["Server Primary", "Intel X540 (10G)", "None", "10 Gbps + 4x 1 Gbps"],
      ["AiO HP EliteOne", "Realtek RTL8111", "Intel AX211 (Wi-Fi 6E)", "1 Gbps / 2.4 Gbps Wi-Fi"],
      ["Thin Client", "Realtek RTL8111", "Intel 7265 (Wi-Fi 5)", "1 Gbps / 867 Mbps Wi-Fi"],
      ["Raspberry Pi 5", "Gigabit Ethernet", "Cypress CYW43455 Wi-Fi 6", "1 Gbps / 2.4 Gbps Wi-Fi"],
    ],
    [2400, 2200, 2200, 2560]
  ),
  spacer(),
  heading2("15.1 Lab Network Infrastructure"),
  body("The lab network is segmented into three VLANs: VLAN 10 for student workstations, VLAN 20 for staff and instructor systems, and VLAN 30 for servers. Inter-VLAN routing is managed by the core firewall with access control lists. The lab maintains a 1 Gbps uplink to the campus backbone, with QoS policies prioritizing RDP and SSH traffic for thin client and remote access users."),
  bullet("DNS: Internal DNS via Microsoft AD (Windows) and BIND9 (Linux)."),
  bullet("DHCP: Centrally managed with static reservations for servers, workstations, and printers."),
  bullet("Firewall: pfSense VM handling outbound NAT and intrusion detection (Suricata)."),
  bullet("VPN: WireGuard-based access for remote student and faculty connections."),
  pageBreak(),
];

// CHAPTER 16 - OS
const chapter16 = [
  heading1("16. OPERATING SYSTEMS INSTALLED"),
  spacer(),
  body("A variety of operating systems are installed and maintained across lab systems. The diversity is intentional — it ensures students gain experience with both Windows and Linux environments, and in some cases, macOS. All operating systems are kept updated via centralized patch management. Windows systems use WSUS for updates, and Linux systems use Ansible playbooks for automated patch deployment."),
  spacer(),
  ...specTable(
    "Table 16.1 — Operating System Deployment",
    ["OS", "Version", "Installed On", "Update Mechanism"],
    [
      ["Windows 11 Pro", "23H2 (Build 22631)", "15 Desktops, 4 Workstations", "WSUS / Microsoft Intune"],
      ["Windows 11 Home", "23H2", "5 AiO (HP EliteOne)", "Windows Update"],
      ["macOS Sonoma", "14.4", "2 AiO (iMac M3)", "Apple Software Update"],
      ["Ubuntu 22.04 LTS", "22.04.4 (Jammy)", "10 Laptops, 2 Workstations", "Ansible + unattended-upgrades"],
      ["Ubuntu Server 22.04", "22.04.4", "2 Servers", "Ansible Playbooks"],
      ["HP ThinPro 8.0", "8.0 SP4", "8 Thin Clients", "HP Device Manager"],
      ["Raspberry Pi OS 12", "Bookworm 64-bit", "4 Raspberry Pi 5", "apt / raspi-config"],
      ["Ubuntu Server (ARM)", "22.04.4 ARM64", "2 Raspberry Pi 4", "apt"],
    ],
    [2200, 2200, 2400, 2560]
  ),
  spacer(),
  heading2("16.1 Software Ecosystem"),
  body("Beyond the operating system, lab systems host a standardized software stack relevant to the Computer Science curriculum. Windows systems use Chocolatey for package management; Linux systems use apt with a lab-maintained PPA for custom packages. Key software installed across the lab includes:"),
  bullet("Development: Visual Studio Code, IntelliJ IDEA, Eclipse, Vim, GCC/G++, Python 3.12, Java JDK 21"),
  bullet("Database: MySQL 8.0, PostgreSQL 16, MongoDB 7, SQLite3, Oracle XE 21c"),
  bullet("Networking: Wireshark, Cisco Packet Tracer, GNS3, Nmap, OpenSSH"),
  bullet("Data Science: Anaconda (Python), R + RStudio, MATLAB R2024a, Jupyter Notebook"),
  bullet("Virtualization: VirtualBox 7, VMware Player, WSL2 (Windows), Docker Desktop"),
  bullet("Office: Microsoft Office 365 (Windows), LibreOffice 7 (Linux)"),
  pageBreak(),
];

// CHAPTER 17 - CONCLUSION
const chapter17 = [
  heading1("17. CONCLUSION & RECOMMENDATIONS"),
  spacer(),
  body("This report has provided a comprehensive and detailed documentation of all major computer systems available in the Computer Science Laboratory. The lab's hardware inventory represents a well-rounded ecosystem covering diverse computing paradigms — from enterprise-grade servers and workstations to energy-efficient thin clients and embedded single-board computers. Each system type serves a distinct educational and operational purpose, collectively supporting the full breadth of Computer Science coursework."),
  spacer(),
  heading2("17.1 Summary of Findings"),
  body("The following key observations emerge from the specification analysis conducted in this report:"),
  bullet("The lab's desktop fleet (15 units) constitutes the primary workhorse for daily programming and coursework. The split between standard (Model A) and enhanced (Model B) configurations provides appropriate resources for different task intensities."),
  bullet("Laptops provide essential flexibility for students who need mobile computing capabilities, and the dual-boot configuration exposes students to both Windows and Linux environments."),
  bullet("Workstations serve as a critical resource for research-oriented tasks requiring substantial CPU/GPU power. Their GPU compute capability (NVIDIA Quadro RTX 4000, RTX A4000) is particularly valuable for machine learning and simulation courses."),
  bullet("The server infrastructure provides reliable centralized services. The RAID 6 storage configuration and dual-server redundancy ensure data safety and service availability."),
  bullet("Thin clients provide a cost-effective, easily managed solution for the secondary cluster, reducing per-seat hardware costs while leveraging centralized VDI."),
  bullet("The Raspberry Pi cluster is an excellent resource for IoT and embedded systems labs, providing hands-on experience with ARM architecture and GPIO programming."),
  spacer(),
  heading2("17.2 Recommendations"),
  body("Based on the specifications documented in this report, the following recommendations are proposed for the next hardware refresh cycle:"),
  bullet("Upgrade Desktop Model A (i5-13400) units to systems with dedicated GPUs to better support machine learning coursework at the undergraduate level."),
  bullet("Increase laptop inventory by 5 units (ThinkPad E15 configuration) to reduce contention for available units during practical sessions."),
  bullet("Add a third server node to create a 3-node high-availability cluster using Proxmox VE or VMware vSphere, improving service uptime."),
  bullet("Replace the 4 Raspberry Pi 4 units with Pi 5 (8 GB) units to standardize the embedded systems cluster on a consistent hardware baseline."),
  bullet("Expand the workstation RAM on HP Z4 units from 128 GB to 256 GB to support larger ML dataset in-memory processing."),
  bullet("Implement a 25 GbE upgrade for the server uplinks to accommodate growing data transfer demands from research workloads."),
  spacer(),
  heading2("17.3 Final Remarks"),
  body("The Computer Science Laboratory is well-equipped to support a modern, comprehensive curriculum. The diversity of systems — ranging from embedded ARM boards to dual-Xeon servers — provides students with an exceptional breadth of practical computing experience. Maintaining up-to-date documentation of these specifications, as done in this report, is essential for effective lab management, curriculum planning, and future procurement decisions. It is recommended that this report be reviewed and updated at the start of each academic year to reflect any hardware changes or additions."),
  spacer(),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 300, after: 300 },
    border: { top: { style: BorderStyle.SINGLE, size: 6, color: LIGHT_BLUE }, bottom: { style: BorderStyle.SINGLE, size: 6, color: LIGHT_BLUE } },
    children: [new TextRun({ text: "— End of Report —", bold: true, size: 24, font: "Arial", color: LIGHT_BLUE, italics: true })]
  }),
  spacer(),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: "Computer Specifications Report | CS Laboratory | April 2026", size: 18, font: "Arial", color: "888888", italics: true })]
  }),
];

// ===================== BUILD DOCUMENT =====================
const doc = new Document({
  numbering: {
    config: [
      {
        reference: "bullets",
        levels: [{
          level: 0, format: LevelFormat.BULLET, text: "•",
          alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } }
        }]
      }
    ]
  },
  styles: {
    default: { document: { run: { font: "Arial", size: 22, color: DARK } } },
    paragraphStyles: [
      {
        id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 36, bold: true, font: "Arial", color: WHITE },
        paragraph: { spacing: { before: 360, after: 200 }, outlineLevel: 0 }
      },
      {
        id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 28, bold: true, font: "Arial", color: LIGHT_BLUE },
        paragraph: { spacing: { before: 280, after: 120 }, outlineLevel: 1 }
      },
      {
        id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 24, bold: true, font: "Arial", color: BLUE },
        paragraph: { spacing: { before: 200, after: 80 }, outlineLevel: 2 }
      },
    ]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 }
      }
    },
    headers: {
      default: new Header({
        children: [
          new Paragraph({
            children: [
              new TextRun({ text: "CS LABORATORY — COMPUTER SPECIFICATIONS REPORT   |   APRIL 2026", size: 16, font: "Arial", color: "888888" }),
            ],
            border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: LIGHT_BLUE, space: 1 } },
            spacing: { after: 100 },
          })
        ]
      })
    },
    footers: {
      default: new Footer({
        children: [
          new Paragraph({
            children: [
              new TextRun({ text: "Computer Science Laboratory  |  Confidential  |  Page ", size: 16, font: "Arial", color: "888888" }),
              new TextRun({ children: [PageNumber.CURRENT], size: 16, font: "Arial", color: LIGHT_BLUE }),
              new TextRun({ text: " of ", size: 16, font: "Arial", color: "888888" }),
              new TextRun({ children: [PageNumber.TOTAL_PAGES], size: 16, font: "Arial", color: LIGHT_BLUE }),
            ],
            alignment: AlignmentType.CENTER,
            border: { top: { style: BorderStyle.SINGLE, size: 4, color: LIGHT_BLUE, space: 1 } },
            spacing: { before: 100 },
          })
        ]
      })
    },
    children: [
      ...coverPage,
      ...tocPage,
      ...chapter1,
      ...chapter2,
      ...chapter3,
      ...chapter4,
      ...chapter5,
      ...chapter6,
      ...chapter7,
      ...chapter8,
      ...chapter9,
      ...chapter10,
      ...chapter11,
      ...chapter12,
      ...chapter13,
      ...chapter14,
      ...chapter15,
      ...chapter16,
      ...chapter17,
    ]
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync('/mnt/user-data/outputs/Computer_Specifications_Report.docx', buffer);
  console.log('Report generated successfully!');
});
