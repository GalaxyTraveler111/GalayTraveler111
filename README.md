pip install python-docx
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT

# Create document
doc = Document()

# ---------------------------
# Helper functions
# ---------------------------

def add_heading(text, size=14):
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.bold = True
    run.font.size = Pt(size)
    p.paragraph_format.space_after = Pt(4)
    return p

def add_section_title(text):
    p = doc.add_paragraph()
    run = p.add_run(text.upper())
    run.bold = True
    run.font.size = Pt(13)
    p.paragraph_format.space_after = Pt(2)
    return p

def add_bullet(text):
    p = doc.add_paragraph(style='List Bullet')
    run = p.add_run(text)
    run.font.size = Pt(11)
    p.paragraph_format.space_after = Pt(1)

def add_paragraph(text, size=11):
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.font.size = Pt(size)
    p.paragraph_format.space_after = Pt(4)
    return p

# ---------------------------
# Header
# ---------------------------

name = doc.add_paragraph()
name_run = name.add_run("RAYON ROSE")
name_run.bold = True
name_run.font.size = Pt(22)
name.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT

contact = doc.add_paragraph()
contact_run = contact.add_run(
    "Tampa, FL • 813-580-8004 • timeneverwaits17@gmail.com\n"
    "IT Support • Infrastructure • Cloud • DevOps (In Progress)"
)
contact_run.font.size = Pt(11)

# ---------------------------
# Profile
# ---------------------------

add_section_title("Profile")

add_paragraph(
    "IT Support professional with experience across hardware, software, network, VOIP, "
    "and system troubleshooting in high-volume environments. Currently expanding skills "
    "in Infrastructure and DevOps with a focus on automation and scalable systems. "
    "Strong communicator who thrives under pressure and supports cross-functional teams. "
    "Certified in A+, Network+, Azure, AWS, and Cloud technologies. Seeking a modern, "
    "growth-oriented technical environment."
)

# ---------------------------
# Technical Skills
# ---------------------------

add_section_title("Technical Skills")

add_paragraph(
    "Infrastructure & Systems: Windows 10/11, Windows Server (Basics), Active Directory, "
    "Azure AD, M365, Device Imaging, PC Setup, Configuration, VOIP, AV Systems"
)

add_paragraph(
    "Networking: TCP/IP, DNS, DHCP, VPN, VLAN (Basics), Network Diagnostics"
)

add_paragraph(
    "Cloud: Azure Essentials, AWS Cloud, IaaS / PaaS / SaaS"
)

add_paragraph(
    "Tools & Platforms: Ticketing Systems, SharePoint, Microsoft Office, Intune (Familiar), "
    "Autopilot (Familiar)"
)

add_paragraph(
    "Strengths: Documentation, Fraud Prevention, Escalation Handling, Project Coordination"
)

# ---------------------------
# Experience
# ---------------------------

add_section_title("Experience")

# Progressive
add_heading("Progressive Insurance — Customer Relationship Manager (CRM)", 12)
add_paragraph("Mar 2025 – Present", size=10)

add_bullet("Troubleshoot access, login, and system issues for remote and onsite associates.")
add_bullet("Deliver clear, efficient technical support with minimal downtime.")
add_bullet("Document issues and resolutions for audits and escalations.")
add_bullet("Partner with IT teams during outages and workflow disruptions.")
add_bullet("Maintain compliance with authentication and security standards.")

# USAA
add_heading("USAA — Relationship Specialist", 12)
add_paragraph("Mar 2014 – Jul 2019", size=10)

add_bullet("Resolved access, authentication, and system-related issues across secure platforms.")
add_bullet("Guided users through troubleshooting for software, account, and device issues.")
add_bullet("Ensured adherence to security, backup, and fraud-prevention protocols.")
add_bullet("Documented technical issues to support escalations and process improvements.")
add_bullet("Recognized for leveraging IT skills to prevent fraud and support team operations.")

# Verizon
add_heading("Verizon — Technical Support Specialist", 12)
add_paragraph("Nov 2009 – Mar 2014", size=10)

add_bullet("Diagnosed hardware, software, network, and connectivity issues for a large user base.")
add_bullet("Troubleshot Windows systems, mobile devices, VOIP, and network problems.")
add_bullet("Guided users through OS configuration, installs, and device resets.")
add_bullet("Documented incidents to support escalations and trend analysis.")

# ---------------------------
# Certifications
# ---------------------------

add_section_title("Certifications")

add_paragraph(
    "CompTIA A+, CompTIA Network+, Certified Cloud Professional, "
    "Microsoft Azure Essentials, AWS Cloud (Foundational)"
)

# ---------------------------
# Education
# ---------------------------

add_section_title("Education")

add_paragraph("Hillsborough Community College — Information Technology")
add_paragraph("Tampa Bay Technical High School — High School Diploma")

# ---------------------------
# Save Document
# ---------------------------

doc.save("RayonRoseIT.docx")

print("Resume generated successfully: RayonRoseIT.docx")
