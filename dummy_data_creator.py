import random
from faker import Faker
from fpdf import FPDF
from docx import Document
from openpyxl import Workbook

# Initialize Faker
fake = Faker()

def generate_timeshare_data(complete=True):
    data = {
        "client_name": fake.name(),
        "email": fake.email(),
        "phone": fake.phone_number(),
        "location": fake.city(),
        "sales_rep": fake.name(),
        "interest_level": random.choice(["High", "Medium", "Low"]),
        "notes": fake.text()
    }
    if not complete:
        for _ in range(random.randint(1, 3)):
            key = random.choice(list(data.keys()))
            del data[key]
    return data

def create_pdf(data, filename):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=12)
    for key, value in data.items():
        pdf.cell(200, 10, txt=f"{key}: {value}", ln=True)
    pdf.output(filename)

def create_word(data, filename):
    doc = Document()
    for key, value in data.items():
        doc.add_paragraph(f"{key}: {value}")
    doc.save(filename)

def create_excel(data, filename):
    workbook = Workbook()
    sheet = workbook.active
    for i, (key, value) in enumerate(data.items(), start=1):
        sheet.cell(row=i, column=1, value=key)
        sheet.cell(row=i, column=2, value=value)
    workbook.save(filename)

def generate_documents(num_docs):
    for i in range(num_docs):
        complete = random.random() > 0.05  # 95% chance to be complete
        data = generate_timeshare_data(complete)
        doc_type = random.choice(["pdf", "docx", "xlsx"])
        filename = f"timeshare_{i}.{doc_type}"
        if doc_type == "pdf":
            create_pdf(data, filename)
        elif doc_type == "docx":
            create_word(data, filename)
        elif doc_type == "xlsx":
            create_excel(data, filename)

generate_documents(1000)
