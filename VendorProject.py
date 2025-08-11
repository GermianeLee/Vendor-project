import datetime
from pathlib import Path
from docxtpl import DocxTemplate  # pip install docxtpl

# Define paths
base_dir = Path(r"C:\Users\GongLee73\PycharmProjects\PythonProject")  # Fixed path string
gen_dir = base_dir / "GeneratedDocs"
word_template_path = base_dir / "VendorContract.docx"

# Create the output directory if it doesn't exist
gen_dir.mkdir(parents=True, exist_ok=True)

# Define the dates
today = datetime.datetime.today()
today_in_one_week = today + datetime.timedelta(days=7)

# Define the context
context = {
    "CLIENT": "Lee",
    "VENDOR": "Palladium",
    "LINE1": "Now You See Me",
    "LINE2": "Now You Dont",
    "AMOUNT": 1234,
    "NONREFUNDABLE": round(1234 * 0.2, 2),
    "TODAY": today.strftime("%Y-%m-%d"),
    "TODAY_IN_ONE_WEEK": today_in_one_week.strftime("%Y-%m-%d"),
}

# Load and render the template
doc = DocxTemplate(word_template_path)
doc.render(context)

# Save the generated document
output_path = gen_dir / "GeneratedContract.docx"
doc.save(output_path)

print(f"Document generated and saved at: {output_path}")
