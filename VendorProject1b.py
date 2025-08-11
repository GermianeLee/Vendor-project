import datetime
import pandas as pd
from pathlib import Path
from docxtpl import DocxTemplate  # pip install docxtpl

# Define paths
base_dir = Path(r"C:\Users\GongLee73\PycharmProjects\PythonProject")
gen_dir = base_dir / "GeneratedDocs"
word_template_path = base_dir / "VendorContract.docx"

# Create the output directory if it doesn't exist
gen_dir.mkdir(parents=True, exist_ok=True)

# Load the Excel data (Make sure your Excel file is at this location)
excel_file_path = base_dir / "data.xlsx"
df = pd.read_excel(excel_file_path)

# Loop through each row in the DataFrame and generate a contract document for each
for index, row in df.iterrows():
    # Calculate "NONREFUNDABLE" as 20% of "AMOUNT"
    nonrefundable = round(row['AMOUNT'] * 0.2, 2)

    # Get today's date and the date one week from today
    today = datetime.datetime.today()
    today_in_one_week = today + datetime.timedelta(days=7)

    # Define the context for filling the template
    context = {
        "CLIENT": row["CLIENT"],
        "VENDOR": row["VENDOR"],
        "LINE1": row["LINE1"],
        "LINE2": row["LINE2"],
        "AMOUNT": row["AMOUNT"],
        "NONREFUNDABLE": nonrefundable,
        "TODAY": today.strftime("%Y-%m-%d"),
        "TODAY_IN_ONE_WEEK": today_in_one_week.strftime("%Y-%m-%d"),
    }

    # Load the Word template
    doc = DocxTemplate(word_template_path)

    # Render the document with the current context
    doc.render(context)

    # Generate a unique output file name (using the index number in the DataFrame)
    output_file_name = f"GeneratedContract_{index + 1}.docx"
    output_path = gen_dir / output_file_name

    # Save the generated document
    doc.save(output_path)

    print(f"Generated contract saved at: {output_path}")
