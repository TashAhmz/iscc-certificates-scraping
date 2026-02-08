from scrape import scrape_all
from datetime import datetime
from styles import apply_styles
import json
from compare import create_certs_added, create_certs_removed, create_certs_changed

# Scrape configuration
DELAY = 5
ROWS_LOADED = 20000

now = datetime.now()
timestamp = now.strftime("%d.%m.%Y_%H.%M")
filename = f"ISCC_Certificates_{timestamp}.xlsx"
output_file = f"out/{filename}"


if __name__ == "__main__":
    scrape_all(delay=DELAY, page_size=ROWS_LOADED, output_file=output_file)
    apply_styles(output_file, "Certificate Database")

    try: 
        with open("src/utils.json", "r") as f1: 
            data = json.load(f1)
    except json.JSONDecodeError:
        print(f"{f1.name} is empty")

    prev_filename = data.get("prev_file_name")
    
    if prev_filename:
        create_certs_added(prev_filename, output_file)
        print()
        create_certs_removed(prev_filename, output_file)
        print()
        create_certs_changed(prev_filename, output_file, ignore_cols=["Map", "Company_Name", "City"])
        print()

        apply_styles(output_file, "Certificates Added")
        apply_styles(output_file, "Certificates Removed")
        apply_styles(output_file, "Certificates Changed")

    try: 
        with open("src/utils.json", "w") as f2:
            json.dump({"prev_file_name": f"{output_file}"}, f2)
    except json.JSONDecodeError:
        print(f"{f2.name} is empty")