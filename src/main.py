from scrape import scrape_all
from datetime import datetime
from styles import apply_styles
import json
from compare import create_certs_added, create_certs_removed, create_certs_changed

# Scrape configuration
DELAY = 0.01
ROWS_LOADED = 200

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
        create_certs_changed(prev_filename, output_file, ignore_cols=["Map", "Company_Name", "City", "Asset_Identifier", "Match_Found", "Suggested_Asset_Identifier", "Matcher_Version", "Match_Status", "Match_Confidence", "Match_Method",
"Auto_Match_Eligible", "Is_Processing_Unit", "Matched_GST_Company", "Best_GST_Company_Candidate", "Company_Match_Method", "Company_Match_Evidence", "Matched_Territory", "Company_Score", "Company_Score_Margin",
"Asset_Company_Score", "Asset_Company_Match_Method", "Asset_Company_Match_Evidence", "Location_Score", "Location_Match_Method", "Overall_Score", "Score_Margin", "Candidate_Count", "Candidate_Site_Count", "Matched_Site_Key",
"Runner_Up_Asset", "Runner_Up_Score", "Review_Reason"])
        
        print()

        apply_styles(output_file, "Certificates Added")
        apply_styles(output_file, "Certificates Removed")
        apply_styles(output_file, "Certificates Changed")

    try: 
        with open("src/utils.json", "w") as f2:
            json.dump({"prev_file_name": f"{output_file}"}, f2)
    except json.JSONDecodeError:
        print(f"{f2.name} is empty")
