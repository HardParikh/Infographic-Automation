import os
from utils import save_to_excel, replace_table, remove_trailing_zeros, create_folder_structure, save_selection_as_pdf, send_email, update_workbook, get_previous_month_year, get_personalized_query
import shutil

BASE_PATH = r"C:\Users\ParikH01\CPGPLC\E15 FLIK LENZ - E15 Channel - Analytics\Python Automation - LENZ Performance Snapshot"



def data_fetching():
    # Fetch the dynamically created folder path for the previous month
    previous_month_folder = create_folder_structure(BASE_PATH, "Hyperlink")

    individual = False
    
    # Define query files, excel files, and destination paths
    query_and_excel_files = [
        {
            'query_files': {
                'Aggregated Monthly Infographic': '../queries/aggregated_audits/aggregated_monthly_infographic.sql',
            },
            'excel_file': '../data/Aggregated Monthly Infographic - Hyperlinks.xlsx',
            'destination': os.path.join(previous_month_folder, 'Aggregated Monthly Infographic - Hyperlinks.xlsx')
        },
        {
            'query_files': {
                'Monthly Infographic': '../queries/monthly_infographic_audits/monthly_infographic.sql',
            },
            'excel_file': '../data/Monthly Infographic - Hyperlinks.xlsx',
            'destination': os.path.join(previous_month_folder, 'Monthly Infographic - Hyperlinks.xlsx')
        },
        {
            'query_files': {
                'Failed Audits': '../queries/health_and_sanitation_audits/qa_failed_audits.sql',
                'Failed Audits-Closed Tickets': '../queries/health_and_sanitation_audits/qa_failed_audits_closed_tickets.sql',
                'Failed Audits-Critical Q Missed': '../queries/health_and_sanitation_audits/qa_failed_audits_critical_missed.sql',
                'All Audits-Open Tickets': '../queries/health_and_sanitation_audits/qa_all_audits_open_tickets.sql',
            },
            'excel_file': '../data/QA Monthly Infographic - Hyperlinks.xlsx',
            'destination': os.path.join(previous_month_folder, 'QA Monthly Infographic - Hyperlinks.xlsx')
        },
        {
            'query_files': {
                'Unit Overview': '../queries/voc_audits/voc_unit_overview.sql',
                'AI Identified Negative Comments': '../queries/voc_audits/voc_negative_comments.sql',
            },
            'excel_file': '../data/VOC Monthly Infographic - Hyperlinks.xlsx',
            'destination': os.path.join(previous_month_folder, 'VOC Monthly Infographic - Hyperlinks.xlsx')
        }
    ]

    # Loop through each set of query files and destinations
    for item in query_and_excel_files:
        # Save to the main "data" folder
        save_to_excel(item['excel_file'], item['query_files'], individual)
        # Save to the dynamic previous month-year folder
        save_to_excel(item['destination'], item['query_files'], individual)


    # Define the names of individuals
    sddr_names = ['Adam Salem', 'Brian Donohue', 'Neil Gardner', 'Peter Soguero', 'Rick Russo', 'Sari Feltman']

    # Define the folder where SQL templates are stored
    qa_query_template_folder = '../queries/health_and_sanitation_audits_individual_sddr'
    voc_query_template_folder = '../queries/voc_audits_individual_sddr'

    # QA query mappings and generation
    qa_sheet_name_mappings = {
        "qa_failed_audits.sql": "Failed Audits",
        "qa_failed_audits_closed_tickets.sql": "Failed Audits-Closed Tickets",
        "qa_failed_audits_critical_missed.sql": "Failed Audits-Critical Q Missed",
        "qa_all_audits_open_tickets.sql": "All Audits-Open Tickets"
    }

    individual = True
    
    for sddr_name in sddr_names:
        formatted_sddr_name = sddr_name.replace(" ", "_").lower()
        
        # Define the output paths
        excel_file_path = f"../data/QA Monthly Infographic - Hyperlinks_{formatted_sddr_name}.xlsx"
        destination_path = os.path.join(previous_month_folder, f"QA Monthly Infographic - Hyperlinks_{formatted_sddr_name}.xlsx")

        # Dictionary to hold personalized queries for each sheet
        query_files = {}
        for query_file_name, sheet_name in qa_sheet_name_mappings.items():
            query_template_file = os.path.join(qa_query_template_folder, query_file_name)
            if os.path.exists(query_template_file):
                personalized_query = get_personalized_query(query_template_file, sddr_name)
                query_files[sheet_name] = personalized_query

        # Save data for each query to separate sheets
        save_to_excel(excel_file_path, query_files, individual)        # Save in data folder
        save_to_excel(destination_path, query_files, individual)       # Save in previous month folder

    # VOC query mappings and generation
    voc_sheet_name_mappings = {
        "voc_unit_overview.sql": "Unit Overview",
        "voc_negative_comments.sql": "AI Identified Negative Comments"
    }

    for sddr_name in sddr_names:
        formatted_sddr_name = sddr_name.replace(" ", "_").lower()
        
        # Define the output paths
        excel_file_path = f"../data/VOC Monthly Infographic - Hyperlinks_{formatted_sddr_name}.xlsx"
        destination_path = os.path.join(previous_month_folder, f"VOC Monthly Infographic - Hyperlinks_{formatted_sddr_name}.xlsx")

        # Dictionary to hold personalized queries for each sheet
        query_files = {}
        for query_file_name, sheet_name in voc_sheet_name_mappings.items():
            query_template_file = os.path.join(voc_query_template_folder, query_file_name)
            if os.path.exists(query_template_file):
                personalized_query = get_personalized_query(query_template_file, sddr_name)
                query_files[sheet_name] = personalized_query

        # Save data for each query to separate sheets
        save_to_excel(excel_file_path, query_files, individual)        # Save in data folder
        save_to_excel(destination_path, query_files, individual)       # Save in previous month folder

 

def updating_tables():
    source_file = '../data/Aggregated Monthly Infographic - Hyperlinks.xlsx'
    source_sheet_name = 'Aggregated Monthly Infographic'
    remove_trailing_zeros(source_file, source_sheet_name, column_name = 'STATISTIC')
    source_table_range = (2, 1, 55, 5)  # Range of cells (min_row, min_col, max_row, max_col)
    target_file = '../reports/Monthly Infographic_Job Aid & Templates.xlsx'
    target_sheet_name = 'reference sheet_all SDDR'
    target_start_cell = 'B7'
    success = replace_table(source_file, source_sheet_name, source_table_range, target_file, target_sheet_name, target_start_cell)
    if success == -1:
        return "FAIL"
    source_file = '../data/Monthly Infographic - Hyperlinks.xlsx'
    source_sheet_name = 'Monthly Infographic'
    remove_trailing_zeros(source_file, source_sheet_name, column_name = 'STATISTIC')
    source_table_range = (2, 1, 325, 6)  # Range of cells (min_row, min_col, max_row, max_col)
    target_file = '../reports/Monthly Infographic_Job Aid & Templates.xlsx'
    target_sheet_name = 'reference sheet_indiv SDDR'
    target_start_cell = 'B7'
    success = replace_table(source_file, source_sheet_name, source_table_range, target_file, target_sheet_name, target_start_cell)
    if success == -1:
        return "FAIL"
    return "SUCCESS"

def save_snapshot_as_pdf():
    month, year = get_previous_month_year()
    
    excel_file = '../reports/Monthly Infographic_Job Aid & Templates.xlsx'
    sheet_name = "All SDDR"
    selection_range = "B2:L54" 
    output_pdf = f'../snapshots/{month}_{year}_Monthly Infographic_Snapshot_All_SDDR.pdf'

    save_selection_as_pdf(excel_file, sheet_name, selection_range, output_pdf)

    excel_file = '../reports/Monthly Infographic_Job Aid & Templates.xlsx'
    sheet_name = "Salem"
    selection_range = "B2:M57" 
    output_pdf = f'../snapshots/{month}_{year}_Monthly Infographic_Snapshot_Adam_Salem.pdf'

    save_selection_as_pdf(excel_file, sheet_name, selection_range, output_pdf)

    excel_file = '../reports/Monthly Infographic_Job Aid & Templates.xlsx'
    sheet_name = "Donohue"
    selection_range = "B2:M57" 
    output_pdf = f'../snapshots/{month}_{year}_Monthly Infographic_Snapshot_Brian_Donohue.pdf'

    save_selection_as_pdf(excel_file, sheet_name, selection_range, output_pdf)

    excel_file = '../reports/Monthly Infographic_Job Aid & Templates.xlsx'
    sheet_name = "Gardner"
    selection_range = "B2:M57" 
    output_pdf = f'../snapshots/{month}_{year}_Monthly Infographic_Snapshot_Neil_Gardner.pdf'

    save_selection_as_pdf(excel_file, sheet_name, selection_range, output_pdf)
    
    excel_file = '../reports/Monthly Infographic_Job Aid & Templates.xlsx'
    sheet_name = "Soguero"
    selection_range = "B2:M57" 
    output_pdf = f'../snapshots/{month}_{year}_Monthly Infographic_Snapshot_Peter_Soguero.pdf'

    save_selection_as_pdf(excel_file, sheet_name, selection_range, output_pdf)
    
    excel_file = '../reports/Monthly Infographic_Job Aid & Templates.xlsx'
    sheet_name = "Russo"
    selection_range = "B2:M57" 
    output_pdf = f'../snapshots/{month}_{year}_Monthly Infographic_Snapshot_Rick_Russo.pdf'

    save_selection_as_pdf(excel_file, sheet_name, selection_range, output_pdf)
    
    excel_file = '../reports/Monthly Infographic_Job Aid & Templates.xlsx'
    sheet_name = "Feltman"
    selection_range = "B2:M57" 
    output_pdf = f'../snapshots/{month}_{year}_Monthly Infographic_Snapshot_Sari_Feltman.pdf'

    save_selection_as_pdf(excel_file, sheet_name, selection_range, output_pdf)
    
def send_email_with_attachment():   
    # # Email configuration
    # SMTP_SERVER = 'smtp.gmail.com'  # SMTP server for Outlook
    # SMTP_PORT = 587                   # Port for TLS
    # SENDER_EMAIL = 'hpari002@ucr.edu'
    # SENDER_PASSWORD = 'zrrw qxmb ueki hlei'

    # # Define email content
    # SUBJECT = 'Monthly Infographic Report'
    # BODY = 'Dear team,\n\nPlease find attached the latest monthly infographic report.\n\nBest regards,\nHard Parikh'
    # PDF_PATH = ''

    # # Recipients filtered by domain(s)
    # RECIPIENTS = [
    #     'hpari002@ucr.edu',
    #     'Hard.Parikh@compass-usa.com',
    #     # 'jrickman@e15group.com'
    # ]

    # # Filter recipients by allowed extensions
    # ALLOWED_EXTENSIONS = ["@compass-usa.com", "@ucr.edu", "@gmail.com", "@e15group.com", ]
    # filtered_recipients = [email for email in RECIPIENTS if any(email.endswith(ext) for ext in ALLOWED_EXTENSIONS)]

    # send_email(SUBJECT, BODY, SENDER_EMAIL, SENDER_PASSWORD, filtered_recipients, PDF_PATH, SMTP_SERVER, SMTP_PORT)
    
    subject = 'Monthly Infographic Report'
    body = 'Dear team,\n\nPlease find attached the latest monthly infographic report.\n\nBest regards,\nHard Parikh'
    sender_email = 'hard.parikh@compass-usa.com'
    sender_password = 'YOUR_PASSWORD'  # Replace with a secure way to store this
    recipients = ['hpari002@ucr.edu', 'Hard.Parikh@compass-usa.com']
    attachment_path = r'C:\path\to\your\attachment.pdf'

    send_email(subject, body, sender_email, sender_password, recipients, attachment_path)
    
def update_hyperlinks():
    original_file_path_qa = os.path.join(BASE_PATH, "Links", "QA_HL.xlsx")
    source_file_path_qa = "../data/QA Monthly Infographic - Hyperlinks.xlsx"
    update_workbook(original_file_path_qa, source_file_path_qa)
    
    original_file_path_voc = os.path.join(BASE_PATH, "Links", "VOC_AI_HL.xlsx")
    source_file_path_voc = "../data/VOC Monthly Infographic - Hyperlinks.xlsx"
    update_workbook(original_file_path_voc, source_file_path_voc)
    
    original_file_path_qa = os.path.join(BASE_PATH, "Links", "QA_AS.xlsx")
    source_file_path_qa = "../data/QA Monthly Infographic - Hyperlinks_adam_salem.xlsx"
    update_workbook(original_file_path_qa, source_file_path_qa)
    
    original_file_path_voc = os.path.join(BASE_PATH, "Links", "VOC_AI_AS.xlsx")
    source_file_path_voc = "../data/VOC Monthly Infographic - Hyperlinks_adam_salem.xlsx"
    update_workbook(original_file_path_voc, source_file_path_voc)

    original_file_path_qa = os.path.join(BASE_PATH, "Links", "QA_BD.xlsx")
    source_file_path_qa = "../data/QA Monthly Infographic - Hyperlinks_brian_donohue.xlsx"
    update_workbook(original_file_path_qa, source_file_path_qa)
    
    original_file_path_voc = os.path.join(BASE_PATH, "Links", "VOC_AI_BD.xlsx")
    source_file_path_voc = "../data/VOC Monthly Infographic - Hyperlinks_brian_donohue.xlsx"
    update_workbook(original_file_path_voc, source_file_path_voc)
    
    original_file_path_qa = os.path.join(BASE_PATH, "Links", "QA_PS.xlsx")
    source_file_path_qa = "../data/QA Monthly Infographic - Hyperlinks_peter_soguero.xlsx"
    update_workbook(original_file_path_qa, source_file_path_qa)
    
    original_file_path_voc = os.path.join(BASE_PATH, "Links", "VOC_AI_PS.xlsx")
    source_file_path_voc = "../data/VOC Monthly Infographic - Hyperlinks_peter_soguero.xlsx"
    update_workbook(original_file_path_voc, source_file_path_voc)
    
    original_file_path_qa = os.path.join(BASE_PATH, "Links", "QA_NG.xlsx")
    source_file_path_qa = "../data/QA Monthly Infographic - Hyperlinks_neil_gardner.xlsx"
    update_workbook(original_file_path_qa, source_file_path_qa)
    
    original_file_path_voc = os.path.join(BASE_PATH, "Links", "VOC_AI_NG.xlsx")
    source_file_path_voc = "../data/VOC Monthly Infographic - Hyperlinks_neil_gardner.xlsx"
    update_workbook(original_file_path_voc, source_file_path_voc)
    
    original_file_path_qa = os.path.join(BASE_PATH, "Links", "QA_RR.xlsx")
    source_file_path_qa = "../data/QA Monthly Infographic - Hyperlinks_rick_russo.xlsx"
    update_workbook(original_file_path_qa, source_file_path_qa)
    
    original_file_path_voc = os.path.join(BASE_PATH, "Links", "VOC_AI_RR.xlsx")
    source_file_path_voc = "../data/VOC Monthly Infographic - Hyperlinks_rick_russo.xlsx"
    update_workbook(original_file_path_voc, source_file_path_voc)
    
    original_file_path_qa = os.path.join(BASE_PATH, "Links", "QA_SF.xlsx")
    source_file_path_qa = "../data/QA Monthly Infographic - Hyperlinks_sari_feltman.xlsx"
    update_workbook(original_file_path_qa, source_file_path_qa)
    
    original_file_path_voc = os.path.join(BASE_PATH, "Links", "VOC_AI_SF.xlsx")
    source_file_path_voc = "../data/VOC Monthly Infographic - Hyperlinks_sari_feltman.xlsx"
    update_workbook(original_file_path_voc, source_file_path_voc)
    
def save_reports_to_sharepoint():
    destination_folder_name = "Reports"
    source_report_path = '../reports/Monthly Infographic_Job Aid & Templates.xlsx'

    destination_folder_path = create_folder_structure(BASE_PATH, destination_folder_name)

    # Define the destination path for the report file
    destination_report_path = os.path.join(destination_folder_path, os.path.basename(source_report_path))

    # Copy the report to the destination folder
    try:
        shutil.copy2(source_report_path, destination_report_path)
        print(f"Report saved successfully at {destination_report_path}")
    except Exception as e:
        print(f"Error saving report: {e}")
        
def save_snapshots_to_sharepoint():
    # Save all snapshots in the "Snapshots" folder
    snapshots_folder_name = "Snapshots"
    snapshots_folder_path = create_folder_structure(BASE_PATH, snapshots_folder_name)

    source_snapshots_dir = '../snapshots'
    if os.path.exists(source_snapshots_dir):
        for snapshot_file in os.listdir(source_snapshots_dir):
            source_snapshot_path = os.path.join(source_snapshots_dir, snapshot_file)
            destination_snapshot_path = os.path.join(snapshots_folder_path, snapshot_file)

            try:
                shutil.copy2(source_snapshot_path, destination_snapshot_path)
                print(f"Snapshot '{snapshot_file}' saved successfully at {destination_snapshot_path}")
            except Exception as e:
                print(f"Error saving snapshot '{snapshot_file}': {e}")
    else:
        print(f"Snapshots folder not found at '{source_snapshots_dir}'")
    
if __name__ == "__main__":
    data_fetching()
    if updating_tables() == "FAIL":
        print("An unexpected error has been encountered while updating the tables. Please check and run the code again.")
    else:
        update_hyperlinks()
        save_snapshot_as_pdf()
        save_reports_to_sharepoint()
        save_snapshots_to_sharepoint()
        # send_email_with_attachment()
