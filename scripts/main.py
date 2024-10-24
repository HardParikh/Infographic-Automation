from utils import save_to_excel, replace_table, remove_trailing_zeros


def data_fetching():
    # Aggregated Monthly Infographic
    aggregated_query_files = {
        'Aggregated Monthly Infographic': '../queries/aggregated_audits/aggregated_monthly_infographic.sql',
    }
    aggregated_excel_file = '../data/Aggregated Monthly Infographic - Hyperlinks.xlsx'
    save_to_excel(aggregated_excel_file, aggregated_query_files)

    # Monthly Infographic Audits
    monthly_query_files = {
        'Monthly Infographic': '../queries/monthly_infographic_audits/monthly_infographic.sql',
    }
    monthly_excel_file = '../data/Monthly Infographic - Hyperlinks.xlsx'
    save_to_excel(monthly_excel_file, monthly_query_files)

    # Health and Sanitation Audits
    health_sanitation_query_files = {
        'Failed Audits': '../queries/health_and_sanitation_audits/qa_failed_audits.sql',
        'Failed Audits-Closed Tickets': '../queries/health_and_sanitation_audits/qa_failed_audits_closed_tickets.sql',
        'Failed Audits-Critical Q Missed': '../queries/health_and_sanitation_audits/qa_failed_audits_critical_missed.sql',
        'All Audits-Open Tickets': '../queries/health_and_sanitation_audits/qa_all_audits_open_tickets.sql',
    }
    health_sanitation_excel_file = '../data/QA Monthly Infographic - Hyperlinks.xlsx'
    save_to_excel(health_sanitation_excel_file, health_sanitation_query_files)

    # VOC Monthly Infographic Audits
    voc_monthly_query_files = {
        'Unit Overview': '../queries/voc_audits/voc_unit_overview.sql',
        'AI Identified Negative Comments': '../queries/voc_audits/voc_negative_comments.sql',
    }
    
    voc_monthly_excel_file = '../data/VOC Monthly Infographic - Hyperlinks.xlsx'
    save_to_excel(voc_monthly_excel_file, voc_monthly_query_files)
 

def updating_tables():
    source_file = '../data/Aggregated Monthly Infographic - Hyperlinks.xlsx'
    source_sheet_name = 'Aggregated Monthly Infographic'
    remove_trailing_zeros(source_file, source_sheet_name, column_name = 'STATISTIC')
    source_table_range = (2, 1, 55, 5)  # Range of cells (min_row, min_col, max_row, max_col)
    target_file = '../reports/10.08.2024_Monthly Infographic_Job Aid & Templates.xlsx'
    target_sheet_name = 'reference sheet_all SDDR'
    target_start_cell = 'B7'
    replace_table(source_file, source_sheet_name, source_table_range, target_file, target_sheet_name, target_start_cell)
    
if __name__ == "__main__":
    data_fetching()
    updating_tables()
