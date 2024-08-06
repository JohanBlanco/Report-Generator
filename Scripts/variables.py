from datetime import datetime
year = str(datetime.now().year)
files_directory_path = '../Files'
template_path = '../Template/Lilly Kanban Management Template.xlsx'
new_file_name = 'Lilly Kanban Management'
reports_path = '../Report History'
lattest_report_path = '../Lattest Report'
log_file_path = '../result.log'
other_PTO_dates = [f"12/26/{year}",f"12/27/{year}",f"12/28/{year}",
                   f"12/29/{year}",f"12/30/{year}",f"12/31/{year}",
                   f"01/01/{year}",f"01/02/{year}"]