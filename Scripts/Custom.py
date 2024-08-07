from enum import Enum
from typing import Callable, List
from datetime import datetime
import re
import pandas as pd
from workalendar.usa import UnitedStates
from Logger import Logger
from variables import *
from Utils import *

class desc_keys(Enum):
    ConfigStartDate = "config start date"
    AllConfigBlockedStartDate = "all config blocked start date"
    AllConfigBlockedEndDate = "all config blocked end date"
    ConfigEndDate = "config end date"
    AllPeerReviewStartDate = "all peer review start date"
    AllPeerReviewEndDate = "all peer review end date"
    AllPeerReviewReworkRequiredStartDate = "all peer review rework required start date"
    AllPRReworkBlockedStartDate = "all pr rework blocked start date"
    AllPRReworkBlockedEndDate = "all pr rework blocked end date"
    AllPeerReviewReworkRequiredEndDate = "all peer review rework required end date"
    AllReadyForDemoDate = "all ready for demo date"
    AllDemoDate = "all demo date"
    AllDemoReworkRequiredStartDate = "all demo rework required start date"
    AllDemoBlockedStartDate = "all demo blocked start date"
    AllDemoBlockedEndDate = "all demo blocked end date"
    AllDemoReworkRequiredEndDate = "all demo rework required end date"
    AllReadyForClientVerificationDate = "all ready for client verification date"
    VerificationAssignedTo = "verification assigned to"
    AllClientVerificationStartDate = "all client verification start date"
    AllClientVerificationEndDate = "all client verification end date"
    AllClientReworkRequiredStartDate = "all client rework required start date"
    AllClientBlockedStartDate = "all client blocked start date"
    AllClientBlockedEndDate = "all client blocked end date"
    AllClientReworkRequiredEndDate = "all client rework required end date"
    VerificationCompleteDate = "verification complete date"
    ReadyToMigrateDate = "ready to migrate date"

desc_keys_values = [member.value for member in desc_keys]
expected_length = len(desc_keys_values)

class columns(Enum):
    Site = "Site"
    ConfigurationInProgress = "03.Configuration In Progress"
    Blocked = "04.Blocked"
    EffectiveConfigurationDays = "Effective Configuration Days"
    NumberOfPeerReviews = "Number of Peer Reviews"
    PeerReviewInProgress = "06.Peer Review In Progress"
    PeerReviewReworkReq = "07.Peer Review - Rework Req."
    BlockedPeerReviewReworkDays = "Blocked Peer Review Rework Days"
    EffectivePeerReviewReworkDays = "Effective Peer Review Rework Days"
    NumberOfReadyForDemo = "Number of Ready for Demo"
    NumberOfDemos = "Number of Demos"
    DemoReworkDays = "Demo Rework Days"
    BlockedDemoReworkDays = "Blocked Demo Rework Days"
    EffectiveDemoReworkDays = "Effective Demo Rework Days"
    NumberOfReadyForClientVerification = "Number of Ready for Client Verification"
    NumberOfClientVerification = "Number of Client Verification"
    VerificationInProgress = "13.Verification In Progress"
    ClientReworkInProgress = "11.Client Rework In Progress"
    BlockedReworkClientVerificationDays = "Blocked Rework Client Verification Days"
    EffectiveReworkClientVerificationDays = "Effective Rework Client Verification Days"
    VerificationCompleteDate = "Verification Complete Date"
    ReadyToMigrateDate = "Ready To Migrate Date"
    ConfigurationComplete = "05.Configuration Complete"
    # BlockedBucketTime = "Blocked Bucket Time"
    # PeerReviewReworkBucketTime = "Peer Review Rework Bucket Time"
    ReadyForDemo = "08.Ready For Demo"
    DemoInProgress = "09.Demo In Progress"
    ReadyForClientVerification = "12.Ready For Client Verification"
    VerificationComplete = "14.Verification Complete"

class CaseInsensitiveDict(dict):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._convert_keys()

    def _convert_keys(self):
        for key in list(self.keys()):
            self._convert_key(key)

    def _convert_key(self, key):
        if isinstance(key, str):
            lower_key = key.lower()
            if lower_key != key:
                self[lower_key] = super().pop(key)
    
    def __setitem__(self, key, value):
        if isinstance(key, str):
            key = key.lower()
        super().__setitem__(key, value)

    def __getitem__(self, key):
        if isinstance(key, str):
            key = key.lower()
        return super().__getitem__(key)

    def __delitem__(self, key):
        if isinstance(key, str):
            key = key.lower()
        super().__delitem__(key)

    def __contains__(self, key):
        if isinstance(key, str):
            key = key.lower()
        return super().__contains__(key)

    def get(self, key, default=None):
        if isinstance(key, str):
            key = key.lower()
        return super().get(key, default)

    def setdefault(self, key, default=None):
        if isinstance(key, str):
            key = key.lower()
        return super().setdefault(key, default)

    def pop(self, key, default=None):
        if isinstance(key, str):
            key = key.lower()
        return super().pop(key, default)

    def update(self, *args, **kwargs):
        super().update(*args, **kwargs)
        self._convert_keys()

class Category(Enum):
    Network = "Network"
    Length = "Lenght"
    Value = "Value"

class AgeingParam:
    def __init__(
        self,
        key: str,
        category: str,
        start_dates: Callable[[any], List[str]],
        end_dates: Callable[[any], List[str]] = None,
        date_pos: bool = None,
    ):
        self.key = key
        self.start_dates = start_dates
        self.category = category  # Initialize required attribute
        self.end_dates = end_dates
        self.date_pos = date_pos

ageing_params = [
    # Configuration
    # Network
    AgeingParam(key=columns.ConfigurationInProgress.value, start_dates=desc_keys.ConfigStartDate.value, end_dates=desc_keys.ConfigEndDate.value, category=Category.Network.value),
    # Network
    AgeingParam(key=columns.Blocked.value, start_dates=desc_keys.AllConfigBlockedStartDate.value, end_dates=desc_keys.AllConfigBlockedEndDate.value, category=Category.Network.value),

    # Peer Review
    # Length
    AgeingParam(key=columns.NumberOfPeerReviews.value, start_dates=desc_keys.AllPeerReviewStartDate.value, category=Category.Length.value),
    # Network
    AgeingParam(key=columns.PeerReviewInProgress.value, start_dates=desc_keys.AllPeerReviewStartDate.value, end_dates=desc_keys.AllPeerReviewEndDate.value, category=Category.Network.value),
    # Network
    AgeingParam(key=columns.PeerReviewReworkReq.value, start_dates=desc_keys.AllPeerReviewReworkRequiredStartDate.value, end_dates=desc_keys.AllPeerReviewReworkRequiredEndDate.value, category=Category.Network.value),
    # Network
    AgeingParam(key=columns.BlockedPeerReviewReworkDays.value, start_dates=desc_keys.AllPRReworkBlockedStartDate.value, end_dates=desc_keys.AllPRReworkBlockedEndDate.value, category=Category.Network.value),

    # Demo
    # Length
    AgeingParam(key=columns.NumberOfReadyForDemo.value, start_dates=desc_keys.AllReadyForDemoDate.value, category=Category.Length.value),
    # Length
    AgeingParam(key=columns.NumberOfDemos.value, start_dates=desc_keys.AllDemoDate.value, category=Category.Length.value),
    # Network
    AgeingParam(key=columns.DemoReworkDays.value, start_dates=desc_keys.AllDemoReworkRequiredStartDate.value, end_dates=desc_keys.AllDemoReworkRequiredEndDate.value, category=Category.Network.value),
    # Network
    AgeingParam(key=columns.BlockedDemoReworkDays.value, start_dates=desc_keys.AllDemoBlockedStartDate.value, end_dates=desc_keys.AllDemoBlockedEndDate.value, category=Category.Network.value),

    # Client Verification
    # Length
    AgeingParam(key=columns.NumberOfReadyForClientVerification.value, start_dates=desc_keys.AllReadyForClientVerificationDate.value, category=Category.Length.value),
    # Length
    AgeingParam(key=columns.NumberOfClientVerification.value, start_dates=desc_keys.AllClientVerificationStartDate.value, category=Category.Length.value),
    # Network
    AgeingParam(key=columns.VerificationInProgress.value, start_dates=desc_keys.AllClientVerificationStartDate.value, end_dates=desc_keys.AllClientVerificationEndDate.value, category=Category.Network.value),
    # Network
    AgeingParam(key=columns.ClientReworkInProgress.value, start_dates=desc_keys.AllClientReworkRequiredStartDate.value, end_dates=desc_keys.AllClientReworkRequiredEndDate.value, category=Category.Network.value),
    # Network
    AgeingParam(key=columns.BlockedReworkClientVerificationDays.value, start_dates=desc_keys.AllClientBlockedStartDate.value, end_dates=desc_keys.AllClientBlockedEndDate.value, category=Category.Network.value),

    # Insert the first date as a value
    # Value
    AgeingParam(key=columns.VerificationCompleteDate.value, start_dates=desc_keys.VerificationCompleteDate.value, category=Category.Value.value),
    # Value
    AgeingParam(key=columns.ReadyToMigrateDate.value, start_dates=desc_keys.ReadyToMigrateDate.value, category=Category.Value.value),

    # Bucket Time
    # Network
    AgeingParam(key=columns.ConfigurationComplete.value, start_dates=desc_keys.AllPeerReviewStartDate.value, end_dates=desc_keys.ConfigEndDate.value, date_pos=0, category=Category.Network.value),
    # Network
    # AgeingParam(key=columns.BlockedBucketTime.value, start_dates=desc_keys.AllConfigBlockedStartDate.value, end_dates=desc_keys.AllConfigBlockedEndDate.value, date_pos=-1, category=Category.Network.value),
    # Network
    # AgeingParam(key=columns.PeerReviewReworkBucketTime.value, start_dates=desc_keys.AllPeerReviewReworkRequiredStartDate.value, end_dates=desc_keys.AllPeerReviewEndDate.value, date_pos=0, category=Category.Network.value),
    # Network
    AgeingParam(key=columns.ReadyForDemo.value, start_dates=desc_keys.AllDemoDate.value, end_dates=desc_keys.AllReadyForDemoDate.value, date_pos=0, category=Category.Network.value),
    # Network
    AgeingParam(key=columns.DemoInProgress.value, start_dates=desc_keys.AllReadyForClientVerificationDate.value, end_dates=desc_keys.AllDemoDate.value, date_pos=0, category=Category.Network.value),
    # Network
    AgeingParam(key=columns.ReadyForClientVerification.value, start_dates=desc_keys.AllClientVerificationStartDate.value, end_dates=desc_keys.AllReadyForClientVerificationDate.value, date_pos=0, category=Category.Network.value),
    # Network
    AgeingParam(key=columns.VerificationComplete.value, start_dates=desc_keys.ReadyToMigrateDate.value, end_dates=desc_keys.VerificationCompleteDate.value, date_pos=0, category=Category.Network.value)
]

def is_leap_year(year: int) -> bool:
    """Determine if a given year is a leap year."""
    return (year % 4 == 0 and year % 100 != 0) or (year % 400 == 0)

def is_valid_date(date_string: str) -> bool:
    """Check if the provided date string is a valid date."""
    valid_date = True
    try:
        date_numbers_strings = date_string.split('/')
        if len(date_numbers_strings) != 3:
            return False

        month, day, year = date_numbers_strings[0],date_numbers_strings[1],date_numbers_strings[2]
        if (len(year) != 2 and len(year) != 4) or len(day) > 2 or len(month) > 2:
            return False

        # Parse the month, day, and year
        month, day, year = map(int, date_numbers_strings)
        
        # Days in each month
        days_in_month = [31, 29 if is_leap_year(year) else 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]

        # Validate the ranges of month, day, and year
        if not (1 <= month <= 12):
            valid_date = False
        elif not (1 <= day <= days_in_month[month - 1]):
            valid_date = False

        # Handle two-digit year conversion
        current_year = datetime.now().year
        century = (current_year // 100) * 100
        if len(str(year)) == 2:
            year += century

        # Validate the year range
        if not (1900 <= year <= century + 999):
            valid_date = False

    except (ValueError, IndexError):
        valid_date = False

    return valid_date

def parse_descriptions(template: pd.DataFrame, logger: Logger):
    """Parse 'Description' column to extract key-value pairs and dates, mapping 'Task ID' to the processed description."""
    description_column = template['Description']
    task_ids = template['Task ID']
    sites = template['Site']
    task_names = template['Task Name']
    
    result_map = {}  # Initialize the result map to store processed data
    
    for task_id, text, site, task_name in zip(task_ids, description_column, sites, task_names):
        description_text = text.lower()
        more_than_one_space_regex = re.compile(r'\s+')
        if pd.isna(description_text):
            description_text = ''  # Replace NaN or None with an empty string if necessary
        
        if len(description_text) == 0:
            logger.INFO("", task_id, task_name, site, "Empty Description")
            continue  # Skip further processing for empty descriptions

        result_map[task_id] = {}
        
        # Split description text into lines
        lines = description_text.split('\n')
        
        for line in lines:
            if ':' in line and len(line.split(':')) == 2:
                pair = line.split(':', 1)
                key = more_than_one_space_regex.sub(' ', pair[0]).lower().strip()  # Key is before the colon
                value = more_than_one_space_regex.sub(' ', pair[1]).strip()  # Value is after the colon
                
                # Check if the key is a valid key
                if key not in desc_keys_values:
                    continue  # Skip keys with invalid characters

                # Extract dates from the value
                date_regex = re.compile(r'\b(\d{1,2}/\d{1,2}/\d{2,4})\b')  # Regex to match MM/DD/YY or MM/DD/YYYY
                dates = []
                
                # If the value contains dates
                if any(date_regex.search(val) for val in value.split(';')):
                    dates = [date_match.group(1) for date_match in date_regex.finditer(value)]
                
                # If value does not contain dates or is not valid
                inValidDates = []
                if not dates:
                    dates = []  # Set to empty list if no valid dates are found or if value is invalid
                else:
                    inValidDates = [date for date in dates if not is_valid_date(date)]

                if len(inValidDates) != 0:
                    logger.ERROR(line, task_id, task_name, site, "Invalid Date")
                
                result_map[task_id][key] = dates  # Store dates in the result map under the key

        # Check for missing keys
        collected_keys = list(result_map[task_id].keys())
        intersection = [key_ for key_ in desc_keys_values if key_ in collected_keys]
        if(len(intersection) == 0):
            logger.INFO("", task_id, task_name, site, "This Task does not contain info for ageing - If this is the idea, just ignore, if not check the format")
        elif len(intersection) != expected_length:
            differences = [key_ for key_ in desc_keys_values if key_ not in collected_keys]
            if len(differences) != 0:
                for difference in differences:
                    issueLines = [line for line in lines if difference in line]
                    if len(issueLines) == 1:
                        logger.ERROR(issueLines[0], task_id, task_name, site, f"The key '{difference}' was found on this line, but it is not properly formatted")
                    elif len(issueLines) == 0:
                        logger.ERROR("", task_id, task_name, site, f"The key '{difference}' was not found")
    return result_map

def assign_sites(template,sites_dict):
    # Asign the site to each row
    template['Site'] = template['Site'].astype(str)
    for file_name, task_IDs in sites_dict.items():
        site = file_name.split(' ')[0]
        template.loc[template['Task ID'].isin(task_IDs), 'Site'] = site

    return template

def one_date(dates, pos):
    """Returns the first element of the array if it is not empty, otherwise returns None."""
    return [dates[pos]] if dates else []

def get_us_holidays_current_year_for_networkdays(cal = UnitedStates()) -> str:
    """Get the US holidays for the current year formatted for NETWORKDAYS.INTL in Excel."""
    # Get the current year
    year = datetime.now().year

    # Get holidays for the current year
    holidays = cal.holidays(year)

    # Format holidays as mm/dd/yyyy
    formatted_holidays = [date.strftime('%m/%d/%Y') for date, name in holidays]
    formatted_holidays += other_PTO_dates
    # Join the holidays with commas and add curly braces
    return '{' + ','.join(f'"{date}"' for date in formatted_holidays) + '}'


holidays_range = get_us_holidays_current_year_for_networkdays()
def insert_network_days_formula(task_id, param: AgeingParam, candidate, template):
    formula = ''
    now = datetime.now()  # Current date
    weekends = 1
    start_dates = candidate[param.start_dates]
    end_dates:list = candidate[param.end_dates]
    

    while len(end_dates) < len(start_dates):
        current_date = now.strftime('%m/%d/%y')
        end_dates.append(current_date)

    # If using only the first date in the ranges
    date_pos = param.date_pos
    if date_pos is not None:
        start_dates = one_date(candidate[param.start_dates], date_pos)
        end_dates = one_date(candidate[param.end_dates], date_pos)
        
            

    for start_date, end_date in zip(start_dates, end_dates):
        # Directly use start_date and end_date if they are already in mm/dd/yy format
        formula += f'NETWORKDAYS.INTL("{start_date}", "{end_date}", {weekends}, {holidays_range})+'

    formula = '=0' if len(formula) == 0 else f"=ABS({formula[:-1]})"  # Remove the trailing '+'

    # Update the template DataFrame
    template[param.key] = template[param.key].astype(str)
    template.loc[template['Task ID'] == task_id, param.key] = formula

def insert_start_dates_length(task_id:str, param:AgeingParam, candidate, template):
    startDates = candidate[param.start_dates]
    lenght = len(startDates)
    template.loc[template['Task ID'] == task_id, param.key] = lenght

def insert_first_date_value(task_id:str, param:AgeingParam, candidate, template):
    value = candidate[param.start_dates]
    template[param.key] = template[param.key].astype(str)
    template.loc[template['Task ID'] == task_id, param.key] = value[0] if len(value) != 0 else ""

def insert_ageing_values(ageingCandidates, template:pd.DataFrame):
    for task_id, candidate in ageingCandidates.items():
        for param in ageing_params:
            category = param.category
            if category == Category.Length.value:
                insert_start_dates_length(task_id, param, candidate, template)
            elif category == Category.Value.value:
                insert_first_date_value(task_id, param, candidate, template)
            elif category == Category.Network.value:
                insert_network_days_formula(task_id, param, candidate, template)

def insert_label_values(template:pd.DataFrame):

    # Continue Here it must be like what I did with parse_descriptions function
    template_columns = template.columns.tolist()
    template['Labels'] = template['Labels'].astype(str)
    labels_column = template['Labels']
    task_ids = template['Task ID']

    for task_id, labels_text in zip(task_ids, labels_column):
        labels = [label.upper().strip() for label in labels_text.split(';')]
        for label in labels:
            if label in template_columns:
                template.loc[template['Task ID'] == task_id, label] = True

def clean_not_candidates(df, task_ids):
    # Create a set of column names from the columns Enum
    valid_columns = {column.value for column in columns}

    # Ensure the DataFrame contains only valid columns and Task ID
    if 'Task ID' not in df.columns:
        raise ValueError("DataFrame must contain 'Task ID' column")

    # Filter the DataFrame to include only valid columns and 'Task ID'
    valid_df_columns = [col for col in df.columns if col in valid_columns or col == 'Task ID']
    df = df[valid_df_columns].copy()  # Make an explicit copy to avoid SettingWithCopyWarning

    # Update the columns to empty strings where 'Task ID' is in the task_ids list
    for column in valid_columns:
        if column in df.columns:
            # Convert the column to string type before assigning empty strings
            df[column] = df[column].astype(str)
            df.loc[df['Task ID'].isin(task_ids), column] = ""

    return df

def excecute(file_paths = None):
    if file_paths is not None:
        store_excel_as_csv(file_paths)
    logger = Logger()
    data_frames, sites_dict = read_files(files_directory_path)
    template = read_excel_file(template_path)
    combined_dataframes = combine_dataframes(data_frames)
    template = move_files_info_to_template(combined_dataframes,template)
    template = assign_sites(template,sites_dict)
    descriptions = parse_descriptions(template, logger)
    tasks_to_be_ignored = logger.get_tasks_ids()
    clean_not_candidates(template, tasks_to_be_ignored)
    ageing_candidates = {key: value for key, value in descriptions.items() if key not in tasks_to_be_ignored}
    insert_ageing_values(ageing_candidates, template)
    insert_label_values(template)
    create_excel_with_content(template_path, reports_path, template, new_file_name, lattest_report_path, tasks_to_be_ignored)
    logger.save_to_file()

if __name__ == '__main__':
    excecute()