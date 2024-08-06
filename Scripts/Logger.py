from variables import log_file_path
from Utils import delete_all_files_in_folder

class Logger:
    INFO = 'INFO'
    WARNING = 'WARNING'
    ERROR = 'ERROR'

    # Class-level attribute to store messages
    messages_by_task_id = {}

    @classmethod
    def initialize_class_attributes(cls):
        """Initialize or reset class-level attributes."""
        cls.messages_by_task_id = {cls.INFO: {}, cls.WARNING: {}, cls.ERROR: {}}

    @classmethod
    def log(cls, level: str, line: str, task_id: str, task_name: str, site: str, issue_description: str):
        """Format and log a message into the internal dictionary, grouped by Task ID and log level."""
        if level not in cls.messages_by_task_id:
            raise ValueError(f"Unsupported log level: {level}")
        
        if task_id not in cls.messages_by_task_id[level]:
            cls.messages_by_task_id[level][task_id] = {
                'task_name': task_name,
                'line': line,
                'site': site,
                'issues': []
            }
        # Append the issue description to the list for the given Task ID
        cls.messages_by_task_id[level][task_id]['issues'].append({'line': line, 'description': issue_description})

    @classmethod
    def INFO(cls, line: str, task_id: str, task_name: str, site: str, issue_description: str):
        """Log a message with INFO level."""
        cls.log(cls.INFO, line, task_id, task_name, site, issue_description)

    @classmethod
    def WARNING(cls, line: str, task_id: str, task_name: str, site: str, issue_description: str):
        """Log a message with WARNING level."""
        cls.log(cls.WARNING, line, task_id, task_name, site, issue_description)

    @classmethod
    def ERROR(cls, line: str, task_id: str, task_name: str, site: str, issue_description: str):
        """Log a message with ERROR level."""
        cls.log(cls.ERROR, line, task_id, task_name, site, issue_description)

    @classmethod
    def get_messages(cls) -> dict:
        """Return the dictionary of messages grouped by log level and Task ID."""
        return cls.messages_by_task_id

    @classmethod
    def save_to_file(cls, filename: str = log_file_path):
        """Save all logged messages to a file, grouped by log level and Task ID, with overwriting."""
        content = ""
        # Define the order of log levels
        log_levels = [cls.ERROR, cls.INFO, cls.WARNING]

        # Iterate through each log level in the specified order
        for level in log_levels:
            if cls.messages_by_task_id[level]:
                content += f"Log Level: {level}\n"
                for task_id, details in cls.messages_by_task_id[level].items():
                    content += f"Task ID: {task_id}\n"
                    content += f"Task Name: {details['task_name']}\n"
                    content += f"Site: {details['site']}\n"
                    for issue in details['issues']:
                        if issue['line']:
                            content += f"Line: {issue['line']}\n"
                        content += f"Issue: {issue['description']}\n"
                    content += "\n\n"
                content += "\n\n"

        # Write the consolidated content to the file
        with open(filename, 'w') as file:
            file.write(content)

    def get_tasks_ids(cls) -> list:
        """Collect all unique task IDs from the logged messages."""
        task_ids = set()
        for level in cls.messages_by_task_id.values():
            for task_id in level.keys():
                task_ids.add(task_id)
        return list(task_ids)

# Initialize the Logger class attributes
Logger.initialize_class_attributes()

