# **sharepoint-api**

A repo for Sharepoint Automation, having class methods for common operations like upload, download, copy &amp; delete files on sharepoint.

## **Usage**

### **Make a setting.cfg file**

Ensure that the sharepoint credentials (client secret, ids) are put into the 'settings.cfg' file, which further should be place at the root folder.
Below is the format of for setting.cfg file

```config=1
[default]
test_site_url = <url-insert-here>
prod_site_url = <url-insert-here>

[test_client_credentials]
client_id = <insert-here>
client_secret = <insert-here>

[prod_client_credentials]
client_id = <insert-here>
client_secret = <insert-here>
```

### **An example to use the SharepointAutomation class**

Below is a sample python code to move a specific sharepoint file.

```python=1
import configparser

from sharepoint_automation import SharepointAutomation

if __name__ == "__main__":

    # Sharepoint Automation constants
    SP_FOLDER_URL = r"/sites/Smart-Connected-Factory-Small-Ag-and-Turf-Sandbox/Shared Documents/copy_folder/"
    SP_FILE_URL = r"/sites/Smart-Connected-Factory-Small-Ag-and-Turf-Sandbox/Shared Documents/sample.py"

    config = configparser.ConfigParser()
    config.read("settings.cfg")

    SHAREPOINT_APP_SETTINGS = {
        "redirect_url": config["default"]["test_site_url"],
        "client_id": config["test_client_credentials"]["client_id"],
        "client_secret": config["test_client_credentials"]["client_secret"],
    }

    sp_obj = SharepointAutomation(SHAREPOINT_APP_SETTINGS)
    sp_obj.test_sharepoint_conection()
    sp_obj.move_file(source_sp_file_url=SP_FILE_URL, destination_sp_folder_url=SP_FOLDER_URL)
```

> Author: Abhishek Dev