# pyJira-AttachmentExtractor
Jira sub-tasks attachment downloader using python and openpyxl

## References:
Python JIRA 
https://jira.readthedocs.io/en/master/

Open Py XL
https://openpyxl.readthedocs.io/en/stable/

## Screenshots:
![User Stories Folders](/images/stories.png)
__UserStories Folders__

![Subtask attachment](/images/subtasks_attc.jpg)
__Subtask attachment__

__Excel file with the list of attachments downloaded__
![Excel file with the list of attachments downloaded](/images/excel.jpg)


### Modify the params in the .py file
jira_options={'server': 'https://<company>.atlassian.net'}

jira=JIRA(options=jira_options,basic_auth=('email@domain.com','accesstoken'))

### Directory to save extracted attachments and excel
dir_to_save = '<Dir to save>'  _ex: D:\\jira-plugin\\CV_
wb.save("<Dir to save excel>") _ex :D:\\jira-plugin\\jira-report.xlsx_

### JQL
jql_query = '<JQL to suit the req>' #project = REC AND issuetype in (Epic, Story, Sub-task) AND "Epic Link" not in (REC-XX, REC-XX) AND Sprint in openSprints()


## How to execute
If you have python environment installed
Python37> python pyJira.py

## To distribute as a .exe file
Bundle as .exe file using pyInstaller
https://pypi.org/project/PyInstaller/ 

###### Limitation:
Currently, this will get only *one* attachment under **each** subtask under the user stories.
Can be extended to get all the attachments.
