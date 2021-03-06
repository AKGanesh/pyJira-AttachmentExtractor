# pyJira-AttachmentExtractor
Jira sub-tasks attachment downloader using python and openpyxl

## Description
In a JIRA project you have different user stories, under each user story you will have sub-tasks.
Sub-tasks will have attachments along with regular id, summary, key, status and other fields.
If you want to download attachments under each subtask under user stories, use this .py file.
This program will create a folder structure (StoryId)-(Sub-taskId)/attachment.
This proram will also generate an excel file with all the attachment details, and embedded hyperlink to particular attachment.

###### Limitation:
Currently, this program will get only *one* attachment under **each** subtask under the user stories.
Can be extended to get all the attachments. 
JQL, JIRA API and python are powerful enough to meet your requirements.

## References:
Python JIRA 
https://jira.readthedocs.io/en/master/

Open Py XL
https://openpyxl.readthedocs.io/en/stable/

## Screenshots:
__UserStories Folders__

![User Stories Folders](/images/stories.png)

__Subtask attachments__

![Subtask attachment](/images/subtasks_attc.jpg)


__Excel file with the list of attachments downloaded__

![Excel file with the list of attachments downloaded](/images/excel.jpg)


### Modify the params in the .py file
jira_options={'server': 'https://<company>.atlassian.net'}

jira=JIRA(options=jira_options,basic_auth=('email@domain.com','accesstoken'))

### Directory to save extracted attachments and excel
dir_to_save = "Dir to save" 

_ex: D:\\jira-plugin\\CV_

wb.save("Dir to save excel") 

_ex :D:\\jira-plugin\\jira-report.xlsx_

### JQL
jql_query = "JQL query to suit the req" 

_ex: project = REC AND issuetype in (Epic, Story, Sub-task) AND "Epic Link" not in (REC-XX, REC-XX) AND Sprint in openSprints()_


## How to execute
If you have python environment installed
Python37> python pyJira.py

## To distribute as a .exe file
Bundle as .exe file using pyInstaller
https://pypi.org/project/PyInstaller/ 


