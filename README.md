# AutomatedNuixUCRT
This version will automate the NuixUCRT so it can be run as a scheduled task
Automate UCRT
Step 1.  Download the latest automateUCRT scripts
	Files included in the script
1.	Settings.json – to be copied to C:\Program Files\Nuix\ScriptAutomate
2.	NuixCaseReports.db3 – to be copied to the location referenced by the “dblocation” node in the ettings.json file - i.e.  "dblocation": "C: \\Nuix\\UCRT\\Scripts"
3.	Database.rb_ – to be copied to the location referenced by the “dblocation” node in the Settings.json file - i.e.  "dblocation": "C: \\Nuix\\UCRT\\Scripts"
4.	SQLite.rb_ – to be copied to the location referenced by the “dblocation” node in the Settings.json file - i.e.  "dblocation": "C: \\Nuix\\UCRT\\Scripts"
Step 2. Modify Settings.json to meet your specifications – i.e.
{
  "useencryptedcreds":false,
  "useWindowsCredentials":false,
  "userName":"",
  "info":"",
  "licenseType":"cloud-server",
  "registryServer":"https://licence-api.nuix.com",
	  "dblocation": "C:\\Users\\ccarlson01\\Documents\\Nuix\\UCRT - CLS\\Backup",
  "caseslocations": "C:\\Nuix\\Sample Cases, C:\\Nuix_WORKING\\Cases",
  "dailycaseslocations": "C:\\Users\\ccarlson01\\Documents\\Nuix\\Sample Cases",
  "historiccaseslocations": "C:\\ProgramData\\Nuix\\NuixCases",
  "nuixversionmapping": [
    {
      "CaseVersion": "7.1.60003",
      "ConsoleVersions": [
        "7.2.1",
        "7.2.2",
        "7.2.3",
        "7.2.4"
      ],
      "ConsoleLocation": "C:\\Program Files\\Nuix\\Nuix 7.2"
    },
    {
      "CaseVersion": "7.5.68935",
      "ConsoleVersions": [
        "7.6.0",
        "7.2.3",
        "7.2.4",
        "7.2.6"
      ],
      "ConsoleLocation": "C:\\Program Files\\Nuix\\Nuix 7.6"
    },
    {
      "CaseVersion":"7.5.68939",
      "ConsoleVersions":["7.6.0",
        "7.2.3",
        "7.2.4",
        "7.2.6"],
      "ConsoleLocation":"C:\\Program Files\\Nuix\\Nuix 7.6"
    },
    {
      "CaseVersion":"8.0.0",
      "ConsoleVersions":["8.0.0"],
      "ConsoleLocation":"C:\\Program Files\\Nuix\\Nuix 8.0"
    },
    {
      "CaseVersion":"8.2.0",
      "ConsoleVersions":["8.2.0"],
      "ConsoleLocation":"C:\\Program Files\\Nuix\\Nuix 8.2"
    },
    {
      "CaseVersion":"8.5.1",
      "ConsoleVersions":["8.6.4"],
      "ConsoleLocation":"C:\\Program Files\\Nuix\\Nuix 8.6"
    },
    {
      "CaseVersion":"8.9.1",
      "ConsoleVersions":["9.0.0.171"],
      "ConsoleLocation":"C:\\Program Files\\Nuix\\Nuix 9.0"
    },
	{
      "CaseVersion":"9.0.0.171",
      "ConsoleVersions":["9.0.0.171"],
      "ConsoleLocation":"C:\\Program Files\\Nuix\\Nuix 9.0"
    },
	{
      "CaseVersion":"9.4.0",
      "ConsoleVersions":["9.4.0.105"],
      "ConsoleLocation":"C:\\Program Files\\Nuix\\Nuix 9.4"
    }
  ],
  "exporttype": "csv",
  "exportfields": "all",
  "exportdirectory": "C:\\Users\\ccarlson01\\Documents\\Nuix\\Automated UCRT",
  "exportfilename":"myexportfile.csv",
  "exportcaseinfo": "all",
  "cleanupfiles": "false",
  "cleanupfilerange": 30,
  "cleanupdirectories": "C:\\Users\\ccarlson01\\Documents\\Nuix\\UCRT - CLS\\Backup,C:\\Users\\ccarlson01\\Documents\\Nuix\\Automated UCRT",
  "cleanupfilestype": "*.log,*.csv",
  "ucrtreportinguser": "nuix-svc",
  "searchterm":"",
  "searchtermfile":"C:\\Users\\ccarlson01\\Documents\\Nuix\\Mizuho\\Bloomberg Search.txt",
  "nuixexportsearchresults":"true",
  "nuixexporttype":"Case Subset",
  "nuixexportworkers":2,
  "nuixexportworkermemory":1024,
  "exportonly":true,
  "reportype": "All",
  "report-run-frequency":"quarterly",
  "batch-run-frequency":"quarterly",
  "q1begindate":"4-1",
  "q1enddate":"6-30",
  "q2begindate":"7-1",
  "q2enddate":"9-30",
  "q3begindate":"10-1",
  "q3enddate":"12-31",
  "q4begindate":"1-1",
  "q4enddate":"3-30",
  "quarteroffset":"1",
  "batchloadhistorylookback":"1",
  "upgradecases": "false",
  "includedaterange": "false",
  "includeannotations": "true",
  "ignorecasehistory":"true",
  "cleanupdatabase": "false",
  "cleanupcaseguids": "all",
  "showsizein": "gb",
  "decimalpointaccuracy":4,
  "scripttorun":"C:\\ProgramData\\Nuix\\ProcessingFiles\\ScriptAutomate\\automateUCRT.rb",
  "nuixAppMemory":"Xmx4g",
  "nuixWorkers":"1",
  "nuixLicense":"eDiscovery Workstation",
  "consoleLocations":"C:\\Program Files\\Nuix\\Nuix 7.8\\nuix_console.exe,C:\\Program Files\\Nuix\\Nuix 8.2\\nuix_console.exe,C:\\Program Files\\Nuix\\Nuix 8.4\\nuix_console.exe,C:\\Program Files\\Nuix\\Nuix 8.6\\nuix_console.exe,C:\\Program Files\\Nuix\\Nuix 8.8\\nuix_console.exe,C:\\Program Files\\Nuix\\Nuix 9.0\\nuix_console.exe"
}
Step 3. Create a task scheduler (C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Administrative Tools\Task Scheduler)
 
In the actions section on the top right of the Task Scheduler click the “Create Task” link
 
 
Name: “AutomateNuixScript”
Description: “This will run the AutomateNuixScript.exe”
Security options
	When running the task, use the following user account: user account that will run the script
	Click the radio button Run whether user is logged on or not
Triggers:
 
Click the “New…” button
 
You can choose “Daily”, “Weekly”, “Monthly” based on your requirements.
Pick the frequency and time that you want to run the task
Click the Actions tab
 
Click the “New…” button
 
Action – “Start a program”
Settings:
	Program/script: "C:\Program Files\Nuix\AutomateNuixScript\AutomateNuixScript.exe" 
Click “OK”
Conditions:
 
Click “OK”
Settings:
 
Click “OK”
 
When you Click the “OK” button the Task Scheduler will ask for the username and password for the user that will be running the task.
Enter the password.
The task will now run on a scheduled basis.


