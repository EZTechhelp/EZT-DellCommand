# EZT-DellCommand

**Synopsis**

Automates checking for or installing Dell driver, firmware and other updates available using DellCommand
* * * 
**Features**

- Ability to scan for available updates and generate a HTML report
- Ability to save the HTML report to a specified directory
- Ability to email the HTML report as an attachment to a specified address
- Ability to install all available updates or specify to install specific types
- HTML report includes links to Dell support site and direct file downloads for each update
- HTML report includes ability to search, filter, or export to formats such as excel, CSV and PDF 
- Reads DellCommand.log in real-time to get accurate status and exit codes
- Outputs all action history, console or error messages to a log file and to iTarian RMM's execution log details

![Example HTML Report](/images/Example-Report.png)

**Notes**

- The script, with default configuration, will only **scan for updates**. The features for installing updates, sending emails as well as other **options can be configured via the iTarian procedure parameters or Script Configurable Variables**. See below for explanation of each parameter
- The script should be **run as the LocalSystem User** as it requires local admin privileges, since we are dealing with drivers and such. It will work under logged in user if said user is a local admin. 
- It is recommended to **leave the option to open the HTML report automatically disabled when running as a procedure**, as obviously you probably don't want it to be opened on a remote machine. The option is there if you want to run the script locally or manually on a machine 
- This is a **powershell script**. As such this will only work on **Windows** endpoints. It is untested on Powershell Core. Requires **Powershell v3** or higher. Tested on **Windows 10**. It may work on Windows 7 but is untested

## Available Versions

**[Self Executable DEMO](https://github.com/EZTechhelp/EZT-DellCommand/raw/main/EZT-DellCommand-DEMO.zip)**  

- Use this if you wish to test and try this script without using it as a procedure
- This version is only configured to scan for updates, generate a report and immediately open the report in your default browser. It otherwise cannot be configured
- Its packaged as a self-executable for quick and easy testing and mobility

**[Powershell Source Code](https://github.com/EZTechhelp/EZT-DellCommand/raw/main/EZT-DellCommand.ps1)** 

- Powershell only version, main script

**[Python Source Code](https://github.com/EZTechhelp/EZT-DellCommand/raw/main/EZT-DellCommand.py)**

- Python only version that is also the source code for the iTarian Procedure

**[iTarian Procedure](https://github.com/EZTechhelp/EZT-DellCommand/raw/main/EZT-DellCommand.json)**

- Basically just the Python version exported from iTarian. Use this to quickly import into your iTarian procedures

## Installation and Configuration

### Installation for iTarian Procedures

1. **Download the ITSM procedure** 
2. Within your ITSM portal, import the procedure under **Configuration Templates - Procedures**
3. Configure desired **procedure name, alert settings**..etc
4. Configure the **default parameters** for the procedure from the **Parameters tab** of the script. See **Configuration Parameters** below for explanations of each parameter
5. Click **Save - Ready to Review - Approve** to finish. **Assign to a profile** and optionally a schedule of your liking
6. **(Recommended)** Run the new procedure on a single **test machine** to ensure its working or configured as expected

#### iTarian Configuration

- This script can be configured by editing the **parameter options** within the iTarian RMM procedure 

#### Powershell Configuration

- If you wish to use the pure Powershell shell script version, use the configuration variables located in the region **Configurable Script Parameters** located near the top of the script 

### Configuration Parameters/Variables

These parameters should be fairly self-explanatory (and include some basic comments to explain them), however below is a detailed rundown

_**Note: 1 = Enabled, 0 = Disabled**_

#### Dell Command Configuration

-  **Install_Dell_Update_Types**
   - Default: none
   - Enter update types to install if detected, comma separated. 
   - Available Types you can use: **bios, firmware, driver, application, others, all**
   - Adding anything not listed above or leaving blank will disable installing updates
   - Adding '**all**' takes priority and will always install all updates 
   - Example to install just bios and driver updates: 
     - "bios,driver"
-  **InstallDellCommand_Update_Severity**
   - Default: all
   - Enter update severity types to install if detected, comma seperated
   - Available Types you can use: **security,critical,recommended,optional,all**
   - If blank, 'All' is assumed
   - Adding '**all**' takes priority and will always install severity updates, but ONLY if Install_Dell_Update_Types is defined and matches an available update. 
   - Example: to install just critical and security updates when Install_Dell_Update_Types is set: 
     - "security,critical"
-  **Dell_Install_Reboot**
   - Default: 0
   - Enables automatic reboot of system after updates are installed
-  **DownloadURL_Dell**
   - Default:  "https://dl.dell.com/FOLDER06986472M/2/Dell-Command-Update-Application-for-Windows-10_DF2DT_WIN_4.1.0_A00.EXE" 
   - Download URL for the Dell Command installer
   - Dell Command is installed silently if not already installed
   - Any updates to Dell Command itself will show as part of application updates in the report and can be installed along with them. You can also manually update this link to download the latest version if you prefer
-  **DownloadLocation_Dell**
   - Default: "C:\DellCommand"
   - Working directory where the Dell Command installer file will be downloaded and executed
   - The folder is not deleted automatically, but any log or XML files will be, leaving only the Dell installer file

#### Report Configuration

-  **Dashboard_Report**
   - Default: 1 
   - Enables generating an HTML Report containing all Dell Updates found
-  **Open_Dashboard_Report**
   - Default: 0
   - Enables immediately opening the report in the default browser after its created. 
-  **Save_Report_Location**
   - Default: "C:\Dell Reports"
   - Directory where the HTML dashboard report should be saved. 
   - If left blank, the report will be deleted after its generated or emailed
-  **Email_Report**
   - Default: 0
   - Enables sending the report via email. It will be attached as a HTML file
   - Requires configuring email variables if enabled
-  **Email_Logs**
   - Default: 0  
   - Enables attaching the script log file to the email sent with the report.
   - The log file is mostly for tracking run history, actions performed, debugging, auditing or troubleshooting
   - Log file only sends if Email_Report is also enabled and configured

#### Email Configuration

_**Important!**_

**When using as an iTarian Procedure** 

It is highly recommended to always **use iTarian parameters to configure passwords** for scripts vs adding it directly to them 

**When using as scripts** 

If you cannot use an RMM system like iTarian to pass credentials at runtime, there are alternative ways to handle this. For example, each time the script runs, it can grab the encrypted credentials from a file located in a secure share, decrypt, authenticate, then discard. This script is not configured currently to do this but I can provide a version that does if requested. Additionally, look into using secure vaults, such as the **Azure Key Vault**

**Email Variables** - Should be self-explanatory and if not, you probably shouldn't be using this

- **SMTP_Username** 
- **SMTP_Password**
- **SMTP_Port**
- **SMTP_Server**
- **SMTP_From**
- **SMTP_To**
- **SMTP_Subject**
  - Default: "DellCommand Update Report"
  - The computer name the script is run on is always added at the end of the title. 
  - Using the above default as an example, running it on a workstation named Workstation01, it would look like "DellCommand Update Report - Workstation01"

#### Log Configuration

-  **LogFile_Directory** 
   - Default: "C:\Logs\"
   - Directory where log file should be created
   - Feel free to change location if desired


