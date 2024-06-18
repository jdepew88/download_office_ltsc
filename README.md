**Instructions for Using Install_Office.ps1 Script**
Download the Script:

Download Install_Office.ps1 script from your GitHub repository or the location where it's hosted.
Prepare Your Environment:

Ensure you have a Windows system with PowerShell installed. The script is compatible with PowerShell on Windows.
Edit Script Variables (Optional):

Open Install_Office.ps1 in a text editor (like Notepad) if you need to modify any variables such as $setupUrl, $configUrl, or $outputDir to customize the script behavior.
Run the Script:

Right-click on Install_Office.ps1 and select "Run with PowerShell". Alternatively, open PowerShell and navigate to the directory containing the script using cd command, then run:
powershell
Copy code
.\Install_Office.ps1
This starts the script execution.
Follow Script Output:

The script will download setup.exe and configuration.xml from specified URLs into the $outputDir directory.

Setup.exe is directly from microsoft and is extracted when you run the Office Deployment Tool exe from https://www.microsoft.com/en-us/download/details.aspx?id=49117

You can like wise make your own configuration script using microsofts own tool (Office Customization Tool) at https://config.office.com/deploymentsettings.  
My script just simplfies the process by installing office per the configuration.xml as I have it set (has basic office suite, no Visio and no Projects).

It will display progress messages indicating when files are downloaded and when Office 2021 LTSC installation starts.
Monitor Installation:

During installation, a window may open to show the progress of Office installation. Follow on-screen prompts if any.

Completion:

Once the script completes, it will display a message indicating that Office installation has finished.
Notes:
Dependencies: The script relies on basic PowerShell functionality and may use wget or Invoke-WebRequest for file downloads. Ensure these utilities are available in your PowerShell environment.

Administrator Privileges: Depending on your system's security settings, you may need to run PowerShell with administrative privileges (Run as administrator) to ensure the script can download files and install Office without issues.

Customization: If you need to customize the Office version, configuration options, or installation directory, make a custon configuration.xml file with the Office Customization Tool at https://config.office.com/deploymentsettings.  

By following these instructions, you can effectively use the Install_Office.ps1 script to automate the download and installation of Office 2021 LTSC on your Windows system. Adjustments can be made based on specific organizational requirements or preferences.
