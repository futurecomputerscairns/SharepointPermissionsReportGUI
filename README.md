# Sharepoint Permissions Report GUI

This script has been designed to allow for running the [SharePoint Online: User Permissions User Permissions Audit Report for a Site Collection](https://www.sharepointdiary.com/2019/09/sharepoint-online-user-permissions-audit-report-using-pnp-powershell.html) in the form of a GUI. I take no credit for the report that the GUI generates, that was entirely made by sharepointdiary.com.  This is by no means a finished product, just a personal requirement that does the job, and is now being shared.

![SharePoint Permissions Report GUI](https://github.com/futurecomputerscairns/SharepointPermissionsReportGUI/blob/master/images/image.jpg)

## Prerequisites

The script will attempt to install the required module (SharePointPnPPowerShellOnline) if it doesn't already exist, when the user clicks on the login button.

## Running the report

1. Enter the site URL in the topmost textbox, i.e. https://contoso.sharepoint.com
2. Click the Login Button - This will launch an Office 365 login form (via Connect-PnPOnline -UseWebLogin) for you to enter your SharePoint Online credentials - IMPORTANT - You must use credentials of a user that has access to read permissions on all sites and subsites.
3. The 'Select Subsite' dropdown menu will be populated with all subsites in the SharePoint site. Select the site you wish to run the report on.
4. Click the '4. Select Save Location' button, and choose where you wish to save the report.
5. Once you are happy with the selections, click '5. Run Report'

Due to the report function currently not running in a runspace, the GUI becomes unresponsive while the report runs. The output window will still populate with the report stage, and the progress bar will update according to completion status.

## To Do:-

1. Run report in runspace to remove unresponsive GUI during report run
2. Add options to utilise all function parameters, such as running recursively
3. Add option to run report on all sites

## Feedback:-
Is always welcome! (constructively always helps :) )



 
