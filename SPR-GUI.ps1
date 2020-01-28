#
# Load .NET Assemblies
#

[void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
[System.Windows.Forms.Application]::EnableVisualStyles();

#
# Functions
#

Function var-check()
{
write-host $fname
}

Set-Variable -name fname -Scope Global -Value 'C:\Temp\'

Function var-check2()
{
write-host $fname2
}

Function Button_ClickLogin($SPAddress)
{

$ModuleCheck = Get-Module -Name SharePointPnPPowerShellOnline -ListAvailable

if ($null -eq $ModuleCheck){

    $TextBox1.AppendText("`nRequired PowerShell module not installed, installing now..")

    Install-Module -Name SharePointPnPPowerShellOnline -Force

    #Check Again
    }
    $ModuleCheck2 = Get-Module -Name SharePointPnPPowerShellOnline -ListAvailable

    if ($null -eq $ModuleCheck2){

    $TextBox1.AppendText("`nIssue installing PowerShell module. Try running the application as Administrator.")

    }

    else {


            $TextBox1.AppendText("`nLogging into $($SPAddress), please login using Office 365 Login Form..")

            $comboBox1.Items.Clear()

            $login = Connect-PnPOnline -URL $SPAddress -UseWebLogin

            $Script:subsites = Get-PnPSubWebs

                ForEach($subsite in $subsites)
                {
 
                  $comboBox1.Items.add($subsite.Title)
    
                }
                $Form1.Controls.Add($comboBox1)

                if ($null -ne $subsites ){ 

                $TextBox1.AppendText("`r`nSuccessfully Logged In.")

                }
                else {

                $TextBox1.AppendText("`r`nIssue Logging In, Please Try Again.")


                }
        }
}


Function Button_ClickSaveLocation()
{
            function Get-FolderName {
            [CmdletBinding()]
            param (
                [Parameter(Mandatory=$false, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, Position=0)]
                [string]$Message = "Select a directory.",

                [string]$InitialDirectory = [System.Environment+SpecialFolder]::MyDocuments,

                [switch]$ShowNewFolderButton
            )

            $browserForFolderOptions = 0x00000041                                  # BIF_RETURNONLYFSDIRS -bor BIF_NEWDIALOGSTYLE
            if (!$ShowNewFolderButton) { $browserForFolderOptions += 0x00000200 }  # BIF_NONEWFOLDERBUTTON

            $browser = New-Object -ComObject Shell.Application
            # To make the dialog topmost, you need to supply the Window handle of the current process
            [intPtr]$handle = [System.Diagnostics.Process]::GetCurrentProcess().MainWindowHandle

            # see: https://msdn.microsoft.com/en-us/library/windows/desktop/bb773205(v=vs.85).aspx
            $folder = $browser.BrowseForFolder($handle, $Message, $browserForFolderOptions, $InitialDirectory)

            $result = $null
            if ($folder) {
                $result = $folder.Self.Path
            }

            # Release and remove the used Com object from memory
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($browser) | Out-Null
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()

            return $result
        }

        $folder = Get-FolderName

        $label3.Text = $folder
        $folderlocation = $folder
        $label3.Refresh()
        
        Set-Variable -name fname -Scope Global -Value $labellocation.Text

        $Button.Enabled = $true
 }


Function Button_Click($fname, $SPAddress)
{
    Add-Type -AssemblyName PresentationCore,PresentationFramework
    $ButtonType = [System.Windows.MessageBoxButton]::YesNo
    $MessageIcon = [System.Windows.MessageBoxImage]::Warning
    $MessageBody = "Warning: While the report is running, output will be displayed, but buttons will become unresponsive. `n`nYou can force close the application using the Task Manager if necessary.`r`n`nAre you sure you want to continue?"
    $MessageTitle = "$($Combobox1.SelectedItem) Report"

    $MessagePrompt = [System.Windows.MessageBox]::Show($MessageBody,$MessageTitle,$ButtonType,$MessageIcon)

    if ($MessagePrompt -eq 'Yes'){



        $i = 0

        $fname2 = $subsites | Where-Object {$_.Title -like $comboBox1.SelectedItem} | Select -ExpandProperty ServerRelativeUrl | Out-String

        Write-Host $fname2

                            #Function to Get Permissions Applied on a particular Object, such as: Web, List or Folder
                        Function Get-PnPPermissions([Microsoft.SharePoint.Client.SecurableObject]$Object)
                        {
                            #Determine the type of the object
                            Switch($Object.TypedObject.ToString())
                            {
                                "Microsoft.SharePoint.Client.Web"  { $ObjectType = "Site" ; $ObjectURL = $Object.URL; $ObjectTitle = $Object.Title }
                                "Microsoft.SharePoint.Client.ListItem"
                                {
                                    $ObjectType = "Folder"
                                    #Get the URL of the Folder 
                                    $Folder = Get-PnPProperty -ClientObject $Object -Property Folder
                                    $ObjectTitle = $Object.Folder.Name
                                    $ObjectURL = $("{0}{1}" -f $Web.Url.Replace($Web.ServerRelativeUrl,''),$Object.Folder.ServerRelativeUrl)
                                }
                                Default 
                                { 
                                    $ObjectType = $Object.BaseType #List, DocumentLibrary, etc
                                    $ObjectTitle = $Object.Title
                                    #Get the URL of the List or Library
                                    $RootFolder = Get-PnPProperty -ClientObject $Object -Property RootFolder     
                                    $ObjectURL = $("{0}{1}" -f $Web.Url.Replace($Web.ServerRelativeUrl,''), $RootFolder.ServerRelativeUrl)
                                }
                            }
    
                            #Get permissions assigned to the object
                            Get-PnPProperty -ClientObject $Object -Property HasUniqueRoleAssignments, RoleAssignments
  
                            #Check if Object has unique permissions
                            $HasUniquePermissions = $Object.HasUniqueRoleAssignments
      
                            #Loop through each permission assigned and extract details
                            $PermissionCollection = @()
                            Foreach($RoleAssignment in $Object.RoleAssignments)
                            { 
                                #Get the Permission Levels assigned and Member
                                Get-PnPProperty -ClientObject $RoleAssignment -Property RoleDefinitionBindings, Member
  
                                #Get the Principal Type: User, SP Group, AD Group
                                $PermissionType = $RoleAssignment.Member.PrincipalType
     
                                #Get the Permission Levels assigned
                                $PermissionLevels = $RoleAssignment.RoleDefinitionBindings | Select -ExpandProperty Name
  
                                #Remove Limited Access
                                $PermissionLevels = ($PermissionLevels | Where { $_ -ne "Limited Access"}) -join "; "
  
                                #Leave Principals with no Permissions assigned
                                If($PermissionLevels.Length -eq 0) {Continue}
  
                                #Check if the Principal is SharePoint group
                                If($PermissionType -eq "SharePointGroup")
                                {
                                    #Get Group Members
                                    $GroupMembers = Get-PnPGroupMembers -Identity $RoleAssignment.Member.LoginName
                  
                                    #Leave Empty Groups
                                    If($GroupMembers.count -eq 0){Continue}
                                    $GroupUsers = ($GroupMembers | Select -ExpandProperty Title | Where { $_ -ne "System Account"}) -join "; "
                                    If($GroupUsers.Length -eq 0) {Continue}
 
                                    #Add the Data to Object
                                    $Permissions = New-Object PSObject
                                    $Permissions | Add-Member NoteProperty Object($ObjectType)
                                    $Permissions | Add-Member NoteProperty Title($ObjectTitle)
                                    $Permissions | Add-Member NoteProperty URL($ObjectURL)
                                    $Permissions | Add-Member NoteProperty HasUniquePermissions($HasUniquePermissions)
                                    $Permissions | Add-Member NoteProperty Users($GroupUsers)
                                    $Permissions | Add-Member NoteProperty Type($PermissionType)
                                    $Permissions | Add-Member NoteProperty Permissions($PermissionLevels)
                                    $Permissions | Add-Member NoteProperty GrantedThrough("SharePoint Group: $($RoleAssignment.Member.LoginName)")
                                    $PermissionCollection += $Permissions
                                }
                                Else #User
                                {
                                    #Add the Data to Object
                                    $Permissions = New-Object PSObject
                                    $Permissions | Add-Member NoteProperty Object($ObjectType)
                                    $Permissions | Add-Member NoteProperty Title($ObjectTitle)
                                    $Permissions | Add-Member NoteProperty URL($ObjectURL)
                                    $Permissions | Add-Member NoteProperty HasUniquePermissions($HasUniquePermissions)
                                    $Permissions | Add-Member NoteProperty Users($RoleAssignment.Member.Title)
                                    $Permissions | Add-Member NoteProperty Type($PermissionType)
                                    $Permissions | Add-Member NoteProperty Permissions($PermissionLevels)
                                    $Permissions | Add-Member NoteProperty GrantedThrough("Direct Permissions")
                                    $PermissionCollection += $Permissions
                                }
                            }
                            #Export Permissions to CSV File
                            $PermissionCollection | Export-CSV $ReportFile -NoTypeInformation -Append
                        }
                        #Function to get sharepoint online site permissions report
                        Function Generate-PnPSitePermissionRpt()
                        {
                            [cmdletbinding()]     
                            Param  
                            (    
                                [Parameter(Mandatory=$false)] [String] $SiteURL, 
                                [Parameter(Mandatory=$false)] [String] $ReportFile,         
                                [Parameter(Mandatory=$false)] [switch] $Recursive,
                                [Parameter(Mandatory=$false)] [switch] $ScanFolders,
                                [Parameter(Mandatory=$false)] [switch] $IncludeInheritedPermissions
                            )  
                            Try {
                                #Connect to the Site
                                Connect-PnPOnline -URL $SiteURL -UseWebLogin
                                #Get the Web
                                $Web = Get-PnPWeb
  
                                $TextBox1.AppendText("`r`nGetting Site Collection Administrators...")
                                #Get Site Collection Administrators
                                $SiteAdmins = Get-PnPSiteCollectionAdmin
          
                                $SiteCollectionAdmins = ($SiteAdmins | Select -ExpandProperty Title) -join "; "
                                #Add the Data to Object
                                $Permissions = New-Object PSObject
                                $Permissions | Add-Member NoteProperty Object("Site Collection")
                                $Permissions | Add-Member NoteProperty Title($Web.Title)
                                $Permissions | Add-Member NoteProperty URL($Web.URL)
                                $Permissions | Add-Member NoteProperty HasUniquePermissions("TRUE")
                                $Permissions | Add-Member NoteProperty Users($SiteCollectionAdmins)
                                $Permissions | Add-Member NoteProperty Type("Site Collection Administrators")
                                $Permissions | Add-Member NoteProperty Permissions("Site Owner")
                                $Permissions | Add-Member NoteProperty GrantedThrough("Direct Permissions")
                
                                #Export Permissions to CSV File
                                Write-Host $ReportFile
                                $Permissions | Export-CSV $ReportFile -NoTypeInformation
                                $progressbar1.Value = 20
    
                                #Function to Get Permissions of Folders in a given List
                                Function Get-PnPFolderPermission([Microsoft.SharePoint.Client.List]$List)
                                {
                                    $TextBox1.AppendText("`r`nGetting Permissions of Folders in the List:$($List.Title)")
             
                                    #Get All Folders from List
                                    $ListItems = Get-PnPListItem -List $List -PageSize 2000
                                    $Folders = $ListItems | Where { ($_.FileSystemObjectType -eq "Folder") -and ($_.FieldValues.FileLeafRef -ne "Forms") -and (-Not($_.FieldValues.FileLeafRef.StartsWith("_")))}
 
                                    $ItemCounter = 0
                                    #Loop through each Folder
                                    ForEach($Folder in $Folders)
                                    {
                                        #Get Objects with Unique Permissions or Inherited Permissions based on 'IncludeInheritedPermissions' switch
                                        If($IncludeInheritedPermissions)
                                        {
                                            Get-PnPPermissions -Object $Folder
                                        }
                                        Else
                                        {
                                            #Check if Folder has unique permissions
                                            $HasUniquePermissions = Get-PnPProperty -ClientObject $Folder -Property HasUniqueRoleAssignments
                                            If($HasUniquePermissions -eq $True)
                                            {
                                                #Call the function to generate Permission report
                                                Get-PnPPermissions -Object $Folder
                                            }
                                        }
                                        $ItemCounter++
                                        $TextBox1.AppendText("`r`nGetting Permissions of Folders in List '$($List.Title)'") 
                                    }
                                    $progressbar1.Value = 40
                                }
  
                                #Function to Get Permissions of all lists from the given web
                                Function Get-PnPListPermission([Microsoft.SharePoint.Client.Web]$Web)
                                {
                                    #Get All Lists from the web
                                    $Lists = Get-PnPProperty -ClientObject $Web -Property Lists
    
                                    #Exclude system lists
                                    $ExcludedLists = @("Access Requests","App Packages","appdata","appfiles","Apps in Testing","Cache Profiles","Composed Looks","Content and Structure Reports","Content type publishing error log","Converted Forms",
                                    "Device Channels","Form Templates","fpdatasources","Get started with Apps for Office and SharePoint","List Template Gallery", "Long Running Operation Status","Maintenance Log Library", "Images", "site collection images"
                                    ,"Master Docs","Master Page Gallery","MicroFeed","NintexFormXml","Quick Deploy Items","Relationships List","Reusable Content","Reporting Metadata", "Reporting Templates", "Search Config List","Site Assets","Preservation Hold Library",
                                    "Site Pages", "Solution Gallery","Style Library","Suggested Content Browser Locations","Theme Gallery", "TaxonomyHiddenList","User Information List","Web Part Gallery","wfpub","wfsvc","Workflow History","Workflow Tasks", "Pages")
              
                                    $Counter = 0
                                    #Get all lists from the web   
                                    ForEach($List in $Lists)
                                    {
                                        #Exclude System Lists
                                        If($List.Hidden -eq $False -and $ExcludedLists -notcontains $List.Title)
                                        {
                                            $Counter++
                                            Write-Progress -PercentComplete ($Counter / ($Lists.Count) * 100) -Activity "Exporting Permissions from List '$($List.Title)' in $($Web.URL)" -Status "Processing Lists $Counter of $($Lists.Count)" -Id 1
  
                                            #Get Item Level Permissions if 'ScanFolders' switch present
                                            If($ScanFolders)
                                            {
                                                #Get Folder Permissions
                                                Get-PnPFolderPermission -List $List
                                            }
  
                                            #Get Lists with Unique Permissions or Inherited Permissions based on 'IncludeInheritedPermissions' switch
                                            If($IncludeInheritedPermissions)
                                            {
                                                Get-PnPPermissions -Object $List
                                            }
                                            Else
                                            {
                                                #Check if List has unique permissions
                                                $HasUniquePermissions = Get-PnPProperty -ClientObject $List -Property HasUniqueRoleAssignments
                                                If($HasUniquePermissions -eq $True)
                                                {
                                                    #Call the function to check permissions
                                                    Get-PnPPermissions -Object $List
                                                }
                                            }
                                        }
                                    }

                                    $progressbar1.Value = 60
                                }
    
                                #Function to Get Webs's Permissions from given URL
                                Function Get-PnPWebPermission([Microsoft.SharePoint.Client.Web]$Web) 
                                {
                                    #Call the function to Get permissions of the web
                                    $TextBox1.AppendText("`r`nGetting Permissions of the Web: $($Web.URL)...")
                                    Get-PnPPermissions -Object $Web
    
                                    #Get List Permissions
                                    $TextBox1.AppendText("`r`n`t Getting Permissions of Lists and Libraries...")
                                    Get-PnPListPermission($Web)
  
                                    #Recursively get permissions from all sub-webs based on the "Recursive" Switch
                                    If($Recursive)
                                    {
                                        #Get Subwebs of the Web
                                        $Subwebs = Get-PnPProperty -ClientObject $Web -Property Webs
  
                                        #Iterate through each subsite in the current web
                                        Foreach ($Subweb in $web.Webs)
                                        {
                                            #Get Webs with Unique Permissions or Inherited Permissions based on 'IncludeInheritedPermissions' switch
                                            If($IncludeInheritedPermissions)
                                            {
                                                Get-PnPWebPermission($Subweb)
                                            }
                                            Else
                                            {
                                                #Check if the Web has unique permissions
                                                $HasUniquePermissions = Get-PnPProperty -ClientObject $SubWeb -Property HasUniqueRoleAssignments
    
                                                #Get the Web's Permissions
                                                If($HasUniquePermissions -eq $true) 
                                                { 
                                                    #Call the function recursively                            
                                                    Get-PnPWebPermission($Subweb)
                                                }
                                            }
                                        }
                                    }

                                    $progressbar1.Value = 80
                                }
  
                                #Call the function with RootWeb to get site collection permissions
                                Get-PnPWebPermission $Web
    
                                $TextBox1.AppendText("`r`n*** Site Permission Report Generated Successfully!***")
                             }
                            Catch {
                                $TextBox1.AppendText("`r`nError Generating Site Permission Report! $($_.Exception.Message)")
                           }
                        }
    
                        #region ***Parameters***
                        $SiteURL= "$($SPAddress)" + "$($fname2)"
                        Write-Host "Site URL is "$SiteURL
                        $ReportFile = $fname + "\$($Combobox1.SelectedItem)-SharepointPermissionReport.csv"
                        #endregion
                        #Call the function to generate permission report
                        #Generate-PnPSitePermissionRpt -SiteURL $SiteURL -ReportFile $ReportFile -ScanFolders
                        Generate-PnPSitePermissionRpt -SiteURL $SiteURL -ReportFile $reportfile -Recursive -IncludeInheritedPermissions

                        $progressbar1.Value = 100

                    }
    
 }

Set-Variable -name fname2 -Scope Global -Value ''

$Form1 = New-Object System.Windows.Forms.Form
$Form1.ClientSize = New-Object System.Drawing.Size(685, 400)
$form1.topmost = $true
$Form1.Text = "Sharepoint Permissions Report"
$Form1.StartPosition = "CenterScreen"
$Form1.SizeGripStyle = "Hide"
$Form1.BackColor = "White"
$terminateScript = $false

$comboBox1 = New-Object System.Windows.Forms.ComboBox
$comboBox1.Location = New-Object System.Drawing.Point(15, 150)
$comboBox1.Size = New-Object System.Drawing.Size(150, 310)
$comboBox1.TabIndex = 3
$comboBox1.AutoCompleteMode =
    [System.Windows.Forms.AutoCompleteMode]::SuggestAppend;
$Form1.Controls.Add($comboBox1)

$ComboBox1_SelectedIndexChanged=
{
$fname2 = $comboBox1.SelectedItem
Write-Host $fname2   
}
$ComboBox1.add_SelectedIndexChanged($ComboBox1_SelectedIndexChanged)

$Button = New-Object System.Windows.Forms.Button
$Button.Enabled = $false
$Button.Location = New-Object System.Drawing.Point(15, 330)
$Button.Size = New-Object System.Drawing.Size(150, 50)
$Button.Text = "5. Run Report"
$Button.add_Click({Button_Click -fname $label3.Text -SPAddress $TextBox2.Text})
$Button.FlatStyle ="Standard"
$Button.TabIndex = 5
$Button.Font = "Arial"
$Form1.Controls.Add($Button)

$Button3 = New-Object System.Windows.Forms.Button
$Button3.Location = New-Object System.Drawing.Point(15, 190)
$Button3.Size = New-Object System.Drawing.Size(150, 50)
$Button3.Text = "4. Select Save Location"
$Button3.add_Click({Button_ClickSaveLocation})
$Button3.FlatStyle ="Standard"
$Button3.TabIndex = 4
$Button3.Font = "Arial"
$Form1.Controls.Add($Button3)

$Button2 = New-Object System.Windows.Forms.Button
$Button2.Location = New-Object System.Drawing.Point(15, 60)
$Button2.Size = New-Object System.Drawing.Size(150, 50)
$Button2.Text = "2. Login to Sharepoint"
$Button2.add_Click({Button_ClickLogin -SPAddress $TextBox2.Text})
$Button2.FlatStyle ="Standard"
$Button2.TabIndex = 2
$Button2.Font = "Arial"
$Form1.Controls.Add($Button2)

$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(15, 120)
$label.Size = New-Object System.Drawing.Size(150, 23)
$label.Text = "3. Select Subsite"
$label.Font = "Arial"
$Form1.Controls.Add($label)

$label2 = New-Object System.Windows.Forms.Label
$label2.Location = New-Object System.Drawing.Point(15, 260)
$label2.Size = New-Object System.Drawing.Size(150, 20)
$label2.Text = "Current Output Location:"
$label2.Font = "Arial"
$Form1.Controls.Add($label2)

$label3 = New-Object System.Windows.Forms.Label
$label3.Location = New-Object System.Drawing.Point(15, 280)
$label3.Size = New-Object System.Drawing.Size(150, 40)
$label3.Text = ""
$label3.Font = "Arial"
$label3.BorderStyle = 'Fixed3D'
$Form1.Controls.Add($label3)

$label4 = New-Object System.Windows.Forms.Label
$label4.Location = New-Object System.Drawing.Point(230, 35)
$label4.Size = New-Object System.Drawing.Size(250, 15)
$label4.Text = "Example: https://company.sharepoint.com"
$label4.Font = "Arial"
$Form1.Controls.Add($label4)

$label5 = New-Object System.Windows.Forms.Label
$label5.Location = New-Object System.Drawing.Point(15, 10)
$label5.Size = New-Object System.Drawing.Size(250, 15)
$label5.Text = "1. Enter Sharepoint Site Address:"
$label5.Font = "Arial"
$Form1.Controls.Add($label5)

$label6 = New-Object System.Windows.Forms.Label
$label6.Location = New-Object System.Drawing.Point(420, 75)
$label6.Size = New-Object System.Drawing.Size(55, 15)
$label6.Text = "Progress:"
$label6.Font = "Arial"
$Form1.Controls.Add($label6)

$TextBox1                        = New-Object system.Windows.Forms.TextBox
$TextBox1.multiline              = $true
$TextBox1.width                  = 482
$TextBox1.height                 = 290
$TextBox1.location               = New-Object System.Drawing.Point(197,103)
$TextBox1.Font                   = 'Microsoft Sans Serif,10'
$TextBox1.ScrollBars             = "Vertical" 
$TextBox1.AcceptsTab             = $false
$TextBox1.AutoSize               = $true
$TextBox1.ReadOnly               = $true
$Textbox1.BorderStyle            = 'Fixed3D'
$TextBox1.Enabled                = $false
$TextBox1.Font = "Arial"
$Form1.Controls.Add($TextBox1)

$TextBox2                        = New-Object system.Windows.Forms.TextBox
$TextBox2.Location = New-Object System.Drawing.Point(15, 30)
$TextBox2.Size = New-Object System.Drawing.Size(200, 23)
$Textbox2.TabIndex = 1
$TextBox2.Font = "Arial"
$textbox2.Add_KeyDown({
    if ($_.KeyCode -eq "Enter") {
        Button_ClickLogin -SPAddress $TextBox2.Text
    }
})
$Form1.Controls.Add($TextBox2)

$progressBar1 = New-Object System.Windows.Forms.ProgressBar
$progressBar1.Name = 'progressBar1'
$progressBar1.Value = 0
$progressBar1.Style="Continuous"
$progressBar1.Location = New-Object System.Drawing.Point(479, 70)
$progressBar1.Size = New-Object System.Drawing.Size(200, 23)

$form1.Controls.Add($progressBar1)

[void]$form1.showdialog()