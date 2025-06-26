# Add assembly for Windows Forms
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Global variables for logging
$global:logFilePath = ""
$global:enableLogging = $true

# Function to write to both console and log file
function Write-LogHost {
    param(
        [string]$Message,
        [string]$ForegroundColor = "White",
        [switch]$NoNewLine
    )
    
    # Write to console
    if ($NoNewLine) {
        Write-Host $Message -ForegroundColor $ForegroundColor -NoNewline
    } else {
        Write-Host $Message -ForegroundColor $ForegroundColor
    }
    
    # Write to log file if logging is enabled and log path is set
    if ($global:enableLogging -and ![string]::IsNullOrEmpty($global:logFilePath)) {
        try {
            $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            $logMessage = "[$timestamp] $Message"
            Add-Content -Path $global:logFilePath -Value $logMessage -Encoding UTF8 -ErrorAction SilentlyContinue
        }
        catch {
            # Silently ignore log write errors to avoid breaking the main script
        }
    }
}

# Function to create the graphical interface
function Show-ConfigurationForm {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "LanLink -Audit Permission SharePoint - Configuração"
    $form.Size = New-Object System.Drawing.Size(500, 470)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    $form.BackColor = [System.Drawing.Color]::WhiteSmoke

    # Title
    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Text = "SharePoint Permissions Report"
    $titleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
    $titleLabel.ForeColor = [System.Drawing.Color]::DarkBlue
    $titleLabel.Location = New-Object System.Drawing.Point(20, 20)
    $titleLabel.Size = New-Object System.Drawing.Size(450, 30)
    $titleLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $form.Controls.Add($titleLabel)

    # Site Collection URL
    $urlLabel = New-Object System.Windows.Forms.Label
    $urlLabel.Text = "Site Collection URL:"
    $urlLabel.Location = New-Object System.Drawing.Point(20, 70)
    $urlLabel.Size = New-Object System.Drawing.Size(150, 20)
    $urlLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $form.Controls.Add($urlLabel)

    $urlTextBox = New-Object System.Windows.Forms.TextBox
    $urlTextBox.Location = New-Object System.Drawing.Point(180, 70)
    $urlTextBox.Size = New-Object System.Drawing.Size(280, 23)
    $urlTextBox.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $urlTextBox.PlaceholderText = "https://tenant.sharepoint.com/sites/sitename"
    $form.Controls.Add($urlTextBox)

    # Tenant ID
    $tenantLabel = New-Object System.Windows.Forms.Label
    $tenantLabel.Text = "Tenant ID:"
    $tenantLabel.Location = New-Object System.Drawing.Point(20, 110)
    $tenantLabel.Size = New-Object System.Drawing.Size(150, 20)
    $tenantLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $form.Controls.Add($tenantLabel)

    $tenantTextBox = New-Object System.Windows.Forms.TextBox
    $tenantTextBox.Location = New-Object System.Drawing.Point(180, 110)
    $tenantTextBox.Size = New-Object System.Drawing.Size(280, 23)
    $tenantTextBox.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $tenantTextBox.PlaceholderText = "00000000-0000-0000-0000-000000000000"
    $form.Controls.Add($tenantTextBox)

    # Client ID
    $clientLabel = New-Object System.Windows.Forms.Label
    $clientLabel.Text = "Client ID (Application ID):"
    $clientLabel.Location = New-Object System.Drawing.Point(20, 150)
    $clientLabel.Size = New-Object System.Drawing.Size(150, 20)
    $clientLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $form.Controls.Add($clientLabel)

    $clientTextBox = New-Object System.Windows.Forms.TextBox
    $clientTextBox.Location = New-Object System.Drawing.Point(180, 150)
    $clientTextBox.Size = New-Object System.Drawing.Size(280, 23)
    $clientTextBox.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $clientTextBox.PlaceholderText = "00000000-0000-0000-0000-000000000000"
    $form.Controls.Add($clientTextBox)

    # Certificate Thumbprint
    $thumbLabel = New-Object System.Windows.Forms.Label
    $thumbLabel.Text = "Certificate Thumbprint:"
    $thumbLabel.Location = New-Object System.Drawing.Point(20, 190)
    $thumbLabel.Size = New-Object System.Drawing.Size(150, 20)
    $thumbLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $form.Controls.Add($thumbLabel)

    $thumbTextBox = New-Object System.Windows.Forms.TextBox
    $thumbTextBox.Location = New-Object System.Drawing.Point(180, 190)
    $thumbTextBox.Size = New-Object System.Drawing.Size(280, 23)
    $thumbTextBox.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $thumbTextBox.PlaceholderText = "A1B2C3D4E5F6..."
    $form.Controls.Add($thumbTextBox)

    # Checkbox to include list items
    $includeItemsCheckBox = New-Object System.Windows.Forms.CheckBox
    $includeItemsCheckBox.Text = "Include list items with unique permissions"
    $includeItemsCheckBox.Location = New-Object System.Drawing.Point(20, 230)
    $includeItemsCheckBox.Size = New-Object System.Drawing.Size(300, 20)
    $includeItemsCheckBox.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $includeItemsCheckBox.Checked = $true
    $form.Controls.Add($includeItemsCheckBox)

    # Checkbox to exclude limited access
    $excludeLimitedCheckBox = New-Object System.Windows.Forms.CheckBox
    $excludeLimitedCheckBox.Text = "Exclude 'Limited Access' permissions"
    $excludeLimitedCheckBox.Location = New-Object System.Drawing.Point(20, 255)
    $excludeLimitedCheckBox.Size = New-Object System.Drawing.Size(300, 20)
    $excludeLimitedCheckBox.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $excludeLimitedCheckBox.Checked = $true
    $form.Controls.Add($excludeLimitedCheckBox)

    # Checkbox to generate HTML
    $generateHtmlCheckBox = New-Object System.Windows.Forms.CheckBox
    $generateHtmlCheckBox.Text = "Generate HTML report (in addition to CSV)"
    $generateHtmlCheckBox.Location = New-Object System.Drawing.Point(20, 280)
    $generateHtmlCheckBox.Size = New-Object System.Drawing.Size(300, 20)
    $generateHtmlCheckBox.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $generateHtmlCheckBox.Checked = $true
    $form.Controls.Add($generateHtmlCheckBox)

    # Informational text with hyperlinks
    $infoLabel1 = New-Object System.Windows.Forms.Label
    $infoLabel1.Text = "Feel free to contribute to this tool at:"
    $infoLabel1.Location = New-Object System.Drawing.Point(20, 315)
    $infoLabel1.Size = New-Object System.Drawing.Size(200, 15)
    $infoLabel1.Font = New-Object System.Drawing.Font("Segoe UI", 8)
    $infoLabel1.ForeColor = [System.Drawing.Color]::Gray
    $infoLabel1.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $form.Controls.Add($infoLabel1)

    $githubLinkLabel = New-Object System.Windows.Forms.LinkLabel
    $githubLinkLabel.Text = "GitHub Repository"
    $githubLinkLabel.Location = New-Object System.Drawing.Point(220, 315)
    $githubLinkLabel.Size = New-Object System.Drawing.Size(100, 15)
    $githubLinkLabel.Font = New-Object System.Drawing.Font("Segoe UI", 8)
    $githubLinkLabel.LinkColor = [System.Drawing.Color]::Blue
    $githubLinkLabel.Add_LinkClicked({
        Start-Process "https://github.com/ofabionobre/UniquePermissionSharePointAuditTool"
    })
    $form.Controls.Add($githubLinkLabel)

    $infoLabel2 = New-Object System.Windows.Forms.Label
    $infoLabel2.Text = "Initially developed by"
    $infoLabel2.Location = New-Object System.Drawing.Point(20, 335)
    $infoLabel2.Size = New-Object System.Drawing.Size(120, 15)
    $infoLabel2.Font = New-Object System.Drawing.Font("Segoe UI", 8)
    $infoLabel2.ForeColor = [System.Drawing.Color]::Gray
    $infoLabel2.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $form.Controls.Add($infoLabel2)

    $nobleLinkLabel = New-Object System.Windows.Forms.LinkLabel
    $nobleLinkLabel.Text = "nobre.cloud"
    $nobleLinkLabel.Location = New-Object System.Drawing.Point(140, 335)
    $nobleLinkLabel.Size = New-Object System.Drawing.Size(70, 15)
    $nobleLinkLabel.Font = New-Object System.Drawing.Font("Segoe UI", 8)
    $nobleLinkLabel.LinkColor = [System.Drawing.Color]::Blue
    $nobleLinkLabel.Add_LinkClicked({
        Start-Process "https://nobre.cloud"
    })
    $form.Controls.Add($nobleLinkLabel)

    # Buttons
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Text = "Execute Audit"
    $okButton.Location = New-Object System.Drawing.Point(180, 370)
    $okButton.Size = New-Object System.Drawing.Size(120, 30)
    $okButton.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $okButton.BackColor = [System.Drawing.Color]::ForestGreen
    $okButton.ForeColor = [System.Drawing.Color]::White
    $okButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $form.Controls.Add($okButton)

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = "Cancel"
    $cancelButton.Location = New-Object System.Drawing.Point(320, 370)
    $cancelButton.Size = New-Object System.Drawing.Size(100, 30)
    $cancelButton.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $cancelButton.BackColor = [System.Drawing.Color]::LightGray
    $cancelButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $form.Controls.Add($cancelButton)

    # Variables to store the results
    $script:SiteCollectionUrl = ""
    $script:TenantID = ""
    $script:ClientID = ""
    $script:Thumbprint = ""
    $script:includeListsItems = $true
    $script:excludeLimitedAccess = $true
    $script:generateHtml = $true
    $script:formResult = $false

    # Button events
    $okButton.Add_Click({
        if ([string]::IsNullOrWhiteSpace($urlTextBox.Text)) {
            [System.Windows.Forms.MessageBox]::Show("Please enter the Site Collection URL.", "Required Field", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            $urlTextBox.Focus()
            return
        }
        if ([string]::IsNullOrWhiteSpace($tenantTextBox.Text)) {
            [System.Windows.Forms.MessageBox]::Show("Please enter the Tenant ID.", "Required Field", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            $tenantTextBox.Focus()
            return
        }
        if ([string]::IsNullOrWhiteSpace($clientTextBox.Text)) {
            [System.Windows.Forms.MessageBox]::Show("Please enter the Client ID.", "Required Field", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            $clientTextBox.Focus()
            return
        }
        if ([string]::IsNullOrWhiteSpace($thumbTextBox.Text)) {
            [System.Windows.Forms.MessageBox]::Show("Please enter the Certificate Thumbprint.", "Required Field", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            $thumbTextBox.Focus()
            return
        }

        $script:SiteCollectionUrl = $urlTextBox.Text.Trim()
        $script:TenantID = $tenantTextBox.Text.Trim()
        $script:ClientID = $clientTextBox.Text.Trim()
        $script:Thumbprint = $thumbTextBox.Text.Trim()
        $script:includeListsItems = $includeItemsCheckBox.Checked
        $script:excludeLimitedAccess = $excludeLimitedCheckBox.Checked
        $script:generateHtml = $generateHtmlCheckBox.Checked
        $script:formResult = $true
        $form.Close()
    })

    $cancelButton.Add_Click({
        $script:formResult = $false
        $form.Close()
    })

    # Allow Enter to confirm
    $form.AcceptButton = $okButton
    $form.CancelButton = $cancelButton

    # Show the form
    $form.ShowDialog()

    return $script:formResult
}

# Initialize logging and result directories
$dateTime = (Get-Date).toString("dd-MM-yyyy-hh-ss")
$invocation = (Get-Variable MyInvocation).Value
$scriptRootPath = Split-Path $invocation.MyCommand.Path
$logDirectoryPath = Join-Path -Path $scriptRootPath -ChildPath "Logs"
$auditResultsPath = Join-Path -Path $scriptRootPath -ChildPath "AuditResults"
$global:logFilePath = Join-Path -Path $logDirectoryPath -ChildPath "Execution-Log_$dateTime.log"

# Ensure log directory exists
if (-not (Test-Path $logDirectoryPath)) {
    New-Item -Path $logDirectoryPath -ItemType Directory -Force | Out-Null
}

# Ensure AuditResults directory exists
if (-not (Test-Path $auditResultsPath)) {
    New-Item -Path $auditResultsPath -ItemType Directory -Force | Out-Null
    Write-LogHost "Created AuditResults directory: $auditResultsPath" -ForegroundColor Green
}

# Show graphical interface
$result = Show-ConfigurationForm

if (-not $result) {
    Write-LogHost "`n❌ Operation cancelled by user." -ForegroundColor Yellow
    Write-LogHost "Script will be terminated." -ForegroundColor Gray
    Write-LogHost "Press any key to exit..." -ForegroundColor Gray
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit 0
}

# Additional verification to ensure variables were defined
if ([string]::IsNullOrWhiteSpace($script:SiteCollectionUrl) -or 
    [string]::IsNullOrWhiteSpace($script:TenantID) -or 
    [string]::IsNullOrWhiteSpace($script:ClientID) -or 
    [string]::IsNullOrWhiteSpace($script:Thumbprint)) {
    Write-LogHost "`n❌ Error: Incomplete configuration information." -ForegroundColor Red
    Write-LogHost "Script will be terminated." -ForegroundColor Gray
    Write-LogHost "Press any key to exit..." -ForegroundColor Gray
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit 1
}

# Continue with the original script using collected variables
Clear-Host

$properties=@{SiteUrl='';SiteTitle='';ListTitle='';SensitivityLabel='';Type='';RelativeUrl='';ParentGroup='';MemberType='';MemberName='';MemberLoginName='';Roles='';}; 

# Variables $excludeLimitedAccess and $includeListsItems already defined by the graphical interface

# Use variables from the graphical interface
$SiteCollectionUrl = $script:SiteCollectionUrl
$TenantID = $script:TenantID
$ClientID = $script:ClientID
$Thumbprint = $script:Thumbprint
$excludeLimitedAccess = $script:excludeLimitedAccess
$includeListsItems = $script:includeListsItems
$generateHtml = $script:generateHtml

# Log the configurations (without showing sensitive data)
Write-LogHost "   Configurations loaded:" -ForegroundColor Cyan
Write-LogHost "   Site Collection URL: $($SiteCollectionUrl.Substring(0, [Math]::Min(50, $SiteCollectionUrl.Length)))..." -ForegroundColor Gray
Write-LogHost "   Tenant ID: $($TenantID.Substring(0, 8))..." -ForegroundColor Gray
Write-LogHost "   Client ID: $($ClientID.Substring(0, 8))..." -ForegroundColor Gray
Write-LogHost "   Thumbprint: $($Thumbprint.Substring(0, 8))..." -ForegroundColor Gray
Write-LogHost "   Include list items: $includeListsItems" -ForegroundColor Gray
Write-LogHost "   Exclude limited access: $excludeLimitedAccess" -ForegroundColor Gray
Write-LogHost "   Generate HTML report: $generateHtml" -ForegroundColor Gray
Write-LogHost ""

$global:siteTitle= "";
# Exclude certain libraries
$ExcludedLibraries = @("Form Templates", "Preservation Hold Library", "Site Assets", "Images", "Pages", "Settings", "Videos","Timesheet"
  "Site Collection Documents", "Site Collection Images", "Style Library", "AppPages", "Apps for SharePoint", "Apps for Office")

$global:permissions =@();
$global:sharingLinks = @();

function Get-ListItems_WithUniquePermissions{
  param(
      [Parameter(Mandatory)]
      [Microsoft.SharePoint.Client.List]$List,
      [Parameter(Mandatory)]
      [string]$SiteUrl
  )
  $selectFields = "ID,HasUniqueRoleAssignments,FileRef,FileLeafRef,FileSystemObjectType"
 
  $Url = $SiteUrl + '/_api/web/lists/getbytitle(''' + $($list.Title) + ''')/items?$select=' + $($selectFields)
  $nextLink = $Url
  $listItems = @()
  $Stoploop =$true
  while($nextLink){  
      do{
      try {
          $response = invoke-pnpsprestmethod -Url $nextLink -Method Get
          $Stoploop =$true
  
      }
      catch {
          Write-LogHost "An error occurred: $_  : Retrying" -ForegroundColor Red
          $Stoploop =$true
          Start-Sleep -Seconds 30
      }
  }
  While ($Stoploop -eq $false)
  
      $listItems += $response.value | where-object{$_.HasUniqueRoleAssignments -eq $true}
      if($response.'odata.nextlink'){
          $nextLink = $response.'odata.nextlink'
      }    else{
          $nextLink = $null
      }
  }

  return $listItems
}

Function PermissionObject($_object,$_type,$_relativeUrl,$_siteUrl,$_siteTitle,$_listTitle,$_memberType,$_parentGroup,$_memberName,$_memberLoginName,$_roleDefinitionBindings,$_sensitivityLabel)
{
  $permission = New-Object -TypeName PSObject -Property $properties; 
  $permission.SiteUrl =$_siteUrl; 
  $permission.SiteTitle = $_siteTitle; 
  $permission.ListTitle = $_listTitle; 
  $permission.SensitivityLabel = $_sensitivityLabel; 
  # Detect if it's a list item by URL
  if ($_Type -eq 0 -and $_relativeUrl -like "*/Lists/*") {
    $permission.Type = "List Item";
  } else {
    $permission.Type =  $_Type -eq 1 ? "Folder" : $_Type -eq 0 ? "File" : $_Type;
  }
  
  # Build the complete item URL
  if ([string]::IsNullOrEmpty($_relativeUrl)) {
    $permission.RelativeUrl = $_siteUrl;
      } else {
      # If it's already a complete URL, use as is
      if ($_relativeUrl.StartsWith("http")) {
        $permission.RelativeUrl = $_relativeUrl;
      } else {
        # For relative paths starting with /sites/ (like ServerRelativeUrl and FileRef)
        # they already contain the complete SharePoint path, so we just need to add the domain
        if ($_relativeUrl.StartsWith("/sites/") -or $_relativeUrl.StartsWith("/")) {
          # Extract only the domain from the site URL
          $uri = [System.Uri]$_siteUrl
          $domain = "$($uri.Scheme)://$($uri.Host)"
          $permission.RelativeUrl = "$domain$_relativeUrl";
        } else {
          # Build complete URL from site URL + relative path
          $baseUrl = $_siteUrl.TrimEnd('/')
          $relativePath = $_relativeUrl.TrimStart('/')
          $permission.RelativeUrl = "$baseUrl/$relativePath";
        }
    }
  }
  
  $permission.MemberType = $_memberType; 
  $permission.ParentGroup = $_parentGroup; 
  $permission.MemberName = $_memberName; 
  $permission.MemberLoginName = $_memberLoginName; 
  $permission.Roles = $_roleDefinitionBindings -join ","; 
  $global:permissions += $permission;
}

Function Extract-Guid ($inputString) {
  $splitString = $inputString -split '\|'
  return $splitString[2].TrimEnd('_o')
}

Function QueryUniquePermissionsByObject($_ctx,$_object,$_Type,$_RelativeUrl,$_siteUrl,$_siteTitle,$_listTitle)
{
  $roleAssignments = Get-PnPProperty -ClientObject $_object -Property RoleAssignments
   switch ($_Type) {
    0 { $sensitivityLabel = $_object.FieldValues["_DisplayName"] }
    1 { $sensitivityLabel = $_object.FieldValues["_DisplayName"] }
    "Site" { $sensitivityLabel = (Get-PnPSiteSensitivityLabel).displayname }
    "Sub-Site" { $sensitivityLabel = (Get-PnPSiteSensitivityLabel).displayname }
    default { " " }
}
  foreach($roleAssign in $roleAssignments){
    Get-PnPProperty -ClientObject $roleAssign -Property RoleDefinitionBindings,Member;
    $PermissionLevels = $roleAssign.RoleDefinitionBindings | Select -ExpandProperty Name;
    # Get all permission levels assigned (Excluding:Limited Access)  
    if($excludeLimitedAccess -eq $true){
       $PermissionLevels = ($PermissionLevels | Where { $_ -ne "Limited Access"}) -join ","  
    }
    $Users = Get-PnPProperty -ClientObject ($roleAssign.Member) -Property Users -ErrorAction SilentlyContinue
    # Get Access type
    $AccessType = $roleAssign.RoleDefinitionBindings.Name
    $MemberType = $roleAssign.Member.GetType().Name; 
    # Get the Principal Type: User, SP Group, AD Group  
    $PermissionType = $roleAssign.Member.PrincipalType  
  if( $_Type -eq 0){
      $sharingLinks = Get-PnPFileSharingLink -Identity $_object.FieldValues["FileRef"]
  }
  if( $_Type -eq 1){
      $sharingLinks = Get-PnPFolderSharingLink -Folder $_object.FieldValues["FileRef"]
  }

    If($PermissionLevels.Length -gt 0) {
      $MemberType = $roleAssign.Member.GetType().Name; 
       # Sharing link is in the format SharingLinks.03012675-2057-4d1d-91e0-8e3b176edd94.OrganizationView.20d346d3-d359-453b-900c-633c1551ccaa
        If ($roleAssign.Member.Title -like "SharingLinks*")
        {
          if($sharingLinks){
          $sharingLinks | where-object {$roleAssign.Member.Title -match $_.Id } | ForEach-Object{
            If ($Users.Count -gt 0) 
            {
                ForEach ($User in $Users)
                {
                PermissionObject $_object $_Type $_RelativeUrl $_siteUrl $_siteTitle $_listTitle "Sharing Links" $roleAssign.Member.LoginName  $user.Title $User.LoginName $_.Link.Type $sensitivityLabel; 
                }
            } 
            else {
              PermissionObject $_object $_Type $_RelativeUrl $_siteUrl $_siteTitle $_listTitle "Sharing Links" $roleAssign.Member.LoginName  $_.Link.Scope "" $_.Link.Type  $sensitivityLabel;
            }
          }  
        }
        }
      ElseIf($MemberType -eq "Group" -or $MemberType -eq "User")
      { 
        $MemberName = $roleAssign.Member.Title; 
        $MemberLoginName = $roleAssign.Member.LoginName;    
        if($MemberType -eq "User")
        {
          $ParentGroup = "NA";
        }
        else
        {
          $ParentGroup = $MemberName;
        }
        (PermissionObject $_object $_Type $_RelativeUrl $_siteUrl $_siteTitle $_listTitle $MemberType $ParentGroup $MemberName $MemberLoginName $PermissionLevels $sensitivityLabel); 
      }

      if(($_Type  -eq "Site" -or $_Type -eq "Sub-Site") -and $MemberType -eq "Group")
      {
        $sensitivityLabel = (Get-PnPSiteSensitivityLabel).DisplayName
                  If($PermissionType -eq "SharePointGroup")  {  
          # Get Group Members  
          $groupUsers = Get-PnPGroupMember -Identity $roleAssign.Member.LoginName                  
          $groupUsers|foreach-object{ 
            if ($_.LoginName.StartsWith("c:0o.c|federateddirectoryclaimprovider|") -and $_.LoginName.EndsWith("_0")) {
              $guid = Extract-Guid $_.LoginName
              
              Get-PnPMicrosoft365GroupOwners -Identity $guid | ForEach-Object {
                $user = $_
                (PermissionObject $_object $_Type $_RelativeUrl $_siteUrl $_siteTitle "" "GroupMember" $roleAssign.Member.LoginName $user.DisplayName $user.UserPrincipalName $PermissionLevels $sensitivityLabel); 
              }
            }
            elseif ($_.LoginName.StartsWith("c:0o.c|federateddirectoryclaimprovider|")) {
              $guid = Extract-Guid $_.LoginName
              
              Get-PnPMicrosoft365GroupMembers -Identity $guid | ForEach-Object {
                $user = $_
                (PermissionObject $_object $_Type $_RelativeUrl $_siteUrl $_siteTitle "" "GroupMember" $roleAssign.Member.LoginName $user.DisplayName $user.UserPrincipalName $PermissionLevels $sensitivityLabel); 
              }
            }

            (PermissionObject $_object $_Type $_RelativeUrl $_siteUrl $_siteTitle "" "GroupMember" $roleAssign.Member.LoginName $_.Title $_.LoginName $PermissionLevels $sensitivityLabel);   
          }
        }
      } 
    }      
  }
}
Function QueryUniquePermissions($_web, $_isSubSite = $false)
{
  ## query list, files and items unique permissions
  Write-LogHost "Querying web $($_web.Title)";
  $siteUrl = $_web.Url; 
 
  Write-LogHost $siteUrl -Foregroundcolor "Red"; 
  
  # For sub-sites, keep the root site title for the file
  if (-not $_isSubSite) {
    $global:siteTitle = $_web.Title; 
  }
  
  # Connect to specific site/sub-site context to search for lists
  try {
    if ($_isSubSite) {
      Write-LogHost "🔗 Connecting to sub-site: $siteUrl" -ForegroundColor Cyan
      Connect-PnPOnline -Url $siteUrl -ClientId $ClientID -Tenant $TenantID -Thumbprint $Thumbprint
    }
    
    $ll = Get-PnPList -Includes BaseType, Hidden, Title,HasUniqueRoleAssignments,RootFolder | Where-Object {$_.Hidden -eq $False -and $_.Title -notin $ExcludedLibraries } 
    Write-LogHost "Number of lists $($ll.Count)";

    # Determine site type based on whether it's a sub-site or not
    $siteType = if ($_isSubSite) { "Sub-Site" } else { "Site" }
    
    # Check if the site/sub-site has unique permissions
    $hasUniquePermissions = Get-PnPProperty -ClientObject $_web -Property HasUniqueRoleAssignments
    
    # For sub-sites, only include in report if it has unique permissions
    if (-not $_isSubSite -or $hasUniquePermissions) {
      Write-LogHost "   ✅ Site with unique permissions - including in report" -ForegroundColor Green
      QueryUniquePermissionsByObject $_web $_web $siteType "" $siteUrl $_web.Title "";
    } else {
      Write-LogHost "   ⏭️  Sub-site inheriting permissions - skipping site, but processing lists/items" -ForegroundColor Yellow
    }
   
    foreach($list in $ll)
    {      
      $listUrl = $list.RootFolder.ServerRelativeUrl; 
      # Exclude internal system lists and check if it has unique permissions 
      if($list.Hidden -ne $True)
      { 
        Write-LogHost $list.Title  -Foregroundcolor "Yellow"; 
        $listTitle = $list.Title; 
        # Check List Permissions 
        if($list.HasUniqueRoleAssignments -eq $True)
        { 
          $Type = $list.BaseType.ToString(); 
          QueryUniquePermissionsByObject $_web $list $Type $listUrl $siteUrl $_web.Title $listTitle;
        }
        
        if($includeListsItems){         
          $collListItem =  Get-ListItems_WithUniquePermissions -List $list -SiteUrl $siteUrl
          $count = $collListItem.Count
          Write-LogHost  "Number of items with unique permissions: $count within list $listTitle" 
          foreach($item in $collListItem) 
          {
              $Type = $item.FileSystemObjectType; 
              $fileUrl = $item.FileRef;  
              $i = Get-PnPListItem -List $list -Id $item.ID
              QueryUniquePermissionsByObject $_web $i $Type $fileUrl $siteUrl $_web.Title $listTitle;
          } 
        }
      }
    }
  }
  catch {
    Write-LogHost "❌ Error processing site $($_web.Title): $($_.Exception.Message)" -ForegroundColor Red
  }
  finally {
    # If a sub-site was processed, reconnect to the root site
    if ($_isSubSite) {
      Write-LogHost "🔙 Reconnecting to root site: $SiteCollectionUrl" -ForegroundColor Gray
      Connect-PnPOnline -Url $SiteCollectionUrl -ClientId $ClientID -Tenant $TenantID -Thumbprint $Thumbprint
    }
  }
}

if(Test-Path $logDirectoryPath){
 
  # ====== AUTHENTICATION WITH APP REGISTRATION + CERTIFICATE ======
  try {
    Connect-PnPOnline -Url $SiteCollectionUrl -ClientId $ClientID -Tenant $TenantID -Thumbprint $Thumbprint
    Write-LogHost "✅ Successfully connected using Application Registration" -ForegroundColor Green
  }
  catch {
    Write-LogHost "❌ Error connecting: $($_.Exception.Message)" -ForegroundColor Red
    Write-LogHost "Please verify:" -ForegroundColor Yellow
    Write-LogHost "  - The Tenant ID is correct" -ForegroundColor Yellow
    Write-LogHost "  - The Client ID is correct" -ForegroundColor Yellow
    Write-LogHost "  - The certificate is installed and the thumbprint is correct" -ForegroundColor Yellow
    Write-LogHost "  - The App Registration has the necessary permissions" -ForegroundColor Yellow
    exit 1
  }
  # =============================================================

  # array storing permissions
  Write-LogHost "🔍 Searching all sites and sub-sites recursively..." -ForegroundColor Cyan
  
  # Get all sites including root site and sub-sites recursively
  $allWebs = Get-PnPSubWeb -Recurse -IncludeRootWeb
  
  Write-LogHost "📊 Found: $($allWebs.Count) sites for audit" -ForegroundColor Green
  
  # Get the root site URL for comparison
  $rootWebUrl = $SiteCollectionUrl
  
  # Process each site (root and sub-sites)
  $currentSite = 1
  foreach ($web in $allWebs) {
    Write-LogHost "🔄 Processing site $currentSite of $($allWebs.Count): $($web.Title)" -ForegroundColor Yellow
    
    # Check if it's a sub-site by comparing with the root site URL
    $isSubSite = $web.Url -ne $rootWebUrl
    
    # Process the current site
    QueryUniquePermissions $web $isSubSite
    
    $currentSite++
  }

  Write-LogHost "Permission count: $($global:permissions.Count)";
  # Use the root site title for the file name (already defined in $global:siteTitle)
  $exportFilePath = Join-Path -Path $auditResultsPath -ChildPath $([string]::Concat($global:siteTitle,"-Permissions_",$dateTime,".csv"));
  
  Write-LogHost "Saving results to AuditResults directory:" $auditResultsPath -ForegroundColor Cyan
  Write-LogHost "Export File Path is:" $exportFilePath -ForegroundColor Cyan
  Write-LogHost "Number of lines exported is :" $global:permissions.Count -ForegroundColor Cyan
 
  $global:permissions | Select-Object SiteUrl,SiteTitle,Type,SensitivityLabel,RelativeUrl,ListTitle,MemberType,MemberName,MemberLoginName,ParentGroup,Roles|Export-CSV -Path $exportFilePath -NoTypeInformation;
  
  # Generate HTML report only if the option is enabled
  if ($generateHtml) {
    try {
      Write-LogHost "Generating HTML report..." -ForegroundColor Cyan
      Generate-HTMLReport -Permissions $global:permissions -ExportPath $exportFilePath -SiteTitle $global:siteTitle -DateTime $dateTime
    }
    catch {
      Write-LogHost "Error generating HTML report: $($_.Exception.Message)" -ForegroundColor Yellow
      Write-LogHost "The CSV file was generated successfully." -ForegroundColor Green
    }
  }
  else {
    Write-LogHost "HTML generation disabled by user option." -ForegroundColor Yellow
  }
  
  Disconnect-PnPOnline
  Write-LogHost "Script completed successfully!" -ForegroundColor Green
  Write-LogHost "Execution log saved to: $global:logFilePath" -ForegroundColor Cyan
  
}
else{
  Write-LogHost "Invalid directory path:" $logDirectoryPath -ForegroundColor "Red";
}

# Function to generate HTML report
function Generate-HTMLReport {
    param(
        [Parameter(Mandatory)]
        [array]$Permissions,
        [Parameter(Mandatory)]
        [string]$ExportPath,
        [Parameter(Mandatory)]
        [string]$SiteTitle,
        [Parameter(Mandatory)]
        [string]$DateTime
    )
    
    try {
        Write-LogHost "   Calculating statistics..." -ForegroundColor Gray
        
        # Calculate statistics for charts
        $totalSubSites = ($Permissions | Where-Object { $_.Type -eq "Sub-Site" } | Select-Object -Property SiteUrl -Unique | Measure-Object).Count
        $totalLibraries = ($Permissions | Where-Object { $_.ListTitle -ne "" } | Select-Object -Property ListTitle -Unique | Measure-Object).Count
        $librariesWithBrokenInheritance = ($Permissions | Where-Object { $_.ListTitle -ne "" -and $_.HasUniqueRoleAssignments -eq $true } | Select-Object -Property ListTitle -Unique | Measure-Object).Count
        
        # Count unique items by RelativeUrl (same link)
        $uniqueListItems = ($Permissions | Where-Object { $_.Type -eq "List Item" } | Select-Object -Property RelativeUrl -Unique | Measure-Object).Count
        $uniqueFolders = ($Permissions | Where-Object { $_.Type -eq "Folder" } | Select-Object -Property RelativeUrl -Unique | Measure-Object).Count
        $uniqueDocuments = ($Permissions | Where-Object { $_.Type -eq "File" } | Select-Object -Property RelativeUrl -Unique | Measure-Object).Count
        
        Write-LogHost "   Generating HTML..." -ForegroundColor Gray

        # Create HTML
        $htmlContent = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>SharePoint Permissions Report - $SiteTitle</title>
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
            padding: 20px;
        }
        
        .container {
            max-width: 1400px;
            margin: 0 auto;
            background: white;
            border-radius: 15px;
            box-shadow: 0 20px 40px rgba(0,0,0,0.1);
            overflow: hidden;
        }
        
        .header {
            background: linear-gradient(135deg, #2c3e50 0%, #34495e 100%);
            color: white;
            padding: 30px;
            text-align: center;
        }
        
        .header h1 {
            font-size: 2.5em;
            margin-bottom: 10px;
            font-weight: 300;
        }
        
        .header p {
            font-size: 1.1em;
            opacity: 0.9;
        }
        
        .stats-section {
            padding: 30px;
            background: #f8f9fa;
        }
        
        .stats-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
            gap: 20px;
            margin-bottom: 30px;
        }
        
        .stat-card {
            background: white;
            padding: 25px;
            border-radius: 10px;
            box-shadow: 0 5px 15px rgba(0,0,0,0.08);
            text-align: center;
        }
        
        .stat-card h3 {
            color: #2c3e50;
            margin-bottom: 15px;
            font-size: 1.2em;
        }
        
        .stat-number {
            font-size: 2.5em;
            font-weight: bold;
            color: #3498db;
            margin-bottom: 10px;
        }
        
        .stat-description {
            color: #7f8c8d;
            font-size: 0.9em;
        }
        
        .data-section {
            padding: 30px;
            background: white;
        }
        
        .search-container {
            margin-bottom: 20px;
        }
        
        .search-input {
            width: 100%;
            padding: 12px 20px;
            border: 2px solid #e0e0e0;
            border-radius: 25px;
            font-size: 16px;
            outline: none;
            transition: border-color 0.3s;
        }
        
        .search-input:focus {
            border-color: #3498db;
        }
        
        .data-table {
            width: 100%;
            border-collapse: collapse;
            background: white;
            border-radius: 10px;
            overflow: hidden;
            box-shadow: 0 5px 15px rgba(0,0,0,0.08);
        }
        
        .data-table th {
            background: linear-gradient(135deg, #2c3e50 0%, #34495e 100%);
            color: white;
            padding: 15px;
            text-align: left;
            font-weight: 500;
        }
        
        .data-table td {
            padding: 12px 15px;
            border-bottom: 1px solid #e0e0e0;
        }
        
        .data-table tr:hover {
            background: #f8f9fa;
        }
        
        .data-table tr:nth-child(even) {
            background: #fafafa;
        }
        
        .footer {
            background: #2c3e50;
            color: white;
            text-align: center;
            padding: 20px;
            font-size: 0.9em;
        }
        
        .badge {
            padding: 4px 8px;
            border-radius: 12px;
            font-size: 0.8em;
            font-weight: bold;
        }
        
        .badge-site { background: #3498db; color: white; }
        .badge-sub-site { background: #2980b9; color: white; }
        .badge-file { background: #e74c3c; color: white; }
        .badge-folder { background: #f39c12; color: white; }
        .badge-list-item { background: #8e44ad; color: white; }
        .badge-user { background: #27ae60; color: white; }
        .badge-group { background: #9b59b6; color: white; }
        
        .link-icon {
            display: inline-block;
            padding: 8px;
            background: #3498db;
            color: white;
            text-decoration: none;
            border-radius: 50%;
            font-size: 14px;
            transition: background-color 0.3s;
        }
        
        .link-icon:hover {
            background: #2980b9;
            color: white;
            text-decoration: none;
        }
        
        .expand-btn {
            background: #2c3e50;
            color: white;
            border: none;
            border-radius: 50%;
            width: 25px;
            height: 25px;
            font-size: 12px;
            cursor: pointer;
            transition: all 0.3s;
            display: inline-flex;
            align-items: center;
            justify-content: center;
        }
        
        .expand-btn:hover {
            background: #34495e;
            transform: scale(1.1);
        }
        
        .group-header {
            background: #f8f9fa !important;
            border-left: 4px solid #3498db;
            font-weight: 500;
        }
        
        .group-child {
            background: #fdfdfd !important;
            border-left: 4px solid #ecf0f1;
        }
        
        .group-child.hidden {
            display: none;
        }
        
        @media (max-width: 768px) {
            .stats-grid {
                grid-template-columns: 1fr;
            }
            
            .header h1 {
                font-size: 2em;
            }
        }
        
        @media (min-width: 1200px) {
            .stats-grid {
                grid-template-columns: repeat(3, 1fr);
            }
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>📊 SharePoint Permissions Report</h1>
            <p>Site Collection: $SiteTitle | Generated on: $DateTime</p>
        </div>
        
        <div class="stats-section">
            <div class="stats-grid">
                <div class="stat-card">
                    <h3>📈 Total Permissions</h3>
                    <div class="stat-number">$($Permissions.Count)</div>
                    <div class="stat-description">Unique permissions found</div>
                </div>
                <div class="stat-card">
                    <h3>🏢 Sub-Sites</h3>
                    <div class="stat-number">$totalSubSites</div>
                    <div class="stat-description">Sub-sites with unique permissions</div>
                </div>
                <div class="stat-card">
                    <h3>📚 Libraries/Lists</h3>
                    <div class="stat-number">$totalLibraries</div>
                    <div class="stat-description">Unique Libraries/Lists with permissions</div>
                </div>
                <div class="stat-card">
                    <h3>📋 List Items</h3>
                    <div class="stat-number">$uniqueListItems</div>
                    <div class="stat-description">Unique list items with permissions</div>
                </div>
                <div class="stat-card">
                    <h3>📁 Folders</h3> 
                    <div class="stat-number">$uniqueFolders</div>
                    <div class="stat-description">Unique folders with permissions</div>
                </div>
                <div class="stat-card">
                    <h3>📄 Documents</h3>
                    <div class="stat-number">$uniqueDocuments</div>
                    <div class="stat-description">Unique documents with permissions</div>
                </div>
            </div>
        </div>
        
        <div class="data-section">
            <div class="search-container">
                <input type="text" id="searchInput" class="search-input" placeholder="🔍 Search permissions...">
            </div>
            <table class="data-table" id="permissionsTable">
                <thead>
                    <tr>
                        <th width="40px"></th>
                        <th>Site</th>
                        <th>Type</th>
                        <th>Item Name</th>
                        <th>Member</th>
                        <th>Member Type</th>
                        <th>Permissions</th>
                        <th>Link</th>
                    </tr>
                </thead>
                <tbody>
"@

    # Helper function to determine the item name based on type
    function Get-ItemName($permission) {
        switch ($permission.Type) {
            "Site" { return $permission.SiteTitle }
            "Sub-Site" { return $permission.SiteTitle }
            { $_ -in @("DocumentLibrary", "GenericList") } { return $permission.ListTitle }
            "List Item" {
                # For list items, extract the item name/ID from the URL
                $url = $permission.RelativeUrl
                if ($url -match '.*/([^/]+)$') {
                    return $matches[1]
                } else {
                    return "List Item"
                }
            }
            { $_ -in @("File", "Folder") } {
                # Extract file/folder name from the URL
                $url = $permission.RelativeUrl
                if ($url -match '.*/([^/]+)$') {
                    return $matches[1]
                } else {
                    return $permission.ListTitle
                }
            }
            default { return $permission.ListTitle }
        }
    }
    
    # Group permissions by RelativeUrl
    Write-LogHost "   Grouping permissions by URL..." -ForegroundColor Gray
    $groupedPermissions = $Permissions | Group-Object -Property RelativeUrl
    
    # Add grouped table data
    $groupIndex = 0
    foreach ($group in $groupedPermissions) {
        $groupIndex++
        $groupId = "group_$groupIndex"
        $firstPermission = $group.Group[0]
        $groupCount = $group.Count
        
        # First line of the group (always visible)
        $typeBadge = switch ($firstPermission.Type) {
            "Site" { "badge-site" }
            "Sub-Site" { "badge-sub-site" }
            "File" { "badge-file" }
            "Folder" { "badge-folder" }
            "List Item" { "badge-list-item" }
            default { "badge-site" }
        }
        
        $memberTypeBadge = switch ($firstPermission.MemberType) {
            "User" { "badge-user" }
            "Group" { "badge-group" }
            default { "badge-user" }
        }
        
        # Expansion button (only appears if there's more than 1 item)
        $expandButton = ""
        if ($groupCount -gt 1) {
            $expandButton = "<button class='expand-btn' onclick='toggleGroup(""$groupId"")' id='btn_$groupId'>+</button>"
        }
        
        $itemName = Get-ItemName $firstPermission
        
        $htmlContent += @"
                    <tr class="group-header" data-group="$groupId">
                        <td>$expandButton</td>
                        <td>$($firstPermission.SiteTitle)</td>
                        <td><span class="badge $typeBadge">$($firstPermission.Type)</span></td>
                        <td>$itemName</td>
                        <td>$($firstPermission.MemberName)</td>
                        <td><span class="badge $memberTypeBadge">$($firstPermission.MemberType)</span></td>
                        <td>$($firstPermission.Roles)</td>
                        <td><a href="$($firstPermission.RelativeUrl)" target="_blank" class="link-icon" title="Open link">🔗</a></td>
                    </tr>
"@
        
        # Child lines of the group (initially hidden)
        if ($groupCount -gt 1) {
            for ($i = 1; $i -lt $groupCount; $i++) {
                $permission = $group.Group[$i]
                
                $typeBadge = switch ($permission.Type) {
                    "Site" { "badge-site" }
                    "Sub-Site" { "badge-sub-site" }
                    "File" { "badge-file" }
                    "Folder" { "badge-folder" }
                    "List Item" { "badge-list-item" }
                    default { "badge-site" }
                }
                
                $memberTypeBadge = switch ($permission.MemberType) {
                    "User" { "badge-user" }
                    "Group" { "badge-group" }
                    default { "badge-user" }
                }
                
                $childItemName = Get-ItemName $permission
                
                $htmlContent += @"
                    <tr class="group-child hidden" data-group="$groupId">
                        <td></td>
                        <td>$($permission.SiteTitle)</td>
                        <td><span class="badge $typeBadge">$($permission.Type)</span></td>
                        <td>$childItemName</td>
                        <td>$($permission.MemberName)</td>
                        <td><span class="badge $memberTypeBadge">$($permission.MemberType)</span></td>
                        <td>$($permission.Roles)</td>
                        <td><a href="$($permission.RelativeUrl)" target="_blank" class="link-icon" title="Open link">🔗</a></td>
                    </tr>
"@
            }
        }
    }

    $htmlContent += @"
                </tbody>
            </table>
        </div>
        
    <div class="footer">
    <p>Report automatically generated by SharePoint Permissions Audit Script</p>
    <p>Initially developed by <a href="https://nobre.cloud" target="_blank">https://nobre.cloud</a></p>
</div>

    <script>
        // Wait for DOM loading
        document.addEventListener('DOMContentLoaded', function() {
            // Add tooltips to buttons
            const buttons = document.querySelectorAll('.expand-btn');
            buttons.forEach(button => {
                button.title = 'Expand group';
            });
        });
        
        // Grouping functionality
        function toggleGroup(groupId) {
            const childRows = document.querySelectorAll('tr[data-group="' + groupId + '"].group-child');
            const button = document.getElementById('btn_' + groupId);
            
            if (!button) return;
            
            let isExpanded = button.textContent === '-';
            
            childRows.forEach(row => {
                if (isExpanded) {
                    row.classList.add('hidden');
                } else {
                    row.classList.remove('hidden');
                }
            });
            
            button.textContent = isExpanded ? '+' : '-';
            button.title = isExpanded ? 'Expand group' : 'Collapse group';
        }
        
        // Enhanced search functionality to work with groups
        document.getElementById('searchInput').addEventListener('keyup', function() {
            const searchTerm = this.value.toLowerCase();
            const table = document.getElementById('permissionsTable');
            const rows = table.getElementsByTagName('tbody')[0].getElementsByTagName('tr');
            const visibleGroups = [];

            // First, check which lines match the search
            for (let i = 0; i < rows.length; i++) {
                const row = rows[i];
                const cells = row.getElementsByTagName('td');
                let found = false;

                for (let j = 0; j < cells.length; j++) {
                    const cellText = cells[j].textContent.toLowerCase();
                    if (cellText.includes(searchTerm)) {
                        found = true;
                        break;
                    }
                }

                if (found) {
                    const groupId = row.getAttribute('data-group');
                    if (groupId && visibleGroups.indexOf(groupId) === -1) {
                        visibleGroups.push(groupId);
                    }
                }
            }

            // Show/hide groups based on search
            for (let i = 0; i < rows.length; i++) {
                const row = rows[i];
                const groupId = row.getAttribute('data-group');
                
                if (searchTerm === '') {
                    // If no search term, show headers and maintain group state
                    if (row.classList.contains('group-header')) {
                        row.style.display = '';
                    } else if (row.classList.contains('group-child')) {
                        // Maintain original state of child groups
                        row.style.display = row.classList.contains('hidden') ? 'none' : '';
                    }
                } else {
                    // If there's a search term, show only relevant groups
                    if (visibleGroups.indexOf(groupId) !== -1) {
                        row.style.display = '';
                        // Expand groups that have matches
                        if (row.classList.contains('group-child')) {
                            row.classList.remove('hidden');
                        }
                        if (row.classList.contains('group-header')) {
                            const button = row.querySelector('.expand-btn');
                            if (button) {
                                button.textContent = '-';
                                button.title = 'Collapse group';
                            }
                        }
                    } else {
                        row.style.display = 'none';
                    }
                }
            }
        });
    </script>
</body>
</html>
"@

    # Save HTML file
    $htmlFilePath = $ExportPath -replace '\.csv$', '.html'
    $htmlContent | Out-File -FilePath $htmlFilePath -Encoding UTF8
    
    Write-LogHost "📊 HTML report generated: $htmlFilePath" -ForegroundColor Green
    Write-LogHost "🌐 Open the HTML file in your browser to view the complete report" -ForegroundColor Cyan
    }
    catch {
        Write-LogHost "❌ Error generating HTML report: $($_.Exception.Message)" -ForegroundColor Red
        throw
    }
}