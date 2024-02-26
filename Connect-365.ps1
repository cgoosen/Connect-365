<#
.SYNOPSIS
  Connect to Office 365 services via PowerShell

.DESCRIPTION
  This script will prompt for your Office 365 tenant credentials and connect you to any or all Office 365 services via remote PowerShell

.INPUTS
  None

.OUTPUTS
  None

.NOTES
  Version:        1.3.1
  Author:         Chris Goosen (Twitter: @chrisgoosen)
  Creation Date:  23 Nov 2023
  Credits:        Ver >= 1.3 ExchangeMFAModule handling by Michel de Rooij - eightwone.com, @mderooij
                  Bugfinder extraordinaire Greig Sheridan - greiginsydney.com, @greiginsydney
                  Various bugfixes: Andy Helsby - github.com/Absoblogginlutely

.LINK
  http://www.cgoosen.com

.EXAMPLE
  .\Connect-365.ps1
#>
$ErrorActionPreference = "Stop"
$ScriptVersion = "1.3.1"
$ScriptName = "Connect365"
$ScriptDisplayName = "Connect-365"
$ScriptURL = "https://github.com/cgoosen/Connect-365/releases/"
$ScriptAuthor = "cgoosen"
$RegistryKeyPath = "HKCU:\Software\" + $ScriptAuthor + "\" + $ScriptName
#region XAML code
$XAML = @"
<Window
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="Connect-365" Height="420" Width="550" ResizeMode="NoResize" WindowStartupLocation="CenterScreen">
    <Grid>
        <DockPanel>
            <Menu DockPanel.Dock="Top">
                <MenuItem Header="_File">
                    <MenuItem Name="Btn_Exit" Header="_Exit" />
                </MenuItem>

                <MenuItem Header="_Edit">
                    <MenuItem Command="Cut" />
                    <MenuItem Command="Copy" />
                    <MenuItem Command="Paste" />
                </MenuItem>

                <MenuItem Header="_Help">
                    <MenuItem Header="_About">
                        <MenuItem Name="Btn_About" Header="_Script Version $ScriptVersion"/>
                        </MenuItem>
                    <MenuItem Name="Btn_Help" Header="_Get Help" />
                </MenuItem>
            </Menu>
        </DockPanel>
        <TabControl Margin="0,20,0,0">
            <TabItem Name="Tab_Connection" Header="Connection Options" TabIndex="12">
                <Grid Background="White">
                    <StackPanel>
                        <StackPanel Height="32" HorizontalAlignment="Center" VerticalAlignment="Top" Width="538" Margin="0,0,0,0">
                            <Label Content="Microsoft 365 Remote PowerShell" HorizontalAlignment="Center" VerticalAlignment="Top" Margin="0" Height="32" FontWeight="Bold"/>
                        </StackPanel>
                        <StackPanel Height="32" HorizontalAlignment="Center" VerticalAlignment="Top" Width="538" Margin="0,0,0,0" Orientation="Horizontal">
                            <Label Content="Username:" HorizontalAlignment="Left" Height="32" Margin="10,0,0,0" VerticalAlignment="Center" Width="70" FontSize="11" VerticalContentAlignment="Center"/>
                            <TextBox Name="Field_User" HorizontalAlignment="Left" Height="22" Margin="0,0,0,0" TextWrapping="Wrap" VerticalAlignment="Center" Width="438" VerticalContentAlignment="Center" FontSize="11" BorderThickness="1" TabIndex="1"/>
                        </StackPanel>
                        <StackPanel Height="32" HorizontalAlignment="Center" VerticalAlignment="Top" Width="538" Margin="0,0,0,0" Orientation="Horizontal">
                            <Label Content="Password:" HorizontalAlignment="Left" Height="32" Margin="10,0,0,0" VerticalAlignment="Center" Width="70" FontSize="11" VerticalContentAlignment="Center"/>
                            <PasswordBox Name="Field_Pwd" HorizontalAlignment="Left" Height="22" Margin="0,0,0,0" VerticalAlignment="Center" Width="438" VerticalContentAlignment="Center" FontSize="11" BorderThickness="1" TabIndex="2"/>
                        </StackPanel>
                        <StackPanel HorizontalAlignment="Center" VerticalAlignment="Top" Width="538" Margin="0,10,0,0">
                            <GroupBox Header="Services:" Width="508" Margin="10,0,0,0" FontSize="11" HorizontalAlignment="Left" VerticalAlignment="Top">
                                <Grid Height="60" Margin="0,10,0,0">
                                    <CheckBox Name="Box_EXO" TabIndex="3" HorizontalAlignment="Left" VerticalAlignment="Top">Exchange Online</CheckBox>
                                    <CheckBox Name="Box_AAD" TabIndex="4" HorizontalAlignment="Center" VerticalAlignment="Top">Azure AD</CheckBox>
                                    <CheckBox Name="Box_Com" TabIndex="5" HorizontalAlignment="Right" VerticalAlignment="Top">Compliance Center</CheckBox>
                                    <CheckBox Name="Box_SPO" TabIndex="6" HorizontalAlignment="Left" VerticalAlignment="Center">SharePoint Online</CheckBox>
                                    <CheckBox Name="Box_MSO" TabIndex="7" HorizontalAlignment="Center" VerticalAlignment="Center" Margin="50,0,0,0">Azure AD MSOnline</CheckBox>
                                    <CheckBox Name="Box_Teams" TabIndex="8" HorizontalAlignment="Right" VerticalAlignment="Center" Margin="0,0,62,0">Teams</CheckBox>
                                    <CheckBox Name="Box_Intune" TabIndex="9" HorizontalAlignment="Left" VerticalAlignment="Bottom">Intune</CheckBox>
                                </Grid>
                            </GroupBox>
                            <GroupBox Header="Options:" Width="508" Margin="10,10,0,0" FontSize="11" HorizontalAlignment="Left" VerticalAlignment="Top">
                                <Grid Height="50" Margin="0,10,0,0">
                                  <CheckBox Name="Box_MFA" TabIndex="10" HorizontalAlignment="Left" VerticalAlignment="Top" IsChecked="True" IsEnabled="False">Use MFA?</CheckBox>
                                  <CheckBox Name="Box_Clob" TabIndex="11" HorizontalAlignment="Center" VerticalAlignment="Top" IsEnabled="False" Margin="20,0,0,0">AllowClobber</CheckBox>
                                    <StackPanel HorizontalAlignment="Left" VerticalAlignment="Bottom" Orientation="Horizontal">
                                        <Label Content="Admin URL:" Width="70"></Label>
                                        <TextBox Name="Field_SPOUrl" Height="22" Width="425" Margin="0,0,0,0" TextWrapping="Wrap" IsEnabled="False" TabIndex="12"></TextBox>
                                    </StackPanel>
                                </Grid>
                            </GroupBox>
                        </StackPanel>
                        <StackPanel Height="45" Orientation="Horizontal" VerticalAlignment="Top" HorizontalAlignment="Center" Margin="0,10,0,0">
                            <Button Name="Btn_Ok" Content="Ok" Width="75" Height="25" VerticalAlignment="Top" HorizontalAlignment="Center" TabIndex="13" />
                            <Button Name="Btn_Cancel" Content="Cancel" Width="75" Height="25" VerticalAlignment="Top" HorizontalAlignment="Center" Margin="40,0,0,0" TabIndex="14" />
                        </StackPanel>
                    </StackPanel>
                </Grid>
            </TabItem>
            <TabItem Name="Tab_Prereq" Header="Prerequisite Checker" TabIndex="11">
                <Grid Background="White">
                    <StackPanel>
                        <StackPanel>
                            <Grid Margin="0,10,0,0">
                                <Label Content="Module" HorizontalAlignment="Left" FontSize="11" FontWeight="Bold"/>
                                <Label Content="Status" HorizontalAlignment="Center" FontSize="11" FontWeight="Bold"/>
                            </Grid>
                            <StackPanel>
                                <Label BorderBrush="Black" BorderThickness="0,0,0,1" VerticalAlignment="Top"/>
                            </StackPanel>
                        </StackPanel>
                        <StackPanel>
                            <Grid Margin="0,10,0,0">
                                <Label Content="Azure AD Version 2" HorizontalAlignment="Left" FontSize="11"/>
                                <TextBlock Name="Txt_AADStatus" HorizontalAlignment="Center" VerticalAlignment="Center" FontSize="11" />
                                <Button Name="Btn_AADMsg" Content="Download now.." Width="125" Height="25" HorizontalAlignment="Right" VerticalAlignment="Center" Margin="0,0,10,0" />
                            </Grid>
                        </StackPanel>
                        <StackPanel>
                            <Grid Margin="0,10,0,0">
                                <Label Content="SharePoint Online" HorizontalAlignment="Left" FontSize="11"/>
                                <TextBlock Name="Txt_SPOStatus" HorizontalAlignment="Center" VerticalAlignment="Center" FontSize="11" />
                                <Button Name="Btn_SPOMsg" Content="Download now.." Width="125" Height="25" HorizontalAlignment="Right" VerticalAlignment="Center" Margin="0,0,10,0" />
                            </Grid>
                        </StackPanel>
                        <StackPanel>
                            <Grid Margin="0,10,0,0">
                                <Label Content="Azure AD MSOnline" HorizontalAlignment="Left" FontSize="11"/>
                                <TextBlock Name="Txt_MSOStatus" HorizontalAlignment="Center" VerticalAlignment="Center" FontSize="11" />
                                <Button Name="Btn_MSOMsg" Content="Download now.." Width="125" Height="25" HorizontalAlignment="Right" VerticalAlignment="Center" Margin="0,0,10,0" />
                            </Grid>
                        </StackPanel>
                            <Grid Margin="0,10,0,0">
                                <Label Content="Exchange Online" HorizontalAlignment="Left" FontSize="11"/>
                                <TextBlock Name="Txt_EXOStatus" HorizontalAlignment="Center" VerticalAlignment="Center" FontSize="11" />
                            <Button Name="Btn_EXOMsg" Content="Download now.." Width="125" Height="25" HorizontalAlignment="Right" VerticalAlignment="Center" Margin="0,0,10,0" />
                        </Grid>
                        <StackPanel>
                          <Grid Margin="0,10,0,0">
                              <Label Content="Teams" HorizontalAlignment="Left" FontSize="11"/>
                              <TextBlock Name="Txt_TeamsStatus" HorizontalAlignment="Center" VerticalAlignment="Center" FontSize="11" />
                          <Button Name="Btn_TeamsMsg" Content="Download now.." Width="125" Height="25" HorizontalAlignment="Right" VerticalAlignment="Center" Margin="0,0,10,0" />
                    </Grid>
                        </StackPanel>
                        <StackPanel>
                          <Grid Margin="0,10,0,0">
                              <Label Content="Intune (MS Graph)" HorizontalAlignment="Left" FontSize="11"/>
                              <TextBlock Name="Txt_IntuneStatus" HorizontalAlignment="Center" VerticalAlignment="Center" FontSize="11" />
                          <Button Name="Btn_IntuneMsg" Content="Download now.." Width="125" Height="25" HorizontalAlignment="Right" VerticalAlignment="Center" Margin="0,0,10,0" />
                    </Grid>
                        </StackPanel>
                    </StackPanel>
                </Grid>
            </TabItem>
        </TabControl>
    </Grid>
</Window>
"@

#endregion

[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
[xml]$XAMLGui = $XAML

$Reader=(New-Object System.Xml.XmlNodeReader $XAMLGui)
$MainWindow=[Windows.Markup.XamlReader]::Load( $Reader )
$XAMLGui.SelectNodes("//*[@Name]") | ForEach-Object {Set-Variable -Name "GUI$($_.Name)" -Value $MainWindow.FindName($_.Name)}

# Functions

Function WriteToRegistry{

    try {
        New-Item -Path "HKCU:\Software" -Name $ScriptAuthor
        New-Item -Path ("HKCU:\Software\" + $ScriptAuthor) -Name $ScriptName

        New-ItemProperty -Path $RegistryKeyPath -PropertyType "String" -Name "Username"
        New-ItemProperty -Path $RegistryKeyPath -PropertyType "String" -Name "Service_EXO"    -Value "False"
        New-ItemProperty -Path $RegistryKeyPath -PropertyType "String" -Name "Service_AAD"    -Value "False"
        New-ItemProperty -Path $RegistryKeyPath -PropertyType "String" -Name "Service_Com"    -Value "False"
        New-ItemProperty -Path $RegistryKeyPath -PropertyType "String" -Name "Service_SPO"    -Value "False"
        New-ItemProperty -Path $RegistryKeyPath -PropertyType "String" -Name "Service_Teams"  -Value "False"
        New-ItemProperty -Path $RegistryKeyPath -PropertyType "String" -Name "Service_Intune" -Value "False"
    }
    catch {
    }

    Set-ItemProperty -Path $RegistryKeyPath -Name "Username"       -Value $GUIField_User.Text
    Set-ItemProperty -Path $RegistryKeyPath -Name "Service_EXO"    -Value $GUIBox_EXO.IsChecked
    Set-ItemProperty -Path $RegistryKeyPath -Name "Service_AAD"    -Value $GUIBox_AAD.IsChecked
    Set-ItemProperty -Path $RegistryKeyPath -Name "Service_Com"    -Value $GUIBox_Com.IsChecked
    Set-ItemProperty -Path $RegistryKeyPath -Name "Service_SPO"    -Value $GUIBox_SPO.IsChecked
    Set-ItemProperty -Path $RegistryKeyPath -Name "Service_MSO"    -Value $GUIBox_MSO.IsChecked
    Set-ItemProperty -Path $RegistryKeyPath -Name "Service_Teams"  -Value $GUIBox_Teams.IsChecked
    Set-ItemProperty -Path $RegistryKeyPath -Name "Service_Intune" -Value $GUIBox_Intune.IsChecked
}

Function ReadFromRegistry{

    Write-Host "ReadFromRegistry"

    try {
        $GUIField_User.Text = Get-ItemPropertyValue -Path $RegistryKeyPath -Name "Username"

        if ( (Get-ItemPropertyValue -Path $RegistryKeyPath -Name "Service_EXO")    -eq "True" ) { $GUIBox_EXO.IsChecked = 1 }
        if ( (Get-ItemPropertyValue -Path $RegistryKeyPath -Name "Service_AAD")    -eq "True" ) { $GUIBox_AAD.IsChecked = 1 }
        if ( (Get-ItemPropertyValue -Path $RegistryKeyPath -Name "Service_Com")    -eq "True" ) { $GUIBox_Com.IsChecked = 1 }
        if ( (Get-ItemPropertyValue -Path $RegistryKeyPath -Name "Service_SPO")    -eq "True" ) { $GUIBox_SPO.IsChecked = 1 }
        if ( (Get-ItemPropertyValue -Path $RegistryKeyPath -Name "Service_MSO")    -eq "True" ) { $GUIBox_MSO.IsChecked = 1 }
        if ( (Get-ItemPropertyValue -Path $RegistryKeyPath -Name "Service_Teams")  -eq "True" ) { $GUIBox_Teams.IsChecked = 1 }
        if ( (Get-ItemPropertyValue -Path $RegistryKeyPath -Name "Service_Intune") -eq "True" ) { $GUIBox_Intune.IsChecked = 1 }
    }
    catch {
    }
}

Function Get-ScriptVersion{
    [CmdletBinding()]
  param(
      [Parameter()]
      [string]$CurrentVersion,
      [string]$ScriptName
  )
  $LatestVersion = Invoke-RestMethod -Method Get -Uri https://cgoosen.azure-api.net/Versions/GetVersion?Name=$ScriptName
  If (!$LatestVersion -or $LatestVersion -eq "Error: Something went wrong.."){
    Write-Host "Unable to perform version check, visit $ScriptURL to check if you're running the latest version" -ForegroundColor Red
  }
  Else {
    If ($LatestVersion -gt $CurrentVersion){
    Write-Host "A newer version of $ScriptDisplayName is available, visit $ScriptURL to download the latest version" -ForegroundColor Red
    }
    Elseif ($LatestVersion -eq $CurrentVersion){
      Write-Host "You are running the latest version of $ScriptDisplayName" -ForegroundColor Green
    }
  }
}

Function Get-Options{
        If ($GUIBox_EXO.IsChecked -eq "True") {
            $Script:ConnectEXO = $true
            $OptionsArray++
    }
        If ($GUIBox_AAD.IsChecked -eq "True") {
            $Script:ConnectAAD = $true
            $OptionsArray ++
    }
        If ($GUIBox_Com.IsChecked -eq "True") {
            $Script:ConnectCom = $true
            $OptionsArray++
    }
        If ($GUIBox_MSO.IsChecked -eq "True") {
            $Script:ConnectMSO = $true
            $OptionsArray++
    }
        If ($GUIBox_SPO.IsChecked -eq "True") {
            $Script:ConnectSPO = $true
            $OptionsArray++
    }
        If ($GUIBox_Teams.IsChecked -eq "True") {
            $Script:ConnectTeams = $true
            $OptionsArray++
    }
        If ($GUIBox_Intune.IsChecked -eq "True") {
            $Script:ConnectIntune = $true
            $OptionsArray++
    }
        If ($GUIBox_MFA.IsChecked -eq "True") {
            $Script:UseMFA = $true
    }
}

Function Get-UserPwd{
  #Password no longer needed as modern auth workflow will prompt for it again
        If (!$Username) {
            $MainWindow.Close()
            Close-Window "Please enter a valid UserName..`nScript failed"
    }
        ElseIf ($OptionsArray -eq "0") {
            $MainWindow.Close()
            Close-Window "Please select a valid option..`nScript failed"
    }
}

Function Connect-EXO{
    If (Get-ModuleInfo-EXOv3 -eq "True") {
        Connect-ExchangeOnline -UserPrincipalName $UserName -ShowBanner:$false
    }
   ElseIf (Get-ModuleInfo-EXO -eq "True") {
        $EXOSession = New-EXOPSSession -ConnectionUri https://outlook.office365.com/powershell-liveid/ -UserPrincipalName $UserName
        If ($Clob) {
        Import-PSSession $EXOSession -AllowClobber
        }
        Else {
        Import-PSSession $EXOSession
        }
   }
}

Function Connect-AAD{
  Connect-AzureAD -AccountId $UserName
}

Function Connect-Com{
    $CCSession = New-EXOPSSession -ConnectionUri https://ps.compliance.protection.outlook.com/powershell-liveid/ -UserPrincipalName $UserName
    If ($Clob) {
      Import-PSSession $CCSession -AllowClobber
    }
    Else {
      Import-PSSession $CCSession
    }
}

Function Connect-MSO{
    Connect-MsolService
}

Function Connect-SPO{
    Connect-SPOService -Url $GUIField_SPOUrl.text
}

Function Connect-Teams{
    Connect-MicrosoftTeams -AccountId $UserName
}

Function Connect-Intune{
    Connect-MSGraph
}

Function Get-ModuleInfo-AAD{
      try {
          Import-Module -Name AzureAD
          return $true
      }
      catch {
          return $false
      }
}

Function Get-ModuleInfo-MSO{
      try {
          Import-Module -Name MSOnline
          return $true
      }
      catch {
          return $false
      }
}

Function Get-ModuleInfo-SPO{
    try {
        Import-Module Microsoft.Online.SharePoint.PowerShell -DisableNameChecking
        return $true
    }
    catch {
        return $false
    }
}

Function Get-ModuleInfo-EXO{
    try {
        Import-Module ExchangeOnlineManagement
        return $true
    }
    catch {
        return $false
    }
}

Function Get-ModuleInfo-EXOv3{
    try {
        Import-Module -Name ExchangeOnlineManagement
        return $true
    }
    catch {
        return $false
    }
}

Function Get-ModuleInfo-Teams{
    try {
        Import-Module -Name MicrosoftTeams
        return $true
    }
    catch {
        return $false
    }
}

Function Get-ModuleInfo-Intune{
    try {
        Import-Module -Name Microsoft.Graph.Intune
        return $true
    }
    catch {
        return $false
    }
}


function Close-Window ($CloseReason) {
    Write-Host "$CloseReason" -ForegroundColor Red
    Exit
}

function Get-FailedMsg ($FailedReason) {
    Write-Host "$FailedReason. Connection failed, please check your credentials and try again.." -ForegroundColor Red
    Exit
}

function Get-PreReq-AAD{
    If (Get-ModuleInfo-AAD -eq "True") {
        $GUITxt_AADStatus.Text = "OK!"
        $GUITxt_AADStatus.Foreground = "Green"
        $GUIBtn_AADMsg.IsEnabled = $false
        $GUIBtn_AADMsg.Opacity = "0"
    }
    else {
        $GUITxt_AADStatus.Text = "Failed!"
        $GUITxt_AADStatus.Foreground = "Red"
        $GUIBtn_AADMsg.IsEnabled = $true
    }
}

function Get-PreReq-MSO{
    If (Get-ModuleInfo-MSO -eq "True") {
        $GUITxt_MSOStatus.Text = "OK!"
        $GUITxt_MSOStatus.Foreground = "Green"
        $GUIBtn_MSOMsg.IsEnabled = $false
        $GUIBtn_MSOMsg.Opacity = "0"
    }
    else {
        $GUITxt_MSOStatus.Text = "Failed!"
        $GUITxt_MSOStatus.Foreground = "Red"
        $GUIBtn_MSOMsg.IsEnabled = $true
    }
}

function Get-PreReq-SPO{
    If (Get-ModuleInfo-SPO -eq "True") {
        $GUITxt_SPOStatus.Text = "OK!"
        $GUITxt_SPOStatus.Foreground = "Green"
        $GUIBtn_SPOMsg.IsEnabled = $false
        $GUIBtn_SPOMsg.Opacity = "0"
    }
    else {
        $GUITxt_SPOStatus.Text = "Failed!"
        $GUITxt_SPOStatus.Foreground = "Red"
        $GUIBtn_SPOMsg.IsEnabled = $true
    }
}

function Get-PreReq-EXO{
    If (Get-ModuleInfo-EXO -eq "True" -or Get-ModuleInfo-EXOv3 -eq "True") {
        $GUITxt_EXOStatus.Text = "OK!"
        $GUITxt_EXOStatus.Foreground = "Green"
        $GUIBtn_EXOMsg.IsEnabled = $false
        $GUIBtn_EXOMsg.Opacity = "0"
    }
    else {
        $GUITxt_EXOStatus.Text = "Failed!"
        $GUITxt_EXOStatus.Foreground = "Red"
        $GUIBtn_EXOMsg.IsEnabled = $true
    }
}

function Get-PreReq-Teams{
    If (Get-ModuleInfo-Teams -eq "True") {
        $GUITxt_TeamsStatus.Text = "OK!"
        $GUITxt_TeamsStatus.Foreground = "Green"
        $GUIBtn_TeamsMsg.IsEnabled = $false
        $GUIBtn_TeamsMsg.Opacity = "0"
    }
    else {
        $GUITxt_TeamsStatus.Text = "Failed!"
        $GUITxt_TeamsStatus.Foreground = "Red"
        $GUIBtn_TeamsMsg.IsEnabled = $true
    }
}

function Get-PreReq-Intune{
    If (Get-ModuleInfo-Intune -eq "True") {
        $GUITxt_IntuneStatus.Text = "OK!"
        $GUITxt_IntuneStatus.Foreground = "Green"
        $GUIBtn_IntuneMsg.IsEnabled = $false
        $GUIBtn_IntuneMsg.Opacity = "0"
    }
    else {
        $GUITxt_IntuneStatus.Text = "Failed!"
        $GUITxt_IntuneStatus.Foreground = "Red"
        $GUIBtn_IntuneMsg.IsEnabled = $true
    }
}
function Get-PreReq{
  Get-PreReq-AAD
  Get-PreReq-MSO
  Get-PreReq-SPO
  Get-PreReq-EXO
  Get-PreReq-Teams
  Get-PreReq-Intune
}

function Get-OKBtn{
  $Script:Username = $GUIField_User.Text
  $Passwd = $GUIField_Pwd.Password
  Get-Options
  Get-UserPwd
  If ($Passwd) {
  	$EncryptPwd = $Passwd | ConvertTo-SecureString -AsPlainText -Force
  	$Script:Credential = New-Object System.Management.Automation.PSCredential($Username,$EncryptPwd)
  }

  WriteToRegistry

  $Script:EndScript = 2
	$MainWindow.Close()
}

function Get-CancelBtn{
    $MainWindow.Close()
    $Script:EndScript = 1
	Close-Window 'Script cancelled'
}

# Event Handlers
$MainWindow.add_KeyDown({
    param
(
  [Parameter(Mandatory)][Object]$Sender,
  [Parameter(Mandatory)][Windows.Input.KeyEventArgs]$KeyPress
)
    if ($KeyPress.Key -eq "Enter"){
    Get-OKBtn
    }

    if ($KeyPress.Key -eq "Escape"){
    Get-CancelBtn
    }
})

$MainWindow.add_Loaded({
    ReadFromRegistry
})

$MainWindow.add_Closing({
    $Script:EndScript++
})

$GUIBtn_Cancel.add_Click({
    Get-CancelBtn
})

$GUIBtn_Ok.add_Click({
    Get-OKBtn
})

$GUITab_Prereq.add_Loaded({

})

$GUIBtn_AADMsg.add_Click({
    try {
        Start-Process -FilePath https://www.powershellgallery.com/packages/AzureAD
    }
    catch {
        $MainWindow.Close()
        Close-Window "An error occurred..`nExiting script"
    }
})

$GUIBtn_MSOMsg.add_Click({
    try {
        Start-Process -FilePath https://www.powershellgallery.com/packages/MSOnline
    }
    catch {
        $MainWindow.Close()
        Close-Window "An error occurred..`nExiting script"
    }
})

$GUIBtn_SPOMsg.add_Click({
    try {
        Start-Process -FilePath http://go.microsoft.com/fwlink/p/?LinkId=255251
    }
    catch {
        $MainWindow.Close()
        Close-Window "An error occurred..`nExiting script"
    }
})

$GUIBtn_EXOMsg.add_Click({
    try {
        Start-Process -FilePath https://www.powershellgallery.com/packages/ExchangeOnlineManagement
    }
    catch {
        $MainWindow.Close()
        Close-Window "An error occurred..`nExiting script"
    }
})

$GUIBtn_TeamsMsg.add_Click({
    try {
        Start-Process -FilePath https://www.powershellgallery.com/packages/MicrosoftTeams
    }
    catch {
        $MainWindow.Close()
        Close-Window "An error occurred..`nExiting script"
    }
})

$GUIBtn_IntuneMsg.add_Click({
    try {
        Start-Process -FilePath https://github.com/Microsoft/Intune-PowerShell-SDK
    }
    catch {
        $MainWindow.Close()
        Close-Window "An error occurred..`nExiting script"
    }
})

$GUIBox_EXO.add_Click({
    $GUIBox_Clob.IsEnabled = "True"
})

$GUIBox_Com.add_Click({
    $GUIBox_Clob.IsEnabled = "True"
})


$GUIBox_SPO.add_Checked({
    $GUIField_SPOUrl.IsEnabled = "True"
    $GUIField_SPOUrl.Text = "Enter your SharePoint Online Admin URL, e.g https://<tenant>-admin.sharepoint.com"
})

$GUIBox_SPO.add_UnChecked({
    $GUIField_SPOUrl.IsEnabled = "False"
    $GUIField_SPOUrl.Text = ""
})

$GUIField_SPOUrl.add_GotFocus({
    $GUIField_SPOUrl.Text = ""
})

$GUIBtn_Exit.add_Click({
    Get-CancelBtn
})

$GUIBtn_About.add_Click({
    Start-Process -FilePath http://cgoo.se/2ogotCK
})

$GUIBtn_Help.add_Click({
    Start-Process -FilePath http://cgoo.se/1srvTiS
})

# Script re-req checks
Write-Host "Starting script version $ScriptVersion..`nLooking for newer version.." -ForegroundColor Green
Get-ScriptVersion -CurrentVersion $ScriptVersion -ScriptName $ScriptName
Write-Host "Done!" -ForegroundColor Green
Write-Host "Looking for installed modules.." -ForegroundColor Green
Get-PreReq
Write-Host "Done!" -ForegroundColor Green

# Load GUI Window
$MainWindow.WindowStartupLocation = "CenterScreen"
$MainWindow.ShowDialog() | Out-Null

# Check if Window is closed
If ($EndScript -eq 1){
    Close-Window 'Script cancelled'
}

# Connect to EXO if required
If ($ConnectEXO -eq "True"){
        Try {
            Connect-EXO
        }
        Catch 	{
            Get-FailedMsg 'Exchange Online error'
        }
}

# Connect to SharePoint Online if required
If ($ConnectSPO-eq "True"){
    Try {
        Connect-SPO
    }
    Catch 	{
        Get-FailedMsg 'SharePoint Online error'
    }
}

# Connect to Security & Compliance Center if required
If ($ConnectCom -eq "True"){
    Try {
        Start-Sleep -Seconds 2
        Connect-Com
    }
    Catch 	{
        Get-FailedMsg 'Security & Compliance Center error'
    }
}

# Connect to AAD if required
If ($ConnectAAD -eq "True"){
    Try {
        Connect-AAD
    }
    Catch 	{
        Get-FailedMsg 'Azure AD error'
    }
}

# Connect to Teams if required
If ($ConnectTeams -eq "True"){
    Try {
        Connect-Teams
    }
    Catch 	{
        Get-FailedMsg 'Teams error'
    }
}

# Connect to Intune if required
If ($ConnectIntune -eq "True"){
    Try {
        Connect-Intune
    }
    Catch 	{
        Get-FailedMsg 'Intune error'
    }
}

# Connect to Azure AD MSOnline if required
If ($ConnectMSO -eq "True"){
    Try {
        Connect-MSO
    }
    Catch 	{
        Get-FailedMsg 'Azure AD MSOnline error'
    }
}

# Notifications/Information
Clear-Host
Write-Host "
Your username is: $UserName" -ForegroundColor Yellow -BackgroundColor Black
Write-Host "You are now connected to:" -ForegroundColor Yellow -BackgroundColor Black
If ($ConnectEXO -eq "True"){
    Write-Host "-Exchange Online" -ForegroundColor Yellow -BackgroundColor Black
}
If ($ConnectAAD -eq "True"){
    Write-Host "-Azure Active Directory" -ForegroundColor Yellow -BackgroundColor Black
}
If ($ConnectCom -eq "True"){
    Write-Host "-Office 365 Security & Compliance Center" -ForegroundColor Yellow -BackgroundColor Black
}
If ($ConnectMSO -eq "True"){
    Write-Host "-Azure AD MSOnline" -ForegroundColor Yellow -BackgroundColor Black
}
If ($ConnectSPO -eq "True"){
    Write-Host "-SharePoint Online" -ForegroundColor Yellow -BackgroundColor Black
}
If ($ConnectTeams -eq "True"){
    Write-Host "-Teams" -ForegroundColor Yellow -BackgroundColor Black
}
If ($ConnectIntune -eq "True"){
    Write-Host "-Intune API" -ForegroundColor Yellow -BackgroundColor Black
}
