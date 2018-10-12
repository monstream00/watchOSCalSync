# ----------------------------------------------------------------------------- 
# Script: Get-OutlookCalendar.ps1 
# Author: monstream00, ed wilson, msft 
# ----------------------------------------------------------------------------- 
Function Get-OutlookCalendar 
{ 
 Add-type -assembly "Microsoft.Office.Interop.Outlook" | out-null
 $olFolders = "Microsoft.Office.Interop.Outlook.OlDefaultFolders" -as [type]  
 $outlook = new-object -comobject outlook.application 
 $namespace = $outlook.GetNameSpace("MAPI") 
 $folder = $namespace.getDefaultFolder($olFolders::olFolderCalendar) 
 $Appointments = $folder.items
 $Appointments.IncludeRecurrences = $true
 $Appointments.Sort("[Start]")
 $Start = (Get-Date).ToShortDateString() + " 00:00"
 $End = (Get-Date).AddDays(+1).ToShortDateString() + " 00:00"
 $filter = "[MessageClass]='IPM.Appointment' AND [Start] >= '$Start' AND [End] <= '$End'"
 echo From $Start To $End
 foreach ($Appointment in $Appointments.Restrict($filter) ) {
	$Appointment | Select-Object -Property Subject, Start, End, Duration, Location
	$Body = @{
		Subject = $Appointment.Subject
		Start = $Appointment.Start
		End = $Appointment.End
		Location = $Appointment.Location
	}
	#Invoke-RestMethod -Method 'Post' -Headers @{Authorization =("Bearer "+ $Authorization.access_token)} -Uri $Url -Body $Body
 }
} #end function Get-OutlookCalendar
#https://blogs.technet.microsoft.com/ronba/2016/05/09/using-powershell-and-the-office-365-rest-api-with-oauth/
Add-Type -AssemblyName System.Windows.Forms
$Url = "https://monstream.vultr.com"
	[Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | out-null
    [Reflection.Assembly]::LoadWithPartialName("System.Drawing")
    $form = New-Object Windows.Forms.Form
    $form.text = "My Form"
    $form.size = New-Object Drawing.size @(700,600)
    $web = New-object System.Windows.Forms.webbrowser
    $web.location = New-object System.Drawing.Point(3,3)
    $web.minimumsize = new-object System.Drawing.Size(20,20)
    $web.size = New-object System.Drawing.size(780,530)
    $web.navigate($Url)
    $form.Controls.Add($web)
    $form.showdialog()
    $doc = $web.document
	echo $doc.Cookie
Get-OutlookCalendar #-Authorization $Authorization
