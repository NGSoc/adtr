<# 
.NAME
    ADRT - Active Directory Report Tool
.DESCRIPTION
    Extract the complete list of all Member Groups.
.EXAMPLE
    PS C:\adrt> .\ad-membergroups.ps1
.NOTES
    Name: Thaigo Freitas
	E-mail: brntt0012@ambev.com.br
.LINK
    
#>

$report = $null
$table = $null
$date = Get-Date -format "yyyy-M-d"
$mounth = Get-Date -format "MMM"
$directorypath = (Get-Item -Path ".\").FullName
$path = "ad-reports\ad-membergroups"
$html = "$path\ad-membergroups-$date.html"
$csv = "$path\ad-membergroups-$date.csv"

#-- Member Groups
$t_mg = (Get-ADGroup -Filter {Name -like "*"} -Properties *).count 
$domain = (Get-ADDomain).Forest

# Config
#$config = Get-Content (JOIN-PATH $directorypath "config\config.txt")
#$company = $config[7]
#$owner = $config[9]

#-- Import Module
Import-Module ActiveDirectory

#-- Show Total
$table += "<center><h3><b>Total Groups: <font color=red>$t_mg</font></b></h3></center>"

#-- Filter
$membergroups = @(Get-ADGroup -Filter {Name -like "*"} -Properties *)

$result = @($membergroups | Select-Object Name, @{n='MemberOf'; e= { $_.memberof | Out-String}}, @{n='Members'; e= { $_.members | Out-String}})

#-- Order by (A-Z)
$result = $result | Sort "Name"

#-- Display result on screen
#$result | ft -auto 

$table += $result | ConvertTo-Html -Fragment
 
$format=
		"
		<html>
		<body>
		<title>$company</title>
		<style>
		BODY{font-family: Calibri; font-size: 12pt;}
		TABLE{border: 1px solid black; border-collapse: collapse; font-size: 12pt; text-align:center;margin-left:auto;margin-right:auto; width='1000px';}
		TH{border: 1px solid black; background: #F9F9F9; padding: 5px;}
		TD{border: 1px solid black; padding: 5px;}
		H3{font-family: Calibri; font-size: 12pt;}
		</style> 
		"
$title=
		"
		<table width='100%' border='0' cellpadding='0' cellspacing='0'>
		<tr>
		<td bgcolor='#F9F9F9'>
		<font face='Calibri' size='5px'><b>Active Directory - Member Groups</b></font>
		<H3 align='center'>Company: <font color=red>$company</font> - Domain: <font color=red>$domain</font> - Date: <font color=red>$date</font> - Owner: <font color=red>$owner</font></H3>
		</td>
		</tr>
		</table>
		</body>
		</html>
		"
$footer=
		"
		<br><br>
		<table width='100%' border='0' cellpadding='0' cellspacing='0'>
		<tr>
		<td bgcolor='#F9F9F9'>
		<font face='Calibri' size='2px'>ADRT - Active Directory Report Tool</font>
		</td>
		</tr>
		</table>
		"
$message = "</table><style>"
$message = $message + "BODY{font-family: Calibri;font-size:16;font-color: #000000}"
$message = $message + "TABLE{margin-left:auto;margin-right:auto;width: 800px;border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}"
$message = $message + "TH{border-width: 1px;padding: 0px;border-style: solid;border-color: black;background-color: #F9F9F9;text-align:center;}"
$message = $message + "TD{border-width: 1px;padding: 0px;border-style: solid;border-color: black;text-align:center;}"
$message = $message + "</style>"
$message = $message + "<table width='300px' heigth='500px' align='center'>"
$message = $message + "<tr><td colspan='2' bgcolor='#DDEBF7' height='40'><b>Active Directory</b></td></tr>"
$message = $message + "<tr><td bgcolor='#F9F9F9' height='40'>Description</td><td bgcolor='#F9F9F9' height='40'>Total</td></tr>"
$message = $message + "<tr><td height='40'>Member Groups</td><td>$t_mg</td></tr>"
$message = $message + "<tr><td colspan='2' bgcolor='#DDEBF7' height='40'><b>Information Security</b></td></tr>"
$message = $message + "</table>"

$report = $format + $title + $table + $footer

#-- Generate HTML file
$report | Out-File $html -Encoding Utf8

#-- Export to CSV
$result | Sort Company | Export-Csv $csv -NoTypeInformation -Encoding Utf8

#-- Send report by email
$Subject = "[ Report-$mounth ] Active Directory - Member Groups"
$SmtpServer	= $config[11]
$Port = $config[13]
$From = $config[15]
$To = $config[17]

Send-MailMessage -From $From -To $To -Subject $Subject -Attachments $html,$csv -bodyashtml -Body $message -SmtpServer $SmtpServer -Port $Port

cls