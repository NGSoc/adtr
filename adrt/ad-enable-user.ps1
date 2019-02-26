<# 
.NAME
    ADRT - Active Directory Report Tool
.DESCRIPTION
    Extract the complete list of All Computers.
.EXAMPLE
    PS C:\adrt> .\ad-enable-user.ps1
.NOTES
    Name: Thaigo Freitas e Alessandro Camargo
	E-mail: brntt0012@ambev.com.br
    E-mail: brntt0002@ambev.com.br
.LINK
    
#>

$report = $null
$table = $null
$date = Get-Date -format "yyyy-M-d"
$mounth = Get-Date -format "MMM"
$directorypath = (Get-Item -Path ".\").FullName
$path = "ad-reports\ad-enable-user"
$html = "$path\ad-enable-user-$date.html"
$csv = "$path\ad-enable-user-$date.csv"

#-- Show Total
$table += "<center><h3><b>Total Users: <font color=red>$t_u</font> - PasswordNeverExpires: <font color=red>$t_pne</font> - Disabled Users: <font color=red>$t_du</font></b></h3></center>"

#-- Import Module
Import-Module ActiveDirectory

GC C:\adrt\usuario.txt | %{ set-ADAccountExpiration -Identity $_ -DateTime '30/07/2019 00:00:00'}

$result = @(GC C:\adrt\usuario.txt | %{ get-aduser -identity $_ -properties sAMAccountName,cn,accountExpires,whenCreated,Enabled,lastLogon} | select sAMAccountName,cn,accountExpires,whenCreated,Enabled,lastLogon)

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
		<font face='Calibri' size='5px'><b>Active Directory</b></font>
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
$report = $format + $title + $table + $footer

#-- Generate HTML file
$report | Out-File $html -Encoding Utf8

#-- Export to CSV
$result | Sort Company | Export-Csv $csv -NoTypeInformation -Encoding Utf8r




