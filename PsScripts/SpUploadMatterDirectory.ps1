#
# SpUploadMatterDirectory.ps1
#

Param(
	[Parameter(Mandatory=$True)]
	[string]$directory,
	[Parameter(Mandatory=$true)]
	[string]$ClientFolder,
	[Parameter(Mandatory=$true)]
	[string]$matterId,
	[Parameter(Mandatory=$true)]
	[string]$SpSite,
	[Parameter(Mandatory=$true)]
	[string]$SpLibrary,
	[string]$LogFile,
	[string]$ErrorLogFile,
	[Parameter(ParameterSetName='Zip')]
	[switch]$zip,
	[Parameter(Mandatory=$true,
				ParameterSetName='Zip')]
	[string]$ZipBase
)

$IncludeExtensions = '*.doc', '*.docx', '*.docm', '*.xls', '*.xlsx', '*.dot', '*.dotx', '*.dotm', '*.pdf', '*.tif', '*.htm', '*.wpd', '*.rtf', '*.txt', '*.msg'

Connect-PnPOnline $SpSite

$matter = Get-PnPListItem -List Matters -query "<View><Query><Where><Eq><FieldRef Name='MatterID' /><Value Type='Text'>$matterId</Value></Eq></Where></Query></View>"
if ($matter.count -ne 1) { Throw }	


Function Remove-InvalidFileNameChars($name)
{
	$invalidChars = [IO.Path]::GetInvalidFileNameChars()
	$invalidChars += '#'
	$replace = [RegEx]::Escape(-join $invalidChars)
	$name -replace '[#&()]'
	
}

function SPUploadFile($filename, $relative_path_regex)
{

		$relative_folder = [RegEx]::matches($filename, $relative_path_regex)
		$relative_folder = Split-Path -Path $relative_folder.groups[1].value
		$sp_folder = $SpLibrary, $ClientFolder, $relative_folder -join "/"
		$sp_folder = $sp_folder.Replace('\', '/')
		$sp_folder = $sp_folder.Replace('//', '/')
		$leafFilename = Split-Path -Path $filename -leaf 
		$sanitized_filename = Remove-InvalidFileNameChars $leafFilename

		Write-Host "PSPath: " -nonewline; write-host $filename

		$file_details = @{
			#Source=$file_path;
			FS_LastWrite=$child.lastwritetimeutc;
			FS_Created=$child.creationtimeutc;
			FS_LastAccess=$child.lastaccesstimeutc
			Matter=$matter.id
		}

		
		try
		{
			
			$fsFile = New-Object System.IO.FileStream ($filename, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::Read);
			$added_file = Add-PnPFile -Folder $sp_folder -Stream $fsFile -FileName $sanitized_filename -Values $file_details -Checkout -CheckInComment "Intial Upload" -ErrorAction Stop
			$log_entry_properties = @{
				Date=Get-Date;
				MatterID=$matterId;
				File=$filename;
				SP_File=$added_file.ServerRelativeUrl

			}
			if ($LogFile)
				{
					$log_entry = New-Object -TypeName psobject -Property $log_entry_properties
					$log_entry | Export-Csv -Path $LogFile -NoTypeInformation -Force -Append
				}
			Write-Host "added file: " $added_file.ServerRelativeUrl
		}
		catch
		{
			Write-Host $_.Exception
			$error_entry = @{
				Date=Get-Date;
				Message=$_.Exception.Message;
				ParamName=$_.Exception.ParamName;
				Filename=$child.FullName;
				InvocationInfo=$_.InvocationInfo
			}
			if ($ErrorLogFIle)
			{
				$error_log = New-Object -TypeName psobject -Property $error_entry
				$error_log | Export-Csv $ErrorLogFile -NoTypeInformation -Force -Append
			}
		}



	try 
		{
			$sp_file = Find-PnPFile -Folder $sp_folder -Match $sanitized_filename -ErrorAction stop	
		}

	Catch

	{
		Write-Host $_.Exception
			$error_entry = @{
				Date=Get-Date;
				Filename=$child.FullName;
				InvocationInfo=$_.Exception.InvocationInfo;
				ErrorDetails=$_.Exception.ErrorDetails;
				Message=$_.Exception.PSMessageDetails
			}
			$error_log = New-Object -TypeName psobject -Property $error_entry
			$error_log | Export-Csv $ErrorLogFile -NoTypeInformation -Force -Append
	}
		

}


foreach($file in Get-ChildItem -File -Recurse -Path $directory -Include $IncludeExtensions) 
{
	SPUploadFile $file.fullname (($directory | Split-Path -leaf), "(.*)" -join "")
}

$SpClientFolder = $SpLibrary, $ClientFolder -join "/"
$FolderQuery = "<View Scope='RecursiveAll'><Query><Where><And><Eq><FieldRef Name='ContentType' /><Value Type='Computed'>Client Folder</Value></Eq><BeginsWith><FieldRef Name='FileRef' /><Value Type='Text'>", $SpClientFolder, "</Value></BeginsWith></And></Where></Query></View>" -join ""
$folders = Get-PnPListItem -list $SpLibrary -Query $FolderQuery
foreach($folder in $folders) 
{
	$item = Set-PnPListItem -Identity $folder -List $SpLibrary -Values @{ Matter=$matter.id }
	$item
}

if ($zip)
{
	#try {
		#$zip_file = Split-Path $directory.BaseName -Leaf
		$destination_path = $ZipBase,  $ClientFolder -join "\"
		Compress-Archive -Path $directory -DestinationPath $destination_path -ErrorAction stop
	#}
	#catch {
		#$_.Exception
	#} 
}