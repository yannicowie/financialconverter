################################CSV to QIF Converter################################

#Version 1.0

#Powershell script for creating QIF files from CSV data. The Script requires a file
#named "data.csv" exists in the working directory (the same folder as the script).
#This CSV file should have five columns with the following headers in row 1:
#"Date","Amount","Type","Details","Particulars". The "Date" column must contain data
#in the format "02/01/19" or "02/01/2019" (i.e. the second of janurary 2019).
#If the CSV data is not in this format the script will not work.

##Created by Yanni Cowie

#####################################################################################

#Creating File
$filename = read-host "Enter a name for the QIF file (without any file extension)"

$filenamext = $filename + ".txt"
$filenameofx = $filename + ".qif"

New-Item -Path . -Name $filenamext -ItemType "file"

#Creating QIF header. This has hardcoded in the type as a bank account. If another account type needs to be used it can be updated here
Add-Content $filenamext {!Type:Bank
^}

#Importing CSV data
[string]$MyDirectory = Get-Location

$csvpath = $MyDirectory + "\data.csv"

[array]$csv = Import-Csv $csvpath

$global:LinesInFile = $csv.count

#Data manipulation function
function Line-Writer {
	
	#Getting info for Line Items
	
	$trndate = $csv[$lineit].Date
	$trnamount = $csv[$lineit].Amount
	$trntype = $csv[$lineit].Type
	$trndetails = $csv[$lineit].Details
	$trnparticulars = $csv[$lineit].Particulars

	#Creating lines
	$line1 = "D" + $trndate
	$line2 = "T" + $trnamount
	$line3 = "L" + $trntype
	$line4 = "P" + $trndetails
	$line5 = "M" + $trnparticulars
	$line6 = "^"

	#Writing the lines to the file
	$line1 | Out-File $filenamext -Append -Encoding ASCII
	$line2 | Out-File $filenamext -Append -Encoding ASCII
	$line3 | Out-File $filenamext -Append -Encoding ASCII
	$line4 | Out-File $filenamext -Append -Encoding ASCII
	$line5 | Out-File $filenamext -Append -Encoding ASCII
	$line6 | Out-File $filenamext -Append -Encoding ASCII
	
	$lineit = $lineit + 1
	
	#Identifying whether there are more lines in file to convert or whether the conversion is complete
	if($lineit -lt $LinesInFile) {
	Line-Writer
	}
	else {
	write-host "Conversion is done"
		#renaming the .txt file to .qif
		Rename-Item -Path $filenamext -NewName $filenameofx
	pause
	}
}

#Setting the line iteration counter to zero and triggering the conversion function
[int]$global:Lineit = 0
Line-Writer