################################CSV to OFX Converter################################

#Version 1.0

#Powershell script for creating OFX files from CSV data. The Script requires a file
#named "data.csv" exists in the working directory (the same folder as the script).
#This CSV file should have five columns with the following headers in row 1:
#"TransactionType","Date","Amount","Name","Memo". The "Date" column must contain data
#in the format "20190102" (i.e. the second of janurary 2019), not a traditional date
#such as "02/01/2018" If the CSV data is not in this format the script will not work.

##Created by Yanni Cowie

#####################################################################################

#Creating File
$filename = read-host "Enter a name for the OFX file (without any file extension)"

	#getting balance info
	write-host "Now you will be asked questions about the account balance data to include in the file" -ForegroundColor Red
	$baldate = read-host "Enter the balance date in the format 20190102 (eg. 2nd Janurary 2019)"
	$ledgerbal = read-host "Enter the Account Balance at that date"
	$availbal = read-host "Enter the Available Balance at that date"

$filenamext = $filename + ".txt"
$filenameofx = $filename + ".ofx"

New-Item -Path . -Name $filenamext -ItemType "file"

#Creating OFX header. This has some hardcoded parameters in it which can be updated below if needed.
Add-Content $filenamext {OFXHEADER:100
DATA:OFXSGML
VERSION:102
SECURITY:NONE
ENCODING:USASCII
CHARSET:1252
COMPRESSION:NONE
OLDFILEUID:NONE
NEWFILEUID:NONE

<OFX>
<SIGNONMSGSRSV1>
<SONRS>
<STATUS>
<CODE>0
<SEVERITY>INFO
</STATUS>
<DTSERVER>20190131043320
<LANGUAGE>ENG
</SONRS>
</SIGNONMSGSRSV1>
<BANKMSGSRSV1>
<STMTTRNRS>
<TRNUID>1772506164
<STATUS>
<CODE>0
<SEVERITY>INFO
</STATUS>
<STMTRS>
<CURDEF>NZD
<BANKACCTFROM>
<BANKID>06
<BRANCHID>1234
<ACCTID>1234567-00
<ACCTTYPE>CHECKING
</BANKACCTFROM>
<BANKTRANLIST>
<DTSTART>20190102
<DTEND>20190130}

write-host "Header Written"

#Importing CSV data
[string]$MyDirectory = Get-Location

$csvpath = $MyDirectory + "\data.csv"

[array]$csv = Import-Csv $csvpath

$global:LinesInFile = $csv.count

#Data manipulation function
function Line-Writer {
	
	#Getting info for Line Items

	$trntype = $csv[$lineit].TransactionType
	$date = $csv[$lineit].Date
	$amount = $csv[$lineit].Amount
	$name = $csv[$lineit].Name
	$memo = $csv[$lineit].Memo

	#Creating lines
	$line1 = "<STMTTRN>"
	$line2 = "<TRNTYPE>" + $trntype
	$line3 = "<DTPOSTED>" + $date
	$line4 = "<TRNAMT>" + $amount
	$line5 = "<FITID>201901290"
	$line6 = "<NAME>" + $name
	$line7 = "<MEMO>" + $memo
	$line8 = "</STMTTRN>"

	#Writing the lines to the file
	$line1 | Out-File $filenamext -Append -Encoding ASCII
	$line2 | Out-File $filenamext -Append -Encoding ASCII
	$line3 | Out-File $filenamext -Append -Encoding ASCII
	$line4 | Out-File $filenamext -Append -Encoding ASCII
	$line5 | Out-File $filenamext -Append -Encoding ASCII
	$line6 | Out-File $filenamext -Append -Encoding ASCII
	$line7 | Out-File $filenamext -Append -Encoding ASCII
	$line8 | Out-File $filenamext -Append -Encoding ASCII
	
	$lineit = $lineit + 1
	
	#Identifying whether there are more lines in file to convert or whether the conversion is complete
	if($lineit -lt $LinesInFile) {
	Line-Writer
	}
	else {
	write-host "Lines Written"
	}
}

#Setting the line iteration counter to zero and triggering the conversion function
[int]$global:Lineit = 0
Line-Writer

#Writing footer
	#Creating lines
	$line1 = "</BANKTRANLIST>"
	$line2 = "<LEDGERBAL>"
	$line3 = "<BALAMT>" + $ledgerbal
	$line4 = "<DTASOF>" + $baldate
	$line5 = "</LEDGERBAL>"
	$line6 = "<AVAILBAL>"
	$line7 = "<BALAMT>" + $availbal
	$line8 = $line4

	#Writing the lines to the file
	$line1 | Out-File $filenamext -Append -Encoding ASCII
	$line2 | Out-File $filenamext -Append -Encoding ASCII
	$line3 | Out-File $filenamext -Append -Encoding ASCII
	$line4 | Out-File $filenamext -Append -Encoding ASCII
	$line5 | Out-File $filenamext -Append -Encoding ASCII
	$line6 | Out-File $filenamext -Append -Encoding ASCII
	$line7 | Out-File $filenamext -Append -Encoding ASCII
	$line8 | Out-File $filenamext -Append -Encoding ASCII
	
	#finishing off the closing tags in bulk
	Add-Content $filenamext {</AVAILBAL>
</STMTRS>
</STMTTRNRS>
</BANKMSGSRSV1>
</OFX>}

#renaming the .txt file to .ofx
Rename-Item -Path $filenamext -NewName $filenameofx
write-host "Footer written"
write-host "Conversion Completed"