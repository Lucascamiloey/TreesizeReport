# Set up file paths
#File containing all paths to report
$data=get-content D:\Scripts\Treesize\PathsToScan.txt
#Path and filename to contain the reports
$workingpath="D:\Scripts\Treesize\results"
$filename="Treesize Report - $(get-date -Format 'ddMMyy').xlsx"
#Setting treesize intallation path
$command="C:\Program Files\JAM Software\TreeSize Professional\Treesize.exe" 

#Checks for previos existing report and renames it
if (test-path $workingpath\$filename){
	Move-Item -Path $workingpath\$filename -Destination $workingpath\$filename-$(get-date -format 'hhmm').old
}
 
foreach ($path in $data){
	#Transforms \\ip\share$\folder to ip-share-folder
	$name=((($path.replace("\","-")).substring(2)).trimend("-")).replace("$","")
	# uncomment for debugging
	#write-host "Scanning $path Using sheetname $name"
	&"$command" /EXCEL $workingpath\$filename /NOGUI /SHEETNAME $name /EXPAND 50MB /HIDESMALLFOLDERS 50MB /SIZEUNIT 2 $path | out-null
}

#Set up mail variables 
$date=get-date -Format 'dd/MM/yyyy'
$toAddress="TO@EMAIL.COM"
$ccAddress="CC@EMAIL.COM"
$fromAddress="FROM@EMAIL.COM"
$subject="$date Treesize Report"
$smtp="SMTP.DOMAIN.COM"
$attachment="$workingpath\$filename"
$body="Treesize report for $date <br> This is an automated email message, please <b>do not reply.</b>"

#Send the email
Send-MailMessage -to $toAddress -cc $ccAddress -from $fromAddress -subject $subject -smtpserver $smtp -body $body -BodyAsHtml -attachment $attachment 


