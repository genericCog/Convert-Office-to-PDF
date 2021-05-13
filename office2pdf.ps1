<#
.SYNOPSIS
	Convert Microsoft Word, Excel, and PowerPoint files to PDF, Uses recursion to convert files in subfolders
	Open a PowerShell (PS) window, Change Directory to the location of this script
	cd 'C:\Users\DC33188\OneDrive - Defense Information Systems Agency\Documents\My_Library'
.DESCRIPTION
	Open a PowerShell (PS) window, Change Directory to the location of this script and run: .\office2pdf.ps1
	This PowerShell (PS) script begins by asking the user for the types of files they want to convert (a, p, w, x)
	After the user enters the file type, they are prompted to enter the folder path
	The conversion process begins and each file is listed in the PS window
	When the conversion process is complete, the PS window displays a count of files converted and the location of the log file
.PARAMETER
	To do in the future
.INPUTS
    user is prompted for file type, directory, and log file
.OUTPUTS
    Output is on console and log file. Creates a new PDF.
.NOTES
  Version:			1.0
  Author:			Adam Cherochak
  Creation Date:	22 APRIL 2021
  Purpose: 			Convert Microsoft Word documents to Adobe PDF
  Useful URLs:		https://stackoverflow.com/questions/16534292/basic-powershell-batch-convert-word-docx-to-pdf
					https://devblogs.microsoft.com/scripting/weekend-scripter-convert-word-documents-to-pdf-files-with-powershell/
					https://devblogs.microsoft.com/scripting/save-a-microsoft-excel-workbook-as-a-pdf-file-by-using-powershell/
.EXAMPLE 1
    Open PowerShell (PS) and change directory to the location of this script. 
	Use the "cd" command: cd C:\Users\nDocuments
	Run this script in the PS window using the command:   .\doc2pdf.ps1
	NOTE: the ".\" is the current directory
#>
# C:\Users\DC33188\Documents\a_Test

$global:userInputPath = $null
$global:userInputFileType = $null
$global:setfiletypeFlag = 0
$global:logFileName = "fileList.log"
$global:logFilePath = $null
$global:dirName = $null
$global:filename = $null
$global:ext = $null
$global:dtgLogFilePath = $null

Function Set_Event_Log($logFilePath){
	if($logFilePath){
		$global:dirName  = [io.path]::GetDirectoryName($logFilePath)
		$global:filename = [io.path]::GetFileNameWithoutExtension($logFilePath)
		$global:ext = [io.path]::GetExtension($logFilePath)
		$global:dtgLogFilePath  = "$dirName\$(get-date -f yyyy-MM-dd)_$filename$ext"
	}else{
		#Write-Output "WARNING! Exiting program. Incorrect Log File Path, dtgLogFilePath is:" $dtgLogFilePath
		exit
	}#end if else logfilepath
}#end Set_Event_Log

Function Convert_PowerPoint($userInputPath){
	If ($userInputPath) {
		Try {			
			$logFilePath = "$userInputPath\$logFileName"
			Set_Event_Log $logFilePath
			Write-Output " "
			Write-Output "######################################################################"
			Write-Output "#####                                                            #####"
			Write-Output "#####                Convert PowerPoint Files                    #####"
			Write-Output "#####                                                            #####"
			Write-Output "######################################################################"
			Write-Output " "
			Write-Output " "
#_____
			$objPowerPoint = New-Object -ComObject PowerPoint.Application # Create a PowerPoint object
			$counter = 0
			$input = Read-Host "Create NEW log file? This will delete the exisiting log file. Please enter 'y' or 'n'."
			if (($input -isnot [string])) { Throw 'You did not provide y or n as input' }
			switch ($input) {
				y{
					if (Test-Path $dtgLogFilePath){
						Remove-Item $dtgLogFilePath
					}else{continue}
				}
				n{continue}
			}#end switch input
			
			# Get all objects of type .pptx in $userInputPath and its subfolders
			Get-ChildItem -Path $userInputPath -Recurse -Filter *.ppt? | ForEach-Object {
			Start-Sleep -m 50
				Write-Host "Converting: " $_.Name #$_.FullName
				$document = $objPowerPoint.Presentations.Open($_.FullName, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse)
				# Define the file path and document name. To save all files in the same folder, use $userInputPath # (EX: # $pdf_filename = "$($userInputPath)\$($_.BaseName).pdf")
				$pdf_filename = "$($_.DirectoryName)\$($_.BaseName).pdf"
			#Write-Host "SAVING DOCUMENT: " $pdf_filename
				$pdf_filename | Out-File -File $dtgLogFilePath -Append
				$opt= [Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType]::ppSaveAsPDF # Save as PDF -- 17 is the literal value of 'wdFormatPDF'
				$document.SaveAs($pdf_filename, $opt)			
				$document.Close()
				$counter = $counter + 1		
			}#end foreach pptx
			Start-Sleep -Seconds 3 
			$objPowerPoint.Quit()		
			[System.Runtime.Interopservices.Marshal]::ReleaseComObject($objPowerPoint)

#_____
			
			Write-Output "     "
			Write-Output "------------------------------------------------------------"
			Write-Output "----- Finished. Number of pptx converted: " $counter 
			Write-Output "----- Log file: "  $dtgLogFilePath
			Write-Output "------------------------------------------------------------"
			Write-Output "     "				
		}
		Catch {
			Write-Output "WARNING! Statement failed in: Convert_PowerPoint"
			Write-Output $error[0]
			Break
		}
	}#end IF 
}#end Function Convert_PowerPoint
Function Convert_Word($userInputPath){
	If ($userInputPath) {
		Try {
			$logFilePath = "$userInputPath\$logFileName"
			Set_Event_Log $logFilePath
			Write-Output " "
			Write-Output "######################################################################"
			Write-Output "#####                                                            #####"
			Write-Output "#####                   Convert Word Files                       #####"
			Write-Output "#####                                                            #####"
			Write-Output "######################################################################"
			Write-Output " "
			Write-Output " "
#_____

			$input = Read-Host "Create NEW log file? This will delete the exisiting log file. Please enter 'y' or 'n'."
			if (($input -isnot [string])) { Throw 'You did not provide y or n as input' }
			switch ($input) {
				y{
					if (Test-Path $dtgLogFilePath){
						Remove-Item $dtgLogFilePath
					}else{continue}
				}
				n{continue}
			}#end switch input	

$counter = 0
Get-ChildItem -Path $userInputPath -Filter *.doc? -Recurse | ForEach-Object { 
	Start-Sleep -m 50
	$objWord = New-Object -ComObject Word.Application
	$objWord.visible = $false
	$document = $objWord.Documents.Open($_.FullName)
	$pdf_filename = "$($_.DirectoryName)\$($_.BaseName).pdf"
	Write-Output "Converting:  $($document.Name)"
	$document.SaveAs([ref] $pdf_filename, [ref] 17)
	$pdf_filename | Out-File -File $dtgLogFilePath -Append
	$document.Close()
	$objWord.Quit()
	$counter++
	Stop-Process -Name winword
}

Start-Sleep -Seconds 3 
#Invoke-Item $userInputPath
	
	
#end getchilditem word
#_____	
			Write-Output "     "
			Write-Output "------------------------------------------------------------"
			Write-Output "----- Finished. Number of docx converted: " $counter 
			Write-Output "----- Log file: "  $dtgLogFilePath
			Write-Output "------------------------------------------------------------"
			Write-Output "     "				
		}
		Catch {
			Write-Output "WARNING! Statement failed in: Convert_Word"
			Write-Output $error[0]
			Break
		}
	}#end IF
}#end Function Convert_Word
Function Convert_Excel($userInputPath){
	If ($userInputPath) {
		Try {
			$logFilePath = "$userInputPath\$logFileName"
			Set_Event_Log $logFilePath
			Write-Output " "
			Write-Output "######################################################################"
			Write-Output "#####                                                            #####"
			Write-Output "#####                  Convert Excel Files                       #####"
			Write-Output "#####                                                            #####"
			Write-Output "######################################################################"
			Write-Output " "
			Write-Output " "

			$excelFilter = Get-ChildItem -Path $userInputPath -include *.xls, *.xlsx -Recurse 
			$xlFixedFormat = "Microsoft.Office.Interop.Excel.xlFixedFormatType" -as [type]
			$objExcel = New-Object -ComObject Excel.Application
			$objExcel.visible = $false
			$counter = 0

			$input = Read-Host "Create NEW log file? This will delete the exisiting log file. Please enter 'y' or 'n'."
			if (($input -isnot [string])) { Throw 'You did not provide y or n as input' }
			switch ($input) {
				y{
					if (Test-Path $dtgLogFilePath){
						Remove-Item $dtgLogFilePath
					}else{continue}
				}
				n{continue}
			}#end switch input

			foreach($wb in $excelFilter){
			Start-Sleep -m 50
				$filepath ="$(($wb.FullName).substring(0, $wb.FullName.lastIndexOf("."))).pdf"
				if ((Test-Path $filepath) -And (Get-Item $filepath).length -gt 3kb) {
					$directoryTestStatus = 0
					echo "WARNING: File already exists, skiping operation for:  $($filepath)"
					#echo "If error persists, contact Adam Cherochak: adam.cherochak@gmail.com"
					# continue
				}else{
					$directoryTestStatus = 1 #echo "Directory Test Passed: $($filepath)"
				}#end IF Test-Path
				if ($directoryTestStatus -eq 1){
					$workbook = $objExcel.workbooks.open($wb.fullname, 3)
					echo "Converting:  $($wb.Name)"
					$workbook.Saved = $true
					$workbook.ExportAsFixedFormat($xlFixedFormat::xlTypePDF, $filepath) 
					$objExcel.Workbooks.close()
					$objExcel.Quit()
					while([System.Runtime.Interopservices.Marshal]::ReleaseComObject($objExcel)){'XL released'}
					$objExcel = New-Object -ComObject Excel.Application
					#Write-Host "Restarting Excel ComObject, please wait."
					$filepath | Out-File -File $dtgLogFilePath -Append

					if ($counter -gt 100) {
						$counter = 0
						$objExcel.Workbooks.close()
						$objExcel.Quit()
						while([System.Runtime.Interopservices.Marshal]::ReleaseComObject($objExcel)){'XL released'}
						$objExcel = New-Object -ComObject Excel.Application
						Write-Host "Restarting Excel ComObject, please wait."			
					}#end IF counter
					$counter = $counter + 1
					 
					continue
				}#end IF directoryTestStatus
			}#end foreach
			Start-Sleep -Seconds 3 
			Write-Output "     "
			Write-Output "------------------------------------------------------------"
			Write-Output "----- Finished. Number of xlsx converted: " $counter 
			Write-Output "----- Log file: "  $dtgLogFilePath
			Write-Output "------------------------------------------------------------"
			Write-Output "     "
		}#end try
		Catch {
			Write-Output "WARNING! Statement failed in: Convert_Excel"
			Write-Output $error[0]
			$error | Out-File -File $dtgLogFilePath -Append 
			Break
		}#end try-catch
	}#end IF userInputPath
}#end Function Convert_Excel

Function Set_Directory{
	#'C:\Users\DC33188\OneDrive - Defense Information Systems Agency\Documents\S3 Ops\a_Test'
	Write-Output " "
	Write-Output "######################################################################"
	Write-Output "#####                     Folder Location                        #####"
	Write-Output "#####   Copy & paste the folder path of the files to convert     #####"
	Write-Output "#####   Example:                                                 #####"
	Write-Output "#####         C:\Users\DC33188\OneDrive - DISA\S3                #####"
	Write-Output "######################################################################"
	Write-Output " "
	Write-Output " "
	$userPath = Read-Host "Enter the folder path: "
		if(!$userPath -or -not(Test-Path -Path $userPath)){
			Write-Output " "	
			Write-Output "WARNING! The location is not valid " $userPath
			Write-Output "Please try again..."
			exit # Set_Directory
		}
		elseif(Test-Path -Path $userPath){
			$global:userInputPath = $userPath
		}#end if/else userPath	
}#end Set_Directory

Function Set_File_Type{
	
	Clear-Host
	
	Write-Output "     "
	Write-Output "######################################################################"
	Write-Output "#####                        File Type                           #####"
	Write-Output "#####      Enter a lowercase letter to convert file type         #####"
	Write-Output "#####                                                            #####"
	Write-Output "#####    Enter 'a' for all. This includes: pptx, xlsx, & docx    #####"
	Write-Output "#####                                                            #####"
	Write-Output "#####   For only one file type, enter the corresponding letter:  #####"
	Write-Output "#####         'p' for PowerPoint                                 #####"
	Write-Output "#####         'w' for Word                                       #####"
	Write-Output "#####         'x' for Excel                                      #####"
	Write-Output "######################################################################"
	Write-Output "     "
	Write-Output "     "
	$userInputFileType = Read-Host "Enter 'a' for all file types OR 'p', 'w', 'x'"
	if ( ($userInputFileType -isnot [string]) ) { 
		Write-Output " "
		Throw "Please use 'a', 'p', 'w', or 'x'" #Write-Warning "Please use 'a', 'p', 'w', or 'x'"
		$global:setfiletypeFlag = 0
		Write-Output " "
		Set_File_Type
	}
	else{
		switch ($userInputFileType) {
			a{
				Write-Output "you selected a" 
				$global:setfiletypeFlag = 1	
				
				Set_Directory #get-set folder location
				
				Write-Output "your folder location is: " $userInputPath
				
				Convert_PowerPoint $userInputPath
				Convert_Word $userInputPath
				Convert_Excel $userInputPath
			}
			p{
				Write-Output "you selected p"
				$global:setfiletypeFlag = 1
				Set_Directory #get-set folder location
				Write-Output "your folder location is: " $userInputPath
				Convert_PowerPoint $userInputPath
			}
			w{
				Write-Output "you selected w"
				$global:setfiletypeFlag = 1
				Set_Directory #get-set folder location
				Write-Output "your folder location is: " $userInputPath
				Convert_Word $userInputPath
			}
			x{
				Write-Output "you selected x"
				Set_Directory #get-set folder location
				Write-Output "your folder location is: " $userInputPath
				$global:setfiletypeFlag = 1
				Convert_Excel $userInputPath
			}
			default {
				Write-Warning "Please use 'a', 'p', 'w', or 'x'"
				$global:setfiletypeFlag = 0
				exit
			}
		}#end switch userInputFileType
	}#end if userInputFileType
}#end Set_File_Type

Function Delete_PDF_Files( $userInputPath ){

		$input = Read-Host "Delete PDF files? y/n"
		if (($input -isnot [string])) { Throw 'You did not provide y or n as input' }
		switch ($input) {
			y{
				if (Test-Path $userInputPath){										
					Get-Content $dtgLogFilePath | ForEach-Object {
						Write-Output "Removing File: " $_
						Remove-Item $_
					}#end if userInputPath
				}else{continue}#end if else
			}#end yes switch
			n{continue}#end no switch
		}#end switch
}#end Delete_PDF_Files

Set_File_Type #call this first to set file paths
Delete_PDF_Files $userInputPath #call function: Delete_PDF_Files
