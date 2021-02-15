############################################################################################
#                                                                                          #
# The sample scripts are not supported under any Microsoft standard support                #
# program or service. The sample scripts are provided AS IS without warranty               #
# of any kind. Microsoft further disclaims all implied warranties including, without       #
# limitation, any implied warranties of merchantability or of fitness for a particular     #
# purpose. The entire risk arising out of the use or performance of the sample scripts     #
# and documentation remains with you. In no event shall Microsoft, its authors, or         #
# anyone else involved in the creation, production, or delivery of the scripts be liable   #
# for any damages whatsoever (including, without limitation, damages for loss of business  #
# profits, business interruption, loss of business information, or other pecuniary loss)   #
# arising out of the use of or inability to use the sample scripts or documentation,       #
# even if Microsoft has been advised of the possibility of such damages.                   #
#                                                                                          #
############################################################################################      
#                                                                                          #
# Created by Santosh.Dakolia@microsoft.com                                                 #
#                                                                                          #
############################################################################################

   <#
        .SYNOPSIS
        This is a Simple Script to Split any CSV file with three different Options
        	1) Split CSV with Number of Record Set.
            2) Split CSV with Number of Batches.
            3) Split CSV with Alphabets.
                    
        .DESCRIPTION
        Once you run the Script it will open "File Open Dialog Box " once you select the file it will give you below Options.
        	1) Split with Number of Record Set.
            2) Split in Batches.
            3) Split with Alphabets.
        
        1) Split with Number of Record Set.
            # As the Name indicates this option will be use to split the Selected File with number of Record Set available.
            
        2) Split in Batches
            # This option will split the CSV into Number of Batches as mention by the User

        3) Split with Alphabets (It is NOT Case Sensitive.
            # This Option will Split the CSV with Single or Range of Alphabets. 
            For Example : 
                        * A
                        * A-C
                        * E-Z
                        * A-Z

        .INPUTS
        None.

        .OUTPUTS
        None.

        .EXAMPLE
        1) Split with Number of Record Set.
        If Total Records in given CSV is 4647 then..

        Please Enter the Number of RecordSet to Split the CSV : 600
        Number of CSV Files created will be : 8
        Current Folder Path : "It will show you the current Primary CSV Path"

        .EXAMPLE
        2) Split with Number of Record Set.
        If Total Records in given CSV is 4647 then..

        Please Enter the Number of Batches to Split the CSV : 4
        Number of CSV Files created will be : 4
        Current Folder Path : "It will show you the current Primary CSV Path"

        .EXAMPLE
        3) Split with Alphabets.
        The File can be split with Single or Range of Alphabets. 
        Sample :-
                * A
                * A-C
                * E-Z
                * A-Z

        .LINK
        None

        .NOTES
        Author : Santosh Dakolia (Santosh.Dakolia@Microsoft.com)
        Last Edit : 15 Feb 2021
        Version 1.0 - Initial Release

    #>
Function Get-FileName {
  [System.Reflection.Assembly]::LoadWithPartialName(“System.windows.forms”) |  Out-Null
  $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
  $OpenFileDialog.filter = “CSV Files (*.CSV)| *.CSV”
  $OpenFileDialog.ShowDialog() | Out-Null
  $OpenFileDialog.filename
  $Global:FP = $OpenFileDialog.filename
}
Function Extract_Folderpath {
            $MCSVE = $MainCSV.Split('\')[-1]
            $MCSV = $MainCSV.Split('\')
            $MCSV = $MCSV | Where-Object { $PSItem -ne $MCSVE }
            $EFPath = $MCSV -join '\'
            return $EFPath
        }
Function Show-Menu {
            param (
                [string]$Title = 'Split CSV Function'
            )
            Clear-Host
            Write-Host ""
            Write-Host "######################## $Title ########################" -ForegroundColor DarkGray
            Write-Host ""
            Write-Host "CSV Total Recordset :: $CSVC" -ForegroundColor DarkYellow
            Write-Host "Current CSV Export Folder Path :: $PPath"  
            Write-Host ""
            Write-Host ""
            Write-Host "Please Press Number to select the Options"
            Write-Host "=========================================" -ForegroundColor DarkCyan
            Write-Host "1) "-ForegroundColor Green -NoNewline; Write-Host "Split CSV with Number of Record Set"
            Write-Host "    # Enter the Number of Record Items." -ForegroundColor DarkCyan
            Write-Host ""
            Write-Host "2) " -ForegroundColor Green -NoNewline; Write-Host "Split CSV with Number of Batches"
            Write-Host "    # Enter the Number of Batches." -ForegroundColor DarkCyan
            Write-Host ""
            Write-Host "3) " -ForegroundColor Green -NoNewline; Write-Host "Split CSV with Alphabets"
            Write-Host "    # Enter the Alphabets to Split the CSV" -ForegroundColor DarkCyan
            Write-Host ""
            Write-Host "4) " -ForegroundColor Green -NoNewline; Write-Host "Enter Folder path to Export CSV"
            Write-Host ""
            Write-Host "Press 'Q' to Quit." -ForegroundColor Red
            Write-Host ""
        }

Get-FileName
Clear-Host
$Fpath = $null
$MainCSV = $Global:FP

If (!($MainCSV)) { Write-Host "Main CSV File Path is Empty" }
If ($MainCSV) {
    If (Test-Path $MainCSV) {
        $ICSV = Import-Csv $MainCSV
        $CSVC = $ICSV.Count
        $PPath = Extract_Folderpath
        Write-Host "Current Folder Path :: " $Fpath
        do {
            Show-Menu
            $input = Read-Host "Please make a selection"
            switch ($input) {
                '1' {
                    #cls
                    Try {
                        [int]$RecordSet = Read-Host "Please Enter the Number of RecordSet to Split the CSV" 
                
                        If ($RecordSet -eq 0 -or $null -eq $RecordSet) {
                            Write-Host "Please enter Number Greater then Zero"
                            Exit
                        }
                        If ($RecordSet -gt $CSVC) {
                            Write-Host "Total Number of Split Recordset Exceeded the Total CSV RecordSet"
                            Exit
                        }
                    }
                    catch {
                        Write-host "Error : Please Enter Only Numbers" -ForegroundColor Red
                        exit
                    }

                    IF ($RecordSet) {
                        #$SplitNum = $num
                        $CSVArray = @()
                        $Batch = [Math]::Ceiling($ICSV.Count / $RecordSet)
                        For ($i = 0; $I -lt $Batch; $i++) {
                            $SR = ($i * $RecordSet)
                            $ER = (($i + 1) * $RecordSet) - 1
                            If ($ER -ge $ICSV.Count) { $ER = $ICSV.Count }
                            $CSVArray += , @($ICSV[$SR..$ER])
                        }
                    }
                    Write-Host "Number of CSV Files created will be : " $CSVArray.count
                    If (!($Fpath)) { $FPath = $PPath }

                    Write-Host "Current Folder Path :: " $Fpath

                    0..(($CSVArray.count) - 1) | ForEach-Object {
                        $CD = $_ + 1
                        $CSVArray[$_] | Export-Csv -Path $FPath\$($CD)_DATARecordset.csv -NoTypeInformation -Force }
                    exit
                } 
                '2' {
                    Try {
                        [int]$Batch = Read-Host "Please Enter the Number of Batches to Split the CSV" 
                
                        If ($Batch -eq 0 -or $null -eq $Batch) {
                            Write-Host "Please enter Number Greater then Zero"
                            Exit
                        }
                        If ($Batch -gt $CSVC) {
                            Write-Host "Total Number of Batch Exceeded the Total CSV RecordSet"
                            Exit
                        }
                    }
                    catch {
                        Write-host "Error : Please Enter Only Numbers" -ForegroundColor Red
                        exit
                    }
                    IF ($Batch) {
                        $CSVArray = @()
                        $Number = [Math]::Ceiling($ICSV.Count / $Batch)
                        For ($i = 0; $I -lt $Batch; $i++) {
                            $SR = ($i * $Number)
                            $ER = (($i + 1) * $Number) - 1
                            If ($ER -ge $ICSV.Count) { $ER = $ICSV.Count }
                            $CSVArray += , @($ICSV[$SR..$ER])
                        }
                    }
                    Write-Host "Number of CSV Files created will be : " $CSVArray.count
                    If (!($Fpath)) { $FPath = $PPath }


                    Write-Host "Current Folder Path :: " $Fpath

                    0..(($CSVArray.count) - 1) | ForEach-Object {
                        $CD = $_ + 1
                        $CSVArray[$_] | Export-Csv -Path $FPath\$($CD)_DATABatch.csv -NoTypeInformation -Force }
                    exit
                } 
                '3' {
                    ################################################Alphabets Logs######################################################################

                    [array]$AB = @('a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i', 'j', 'k', 'l', 'm', 'n', 'o', 'p', 'q', 'r', 's', 't', 'u', 'v', 'w', 'x', 'y', 'z')
                    $SRR = $null
                    $ERR = $null
                    $SR = $null
                    $ER = $null

                    $DT = @()

                    $inputRange = Read-Host "Pleas Enter Range of Alphabets" 
                    $inputRange = $inputRange.Replace(" ", "")
                    $Range = $inputRange.Split('-')
                    $SRR = $Range[0]
                    $ERR = $Range[1]

                    $inputSR = $SRR
                    Switch ($inputSR) {
                        a { $SR = 0 }
                        b { $SR = 1 }
                        c { $SR = 2 }
                        d { $SR = 3 }
                        e { $SR = 4 }
                        f { $SR = 5 }
                        g { $SR = 6 }
                        h { $SR = 7 }
                        i { $SR = 8 }
                        j { $SR = 9 }
                        k { $SR = 10 }
                        l { $SR = 11 }
                        m { $SR = 12 }
                        n { $SR = 13 }
                        o { $SR = 14 }
                        p { $SR = 15 }
                        q { $SR = 16 }
                        r { $SR = 17 }
                        s { $SR = 18 }
                        t { $SR = 19 }
                        u { $SR = 20 }
                        v { $SR = 21 }
                        w { $SR = 22 }
                        x { $SR = 23 }
                        y { $SR = 24 }
                        z { $SR = 25 }

                    }

                    $inputER = $ERR
                    Switch ($inputER) {
                        a { $ER = 0 }
                        b { $ER = 1 }
                        c { $ER = 2 }
                        d { $ER = 3 }
                        e { $ER = 4 }
                        f { $ER = 5 }
                        g { $ER = 6 }
                        h { $ER = 7 }
                        i { $ER = 8 }
                        j { $ER = 9 }
                        k { $ER = 10 }
                        l { $ER = 11 }
                        m { $ER = 12 }
                        n { $ER = 13 }
                        o { $ER = 14 }
                        p { $ER = 15 }
                        q { $ER = 16 }
                        r { $ER = 17 }
                        s { $ER = 18 }
                        t { $ER = 19 }
                        u { $ER = 20 }
                        v { $ER = 21 }
                        w { $ER = 22 }
                        x { $ER = 23 }
                        y { $ER = 24 }
                        z { $ER = 25 }
                    }

                    $Para = ($ICSV | Get-Member -MemberType NoteProperty).Name
                    $Para = $Para | Select-Object | Sort-Object 

                    $PData = @()

                    $Para | ForEach-Object {
                        $PData += $PSItem
                        $PDataC = $PData.count
                        Write-Host $PDataC") " $PSItem -ForegroundColor Green
                    }

                    [int]$Col = Read-Host "Please press the Number to select the Column Name to Split CSV Alphabeticaly"

                    $r = 1..($Para.Count)

                    if ($Col -in $r) {

                        $n = $Col - 1

                        $CCol = $PData[$n]

                    }
                    else { write-host "Selected Number out of Range" -ForegroundColor Red }

                    Write-Host $CCol -ForegroundColor Yellow

                    if (!($ER)) { $ER = $SR }


                    $SR..$ER | ForEach-Object {
                        $CDA = $AB[$PSItem]
                        $ICSV | ForEach-Object {
                            $CN = $PSItem
                            $CT = $CN | Where-Object { $_.($CCol) -like ($($CDA) + '*') }
                            $DT += $CT
                        }
                    }
                    if (!($ERR)) { $ERR = $SRR }

                    $RR = $srr.ToUpper() + "-" + $ERR.ToUpper()

                    "`n"

                    Write-Host "Total Items Count :: $($ICSV.Count)"
                    "`n"

                    Write-Host "Count for User Input Alphabet Range :: $RR :: $($DT.Count)" -ForegroundColor Cyan
                    If (!($Fpath)) { $FPath = $PPath }


                    Write-Host "Current Folder Path :: " $Fpath


                    $DT | Export-Csv -Path $FPath\$($RR)_Alphabatical_data.csv -NoTypeInformation -Force
                    exit

                    ######################################################################################################################
                } 
                '4' {
                    Try {
                        $Fpath = Read-Host "Please Enter Folder Path to Export Split CSV Files" -ErrorAction Stop
                        If (!(Test-Path $Fpath)) {
                            Write-Host "Folder Path does not exists"
                            pause
                            Show-Menu
                        }
                        else {
                            Write-Host "Below Folder Path Added Successfully"
                            Write-Host $Fpath
                        }
                        Set-Variable -Scope script -Name ppath -Value ($Fpath)
                        pause
                        Show-Menu
                    }
                    Catch { Write-Host "Please Enter Correct Format for Folder Path" -ForegroundColor Red }
                
                    Show-Menu

                } 
                'q' {
                    return
                }
            }
            pause
        }
        until ($input -eq 'q')
    }
    else { Write-Host "Please enter Correct Path for CSV File" }
}


