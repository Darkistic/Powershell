$userProfile = $env:USERPROFILE
$dateToday =  get-date -Format "dd/MM/yyyy"
$path = "$userProfile\Documents\Reports\$pattern WordSearch $(get-date -f dd-MM-yyy).txt"
$array = @()
$Document = [string]
$pattern = [string]
$fileExtension = [string]
$searchResult = [string]
$file =$Document
$pwd = pwd
#-------VARs


function retrieveParameters{
    $global:pattern = Read-Host "What text is present on the file you are looking for?"
    $global:fileExtension = Read-Host "What is the file extension (docx,txt,pdf,xlsx..) ?"
    }#get paraments from user input
function searchArray {
    $global:searchResult = Get-ChildItem -Recurse  -filter *.$global:fileExtension | ForEach-Object {$_.FullName
    }
} #-- Get the files and place them on a array


function checkFile{

    ForEach ($x in $searchResult) {
        $row = New-Object Object
        $pattern = $global:pattern
        Write-Host "Analysing file $x" -ForegroundColor yellow
    
             Try{$result = findWord $x $pattern}
             Catch{}
             if ($result -eq "True")
                 {write-host "'$global:pattern' is present in the file" -ForegroundColor green  }

             else
                 {write-host "The words `$global:pattern` does NOT exist in the file" -ForegroundColor Red  }
        #--creates array to be added to the file     
        $row | Add-Member -MemberType NoteProperty -Name "Word Searched" -Value $pattern
        $row | Add-Member -MemberType NoteProperty -Name "File Location" -Value $x             
        $global:array += $row
        

        


    } #---------------------close ForEach


}#---Run findWord on each object inside the list SearchArray 

function findWord ([string]$file,[string]$FindText) {
 $Document = $x
 $MatchCase = $False
 $MatchWholeWord = $True
 $MatchWildcards = $False
 $MatchSoundsLike = $False
 $MatchAllWordForms = $False
 $Forward = $True
 $Wrap = $FindContinue
 $Format = $False

 $Word = New-Object -comobject Word.Application
 $Word.Visible = $False

 $OpenDoc = $Word.Documents.Open($Document)
 $Selection = $Word.Selection

 $Selection.Find.Execute(
  $FindText,
  $MatchCase,
  $MatchWholeWord,
  $MatchWildcards,
  $MatchSoundsLike,
  $MatchAllWordForms,
  $Forward,
  $Wrap,
  $Format
  
 )

 $OpenDoc.Close()
} #--------Open Word Doc and Search for $pattern $fileExtension

function releaseMemory{
 # release memory
 $Word = $null
 # call garbage collection
 [gc]::collect()
 [gc]::WaitForPendingFinalizers()
 }#-- refresh the memory

function introduction{
sleep .4
cls
Write-Host "------------------------------------------------------------------" -ForegroundColor Yellow
Write-Host "You are about to perform a word search on the following location:" -foreground yellow -BackgroundColor Black
Write-Host ""
Write-Host $pwd -ForegroundColor Yellow
Write-Host ""
Write-Warning "If you want to search for the pattern on a different location `nplease change directories and run the script again." 
Write-Host "------------------------------------------------------------------" -ForegroundColor Yellow 
}#- makes it look good
function footer{
Write-Host "------------------------------------------------------------------" -ForegroundColor Yellow
Write-Host "                        Search completed                         " -BackgroundColor green -ForegroundColor Black
Write-Host ""
Write-Host "        Please check the .log file for the final results        " -ForegroundColor green
Write-Host ""
Write-Host "        Successful results for .$fileExtension files  containing the pattern '$global:pattern' are saved in the following location" -ForegroundColor green
Write-Host "        $path" -foreground Yellow
Write-Host "------------------------------------------------------------------" -ForegroundColor Yellow
}#- makes it look good

function createLog{$global:array | Export-Csv -Path $path -NoTypeInformation
} #-guess what this does..
function Word-Search($global:pattern, $global:fileExtension){
introduction
if ($global:pattern -eq $null -and $global:fileExtension -eq $null){retrieveParameters}
searchArray
checkFile
footer
createLog
releaseMemory

} #--Main Function

Word-Search
