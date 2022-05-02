#$userProfile = $env:USERPROFILE
$dateToday =  get-date -Format "dd/MM/yyyy"
$path = "C:\Test\Word Search $(get-date -f dd-MM-yyy).txt"
$array = @()
$Document = [string]
$pattern = [string]
$searchResult = 0
$file =$Document
$pwd = pwd
$filesFound = @()


function searchArray {
    $global:pattern = Read-Host "What is text is present on the file?"
    $fileExtension = Read-Host "What is the file extension"
    $global:searchResult = Get-ChildItem -Recurse  -filter *.$fileExtension | ForEach-Object {$_.FullName
    #Write-Host $global:searchResult
    }
} #-- Get the files and place them on a array

function checkFile{

    ForEach ($x in $searchResult) {
        $pattern = $global:pattern
        Write-Host "Analysing file $x" -ForegroundColor yellow
    
             Try{$result = findWord $x $pattern}
             Catch{}
             if ($result -eq "True")
                 {write-host "String Exist" -ForegroundColor green  }

             else
                 {write-host "String Does Not Exist" -ForegroundColor red  }
             $global:filesFound = $x + $x
        

        


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
} #--------Open Word Doc and Search for $pattern

function releaseMemory{
 # release memory
 $Word = $null
 # call garbage collection
 [gc]::collect()
 [gc]::WaitForPendingFinalizers()
 }

function introduction{
sleep 1
cls
Write-Host "------------------------------------------------------------------" -ForegroundColor Yellow
Write-Host "You are about to perform a word search on the following location:" -foreground yellow -BackgroundColor Black
Write-Host ""
Write-Host $pwd -ForegroundColor Yellow
Write-Host ""
Write-Warning "If you want to search for the pattern on a different location please change directories and run the script again" 
Write-Host "------------------------------------------------------------------" -ForegroundColor Yellow 
}
function footer{
Write-Host "            Search completed           " -BackgroundColor green -ForegroundColor Black
Write-Host please check the .log file for the final results -ForegroundColor green
Write-Host "file location: $path" -foreground Yellow
}

function createLog{$array | Export-Csv -Path $path -NoTypeInformation
}
function Start-Search{
introduction
searchArray
checkFile
footer
createLog
releaseMemory

}



Start-Search
