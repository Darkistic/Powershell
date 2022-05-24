#-------THIS IS THE FTP VERSION@@@@@@@@@@@@@@@@@@@@@@
#-------THIS IS THE FTP VERSION@@@@@@@@@@@@@@@@@@@@@@
#-------THIS IS THE FTP VERSION@@@@@@@@@@@@@@@@@@@@@@
#-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+

#-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+


$userProfile = $env:USERPROFILE
$dataPath ="$userProfile\Downloads\data.csv"
$myCsvFile =Import-CSV "$dataPath" -Header "Employee ID", "Full Name", "Position", "Access Role", "Manager ID", "Manager Full Name", "Manager Position"
$art = @"

   _,="(  // )"=,_
_,'    \_>'\_/    ',_ 
.7,     {  }     ,\.
 '/:,  .m  m.  ,:\'
   ')",(/  \),"('
      '{'!!'}' - All is looking good here, you should tread yourself with a coffee         
       (   ) ) 
        ) ( (             It's your lucky day, the report is empty.
    PB (____)___
    .-'---------|  
   ( C|/\/\/\/\/|
    '-./\/\/\/\/|
      '_________'
       '-------'
"@
#------------------- VARS --------------------------
function Search-Export {
$array = @()

    ForEach($object in $myCsvFile)

    {
    $adUser

    $row = New-Object Object
    If($object."Manager Position" -eq "Support Worker" -or 
           $object."Manager Position" -eq "Bank Support Worker" -or
           $object."Manager Position" -eq "Bank Nurse" -or
           $object."Manager Position" -eq "Senior Support Worker" -or
           $object."Manager Position" -eq "No Manager" -or
           $object."Manager Full Name" -eq "$null" -and
           $object.'Reportee Position Name' -ne "Chief Executive") 

        {
        $ID = $object."Employee ID"
        $name = $object."Full Name"
        $position = $object."Position"
        $mngId = $object."Manager ID"
        $mngName = $object."Manager Full Name"
        $mngPosition = $object."Manager Position"



        $row | Add-Member -MemberType NoteProperty -Name "Employee ID" -Value $ID
        $row | Add-Member -MemberType NoteProperty -Name "Full Name" -Value $name
        $row | Add-Member -MemberType NoteProperty -Name "Position" -Value $position
        $row | Add-Member -MemberType NoteProperty -Name "Manager ID" -Value $mngID
        $row | Add-Member -MemberType NoteProperty -Name "Manager Full Name" -Value $mngName
        $row | Add-Member -MemberType NoteProperty -Name "Manager Position" -Value $mngPosition

        $array += $row
        
        }# iteration throught list finished

        
        #set new vars for fle and path
        $dateToday =  get-date -Format "dd/MM/yyyy"
        $global:path = "$userProfile\Documents\Reports\Wrong Reporting Manager SFTP $(get-date -f dd-MM-yyy).csv"
        #---Tries to create a new folder to output the result
Try{
        Write-Host "Thinking..." -Foreground green
        New-Item -ItemType Directory -Name "Reports" -Path "$userProfile\Documents\" -ErrorAction Stop

        }
Catch { }
}#---close search and export 
        
        $array | Export-Csv -Path $path -NoTypeInformation
        


} #--close Search-Export

function Import-Result {
Clear-Host
sleep -Milliseconds 1000
$out = Import-Csv $global:path -Header "Employee ID", "Full Name", "Position", "Manager ID", "Manager Full Name", "Manager Position"

Try{
        ForEach ($x in $out){
        $name = $x.'Full Name'
        $empID = $x."Employee ID"
        $mgName = $x.'Manager Full Name'
        if( $mgName.Length -lt 1 ){
Write-host "$name ($empID)       is reporting to a            Ghost Manager"  -ForegroundColor Magenta -BackgroundColor black}
        Else{
Write-Host " $name ($empID) is reporting to the WRONG Manager ---> $mgName" -ForegroundColor Red -BackgroundColor black           
            
            }#close FoEach - Else

        }#--close ForEach

}#--close Try

Catch{Write-Host "I could not read file check if the file was created"}
}

function Check-Result {
Try{
    $out = Import-CSv $global:path
    if($out -eq $null){
    Clear-Host
        While ($x -le 10){
        $sleep = .5
        Write-Host $art -ForegroundColor Yellow
        sleep $sleep
        Clear-Host
        Write-Host $art -ForegroundColor Green
        sleep $sleep
        Clear-Host
        $x += 1}
        }

    else{
    Write-Warning "There are employees reporting to the wrong Manager."
    Write-Host "---------------------------------------------------------------" -ForegroundColor Red 
    Write-host "--------------------------------------------------------------" -foreground Yellow
    Write-Host "Please check each of their UNITS to see who they should be reporting to." -ForegroundColor Yellow
    Write-Host " You will find the report on the following location: "  -foreground Green
    Write-Host "$path"  -foreground yellow
    Write-Host "---------------------------------------------------------------" -ForegroundColor Red
    }
}#---close try
Catch{Write-Host "No records found"}
}#close function

Function Email ($attachment,$subject,$body,[string[]]$to){
    $From = "paulo.bazzo@fitzroy.org"
    $SMTPServer = "smtp.office365.com"
    $SMTPPort = "587"
    $To = "paulo.bazzo@fitzroy.org"
    $subject = "Reporting Managers - Daily Check"
    $body = "hello"
    $Creds = (Get-Credential -Credential "$From")
    
   

    Try
    {
        Send-MailMessage -From $From -to $To -Subject $subject `
        -Body $body -SmtpServer $SMTPServer -port $SMTPPort -ErrorAction Stop -Attachments $path -UseSsl -Credential $MyCredential -DeliveryNotificationOption never
<#
        DeliverNotificationOption
        -- None: No notification
        -- OnSuccess: Notify if the delivery is successful
        -- OnFailure: Notify if the delivery is unsucessful
        -- Delay: Notify if the delivery is delayed
        -- Never: Never notify
        
#>        
    }
    Catch
    { Write-Host "something went wrong" }

}

function Run {

Search-Export
Import-Result
Check-Result
#Email
}


Run
