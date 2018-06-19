#============================================================================================

#Script to automatically create or disable students after changes in SIMS

#============================================================================================




#============================================================================================
#Compare Difference between 'SIMS Student List OLD.csv' and 'SIMS Student List NEW.csv'
#============================================================================================


$file1 = Import-Csv -Path ".\SIMS Student List OLD.csv"
$file2 = Import-Csv -Path ".\SIMS Student List NEW.csv"

Compare-Object $file1 $file2 -property "Year","DOA","Name","Legal Surname","Forename" -IncludeEqual | export-csv -path ".\Comparison results.csv" -NoTypeInformation

$comparisons = Import-Csv -Path ".\Comparison results.csv"


#============================================================================================
#IF statement to determine if there are any changes in SIMS (Compares NEW and OLD lists)
#============================================================================================


$comp = $comparisons.SideIndicator
if (($comp -notcontains "=>") -or ($comp -notcontains "<=")) {
    Write-Host "No user changes have been made in SIMS. Exiting script." -ForegroundColor Yellow
    exit
}
else {
    Write-Host "User changes have been made in SIMS. Active Directory will update with changes:" -ForegroundColor Yellow

    $csvPath1 = ".\SIMS NEW Additions.csv"
    $csvPath2 = ".\SIMS Off Roll.csv"

    Clear-Content -Path $csvPath1,$csvPath2

    $hLine = "{0},{1},{2},{3},{4}" -f "Year","DOA","Name","Legal Surname","Forename"
    $hLine | Add-Content -Path $csvPath1,$csvPath2

    foreach ($comparison in $comparisons) 
    {
        $year = $comparison.Year
        $DOA = $comparison.DOA
        $name = $comparison.Name
        $surname = $comparison."Legal Surname"
        $firstname = $comparison.Forename
        $sideIndicator = $comparison.SideIndicator
        $xLine = "{0},{1},{2},{3},{4}" -f $year,$DOA,$name,$surname,$firstname
        
        if ($sideIndicator -eq "=>") {
            $xLine | Add-Content -Path $csvPath1
            }
        else {
            $xLine | Add-Content -Path $csvPath2
            }
    }


    Remove-Item -Path 'C:\Scripts\Create Students - SIMS integration\SIMS Student List OLD.csv'
    Get-Item -Path 'C:\Scripts\Create Students - SIMS integration\SIMS Student List NEW.csv' | Rename-Item -NewName "SIMS Student List OLD.csv"


    #Clear Change.csv
    Clear-Content -Path '.\Changes.csv'
    $heading1 = "{0},{1},{2},{3},{4},{5}" -f "Year","Surname","Firstname","Username","Password","Change"
    $heading1 | Add-Content -Path '.\Changes.csv'


#============================================================================================
#Create accounts for users in 'SIMS NEW Additions.csv' (Curriculum and Controlled Assessment)
#============================================================================================


    Import-Module ActiveDirectory

    $users = Import-Csv -Path ".\SIMS New Additions.csv"

    foreach ($user in $users)
        {

            #Variables from 'SIMS NEW Additions.csv'
            $surname = $user."Legal Surname"
            $firstname = $user.Forename
            $year = $user."Academic Year Of Entry".Substring(2,2)


            #Variables
            $f3Letters = $firstname.Substring(0,3)
            $usernameA = $year + $surname + $f3Letters
            $memberOfA = $year + ",RDP Student Allow,Students"
            $UPN = $usernameA + "@weyvalley.dorset.sch.uk"
            $displayname = $firstname + " " + $surname
            $shareNameA = $usernameA + "$"
            $homeDirectoryA = "\\stusrv\" + "$shareNameA"
            $targetOUA = "OU=Year " + "$year" + ",OU=Students,DC=nsnet,DC=net"
            $uncPathA = "\\stusrv\users`$\Students\$year\$usernameA"
            $homeDrive = "N:"
            

            #Controlled assessment variables
            $usernameB = $year + "CA-" + $surname + $f3Letters
            $shareNameB = $usernameB + "$"
            $homeDirectoryB = "\\stusrv\" + "$shareNameB"
            $targetOUB = "OU=Assess" + "$year" + ",OU=Students,DC=nsnet,DC=net"
            $uncPathB = "\\stusrv\users`$\Students\" + $year + "Assessed\$usernameB"
            $memberOfB = $year + "Assess,Students"


            #Generate Password
            $random1 = Import-csv -path "C:\Scripts\Password Generator\Passphrase List.csv" | get-random
            $random2 = Import-csv -path "C:\Scripts\Password Generator\Passphrase List.csv" | get-random
        
            $randomWord1 = $random1.word
            $randomWord2 = $random2.word
        
            $password = "$randomWord1" + "." + "$randomWord2"


            #Share Variables
            $folderPathA = "D:\Users\Students\$year\$usernameA"
            $netShareA = {
            param($shareNameA,$folderPathA)
            Net Share $shareNameA=$folderPathA /GRANT:EVERYONE`,FULL }

            $folderPathB = "D:\Users\Students\$year`Assessed\$usernameB"
            $netShareB = {
            param($shareNameB,$folderPathB)
            Net Share $shareNameB=$folderPathB /GRANT:EVERYONE`,FULL }


            #Create student user account
            Try
            {
                New-ADUser -Name $usernameA -GivenName $firstname -Surname $surname -UserPrincipalName $UPN -SamAccountName $usernameA -DisplayName $displayname -Description $displayname -HomeDrive $homeDrive -HomeDirectory $homeDirectoryA -Path $targetOUA -AccountPassword (convertto-securestring $password -AsPlainText -Force) -Enabled $false
            }

            Catch
            {
                Get-Date -UFormat "%d-%m-%Y / %T > Failed to create user $usernameA`: $_" | Add-Content $errorLog
            }


            #Set Password to change at next logon
            Try
            {
                Set-ADUser -Identity $usernameA -ChangePasswordAtLogon $True
            }

            Catch
            {
                Get-Date -UFormat "%d-%m-%Y / %T > Failed to set password for user $usernameA`: $_" | Add-Content $errorLog
            }

            #Set User Proxy Address and Script Path
            Try
            {
                Set-ADUser -Identity $usernameA -Add @{
                'proxyAddresses' = $UPN | ForEach-Object { "SMTP:$_" } 
                'scriptPath' = "Main.bat" }
            }

            Catch
            {
                Get-Date -UFormat "%d-%m-%Y / %T > Failed to set Proxy Address and Script Path for user $usernameA`: $_" | Add-Content $errorLog
            }

            #Create User HomeDirectory
            Try
            {
                New-Item -Path $uncPathA -ItemType Directory
            }

            Catch
            {
                Get-Date -UFormat "%d-%m-%Y / %T > Failed to create Home Directory for user $usernameA`: $_" | Add-Content $errorLog
            }

            #Create Share
            Try
            {
                Invoke-Command -ComputerName STUSRV -ScriptBlock $netShareA -ArgumentList $shareNameA,$folderPathA
            }

            Catch
            {
                Get-Date -UFormat "%d-%m-%Y / %T > Failed to create Share for user $usernameA`: $_" | Add-Content $errorLog
            }

            #Permission Variables (This will break if moved further up)
            $aclA = (Get-Item $uncPathA).GetAccessControl('Access')
            $arA = New-Object System.Security.AccessControl.FileSystemAccessRule($usernameA,'FullControl','ContainerInherit,ObjectInherit','None','Allow')


            #Set Security Permissions
            Try
            {
                $aclA.SetAccessRule($arA)
                Set-Acl -Path $uncPathA -AclObject $aclA
            }

            Catch
            {
                Get-Date -UFormat "%d-%m-%Y / %T > Failed to set Security Permissions for user $usernameA`: $_" | Add-Content $errorLog
            }

            #Add User to Member Groups
            Try
            {
                $memberOfA.split("{,}") | Add-ADGroupMember -Members $usernameA 
            }

            Catch
            {
                Get-Date -UFormat "%d-%m-%Y / %T > Failed to add user $usernameA to Groups: $_" | Add-Content $errorLog
            }

            #Add Account Details to Changes.csv
            Try 
            {
                $csvPath = ".\Changes.csv"
                $nLine = "{0},{1},{2},{3},{4},{5}" -f $year,$surname,$firstname,$usernameA,$password,"Addition"
                $nLine | Add-Content -Path $csvPath
            }

            Catch
            {
                Get-Date -UFormat "%d-%m-%Y / %T > Failed to add $usernameA details to Changes.csv: $_" | Add-Content $errorLog
            }


            #===========================================================================


            #Create controlled assessment account
            Try
            {
                New-ADUser -Name $usernameB -GivenName $firstname -Surname $surname -UserPrincipalName $usernameB -SamAccountName $usernameB -DisplayName $displayname -Description $displayname -HomeDrive $homeDrive -HomeDirectory $homeDirectoryB -Path $targetOUB -AccountPassword (convertto-securestring $password -AsPlainText -Force) -Enabled $false
            }

            Catch
            {
                Get-Date -UFormat "%d-%m-%Y / %T > Failed to create user $usernameB`: $_" | Add-Content $errorLog
            }


            #Set Password to change at next logon
            Try
            {
                Set-ADUser -Identity $usernameB -ChangePasswordAtLogon $True
            }

            Catch
            {
                Get-Date -UFormat "%d-%m-%Y / %T > Failed to set password for user $usernameB`: $_" | Add-Content $errorLog
            }


            #Set User Script Path
            Try
            {
                Set-ADUser -Identity $usernameB -Add @{ 
                'scriptPath' = "Main.bat" }
            }

            Catch
            {
                Get-Date -UFormat "%d-%m-%Y / %T > Failed to set Script Path for user $usernameB`: $_" | Add-Content $errorLog
            }


            #Create User HomeDirectory
            Try
            {
                New-Item -Path $uncPathB -ItemType Directory
            }

            Catch
            {
                Get-Date -UFormat "%d-%m-%Y / %T > Failed to create Home Directory for user $usernameB`: $_" | Add-Content $errorLog
            }


            #Create Share
            Try
            {
                Invoke-Command -ComputerName STUSRV -ScriptBlock $netShareB -ArgumentList $shareNameB,$folderPathB
            }

            Catch
            {
                Get-Date -UFormat "%d-%m-%Y / %T > Failed to create Share for user $usernameB`: $_" | Add-Content $errorLog
            }


            #Permission Variables (This will break if moved further up)
            $aclB = (Get-Item $uncPathB).GetAccessControl('Access')
            $arB = New-Object System.Security.AccessControl.FileSystemAccessRule($usernameB,'FullControl','ContainerInherit,ObjectInherit','None','Allow')


            #Set Security Permissions
            Try
            {
                $aclB.SetAccessRule($arB)
                Set-Acl -Path $uncPathB -AclObject $aclB
            }

            Catch
            {
                Get-Date -UFormat "%d-%m-%Y / %T > Failed to set Security Permissions for user $usernameB`: $_" | Add-Content $errorLog
            }


            #Add User to Member Groups
            Try
            {
                $memberOfB.split("{,}") | Add-ADGroupMember -Members $usernameB 
            }

            Catch
            {
                Get-Date -UFormat "%d-%m-%Y / %T > Failed to add user $usernameB to Groups: $_" | Add-Content $errorLog
            }


            #Add Account Details to Changes.csv
            Try 
            {
                $csvPath = ".\Changes.csv"
                $nLine = "{0},{1},{2},{3},{4},{5}" -f $year,$surname,$firstname,$usernameB,$password,"Addition"
                $nLine | Add-Content -Path $csvPath
            }

            Catch
            {
                Get-Date -UFormat "%d-%m-%Y / %T > Failed to add $usernameB details to Changes.csv: $_" | Add-Content $errorLog
            }
        }



#============================================================================================
#Disable accounts from 'SIMS Off Roll.csv' (Curriculum and Controlled Assessment)
#============================================================================================


    $users = Import-Csv -Path '.\SIMS Off Roll.csv'

    foreach ($user in $users)
        {
            
            #Variables from 'SIMS Off Roll.csv'
            $surname = $user."Legal Surname"
            $firstname = $user.Forename
            $year = $user."Academic Year Of Entry".Substring(2,2)


            #Variables
            $f3Letters = $firstname.Substring(0,3)
            $usernameA = $year + $surname + $f3Letters
            $displayname = $firstname + " " + $surname
            $date = Get-Date -UFormat "%d-%m-%Y"


            #Controlled Assessment variables
            $usernameB = $year + "CA-" + $surname + $f3Letters


            #Disable curriculum account
            Set-ADUser -identity $usernameA -Description "$displayname - $date OFF ROLL" -Enabled $false
            

            #Add Account Details to Changes.csv
            $csvPath = ".\Changes.csv"
            $nLine = "{0},{1},{2},{3},{4},{5}" -f $year,$surname,$firstname,$usernameA,"N/A","Disabled"
            $nLine | Add-Content -Path $csvPath


            #Disable controlled assessment account
            Set-ADUser -identity $usernameB -Description "$displayname - $date OFF ROLL" -Enabled $false


            #Add Account Details to Changes.csv
            $csvPath = ".\Changes.csv"
            $nLine = "{0},{1},{2},{3},{4},{5}" -f $year,$surname,$firstname,$usernameB,"N/A","Disabled"
            $nLine | Add-Content -Path $csvPath
            
        }



#============================================================================================
#Create report of changes to 'Change Report.txt'
#============================================================================================


    $changes = Import-Csv '.\Changes.csv'

    $table = @()
    $reportDate = Get-Date -UFormat "%d-%m-%Y"
    Get-Date -UFormat "-----%d-%m-%Y-----" | Out-File ".\Reports\Change Report $reportDate.txt"
    $heading2 = "The following changes have been made in Active Directoy:"
    $heading2 | Out-File ".\Reports\Change Report $reportDate.txt" -Append

    foreach ($change in $changes)
    {
        $year = $change.Year
        $surname = $change.Surname
        $firstname = $change.Firstname
        $username = $change.Username
        $password = $change.Password
        $status = $change.Change

        $fullYear = "20" + $year

        $objContent = New-Object System.Object
        $objContent | Add-Member -Type NoteProperty -Name Year -Value $fullYear
        $objContent | Add-Member -Type NoteProperty -Name Surname -Value $surname
        $objContent | Add-Member -Type NoteProperty -Name Firstname -Value $firstname
        $objContent | Add-Member -Type NoteProperty -Name Username -Value $username
        $objContent | Add-Member -Type NoteProperty -Name Password -Value $password
        $objContent | Add-Member -Type NoteProperty -Name Change -Value $status 
        $table += $objContent
    }

    $table | Sort-Object Year -Descending | Format-Table -AutoSize -GroupBy Change -Property Year,Surname,Firstname,Username,Password | Out-File ".\Reports\Change Report $reportDate.txt" -Append



#============================================================================================
#Send mail message of report to technicians
#============================================================================================


    #Mail variables
    $mailFrom = 'userreport@weyvalley.dorset.sch.uk'
    $mailTo = 'hassallj@weyvalley.dorset.sch.uk','tibbeya@weyvalley.dorset.sch.uk'
    $mailSubject = "Changes have been made in Active Directory"
    $mailBody = "Attached is a report of changes made in Active Directory from data exported from SIMS."
    $mailAttachemtns = ".\Reports\Change Report $reportDate.txt"
    $mailServer = 'win2k8admin.nsnet.net'

    #Send email message
    Send-MailMessage -From $mailFrom -To $mailTo -Subject $mailSubject -BodyAsHtml -Body $mailBody -Attachments $mailAttachemtns -SmtpServer $mailServer

}