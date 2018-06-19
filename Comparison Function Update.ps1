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