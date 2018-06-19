#============================================================================================
#Decide what year prefix a user should be assigned
#============================================================================================

$dateM = Get-Date -UFormat "%B"
$currentY = Get-Date -UFormat "%Y"

$WildcardM = "*$dateM*"
$monthJtoA = "January,February,March,April,May,June,July,August"

#prefixes for Year 7,8,9,10 and 11

if ($monthJtoA -like $WildcardM) {
    $year7 = $currentY-1
    $year8 = $currentY-2
    $year9 = $currentY-3
    $year10 = $currentY-4
    $year11 = $currentY-5
}
else {
    $year7 = $currentY
    $year8 = $currentY-1
    $year9 = $currentY-2
    $year10 = $currentY-3
    $year11 = $currentY-4
}

$users = Import-Csv -Path '.\SIMS New Additions.csv'

foreach ($user in $users) {
    $year = $user.Year
    $DOA = $user.DOA
    $name = $user.name
    $surname = $user.'Legal Surname'
    $firstname = $user.Forename

    if ($year = 'Year 7') {
        
    }
}
