[CmdletBinding()]
Param(
    [Parameter(Mandatory=$False)]
    [int] $ExclusionPeriod = 0
)
[void][reflection.assembly]::LoadWithPartialName("Microsoft.UpdateServices.Administration")
try{
    $ws = [Microsoft.UpdateServices.Administration.AdminProxy]::GetUpdateServer()
}
catch [System.Exception]{
    Write-Host "Failed to connect."
    Write-Host "Error:" $_.Exception.Message
    $ws = $null
}
if($ws -eq $null){return}
Write-Host "Connected"
$ca = 0
$cs = 0
$csp = 0
$cd = 0
try{
    Write-Host "Getting Updates"
	$au = $ws.GetUpdates()
}
catch [System.Exception]{
    Write-Host "Failed to get updates."
	Write-Host "Error:" $_.Exception.Message
    return
}
Write-Host "Processing Updates"
foreach($u in $au){
    $ca++
    if ($u.IsDeclined){
        $cd++
    }
    if (!$u.IsDeclined -and $u.IsSuperseded){
        $cs++     
        if ($u.CreationDate -lt (get-date).AddDays(-$ExclusionPeriod)){
		    $csp++
        }		       
    }
}
Write-Host "All Updates: " $ca
Write-Host "All Superseded updates (Not Declined): " $cs
Write-Host "Superseded updates (Older than $ExlusionPeriod days): $csp"
function declineUpdates{
    $i = 0
    $ud = 0
    Write-Host "Declining Updates"
    foreach($u in $au){
        if($u.IsDeclined -and $u.IsSuperseded){
            if($u.CreationDate -lt (Get-Date).AddDays(-$ExclusionPeriod)){
                $i++
                $pc = "{0:N2}" -f (($ud/$ca) * 100)
                Write-Progress -Activity "Declining Updates" -Status "Declining update #$i/$ca - $($u.Id.UpdateId.Guid)"`
                    -PercentComplete $pc -CurrentOperation "$($pc)% complete"
                try{
                    $u.Decline()
                    $ud++
                }
                catch [System.Exception]{
                    Write-Host "Failed to decline update $($u.Id.UpdateId.Guid)"
                }
            }
        }
    }
    Write-Host "Declined $ud updates"
}
$c = Read-Host "Do you want to continue and decline all updates older than $ExclusionPeriod days? y or n "
if($c -eq "n"){"";"Exiting session";return}
elseif("n" -and "y" -ne $c){Write-Host "Invalid option, exiting session";return}
else{
    declineUpdates
}