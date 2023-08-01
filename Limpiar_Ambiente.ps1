Add-Type -AssemblyName System.Windows.Forms

function Show-YesNoMessageBox {
    [CmdletBinding()]
    param (

        $mensaje
    )
    [System.Windows.Forms.DialogResult]$result = [System.Windows.Forms.MessageBox]::Show(
        $mensaje,
        "Confirmation",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )

    return ($result -eq [System.Windows.Forms.DialogResult]::Yes)
}


$start_runtime = Get-Date
Import-Module smlets
#$servidor = "s1-dixx-ssm04"
$servidor =  "s1-hixx-ssm01"
$Count_only = $false 



$IRClass = Get-SCSMClass -Name System.WorkItem.Incident$ -ComputerName $servidor 
$SRClass = Get-SCSMClass -Name System.WorkItem.ServiceRequest$ -ComputerName $servidor
$MAclass = Get-SCSMClass -Name System.WorkItem.Activity.ManualActivity$ -ComputerName $servidor


$irObjects = Get-SCSMObject -Class  $IRclass -ComputerName $servidor
$srObjects   = Get-SCSMObject -Class  $SRClass -ComputerName $servidor
$maObjects = Get-SCSMObject -Class $MAclass  -ComputerName $servidor


$output_counts = @"
    Incidentes: $($irObjects.Count)
    Acitividades: $($maObjects.Count)
    Solicitudes: $($srObjects.Count)
"@
write-host $output_counts




if($Count_only -eq $false){
    # Call the function to prompt the user
    $confirmation = Show-YesNoMessageBox "El servidor es $($servidor.ToUpper()), continuar?"
    if ($confirmation){
    $confirmation = Show-YesNoMessageBox "El servidor es $($servidor.ToUpper()), seguro que quiere continuar?"
    }
    if ($confirmation) {
        $irObjects  | Remove-SCSMObject -Force -ComputerName $servidor 
        $maObjects | Remove-SCSMObject -Force -ComputerName $servidor 
        $srObjects | Remove-SCSMObject -Force -ComputerName $servidor 
    }
}

$end_runtime = Get-Date


$total_runtime = $end_runtime - $start_runtime

# Display the total run time in seconds
Write-Host "Total run time: $($total_runtime.TotalMinutes) Minutes."