


# Import-Module .\MigracionSCSM\MigracionSCSM.psd1  -Force
# show-variable 


# $filtro = switch ($status)
# {
#     iractive { "Status -eq '{0}'" -f $IrActive ; break}
#     irprogress {"Status -eq '{0}'" -f $IrProgress ; break}
#     irpending { "Status -eq '{0}'" -f $IrPending ; break}
#     irpadre { "Status -eq '{0}'" -f $IrPadre ; break}
#     srInProgress {"Status -eq '{0}'" -f $SRInProgress ; break}
#     irNoCloseNoResolved{  "Status -ne '{0}' -and Status -ne '{1}'" -f $IrResolved, $Irclose  ; break}
#     srNoCompletedCancelClosed{ "Status -ne '{0}' -and Status -ne '{1}'-and Status -ne '{2}'" -f $SRCompleted, $SRCanceled, $SRClosed ; break}
# }

# Write-Host $estado
# Write-Host $filtro


# $SRCompleted = "hola"
# $SRCanceled = "como"
# $SRClosed = "va?"
# "Status -ne '{0}' -and Status -ne '{1}'-and Status -ne '{2}'" -f $SRCompleted, $SRCanceled, $SRClosed


# "Status -ne '$($SRCompleted)' -and Status -ne '$($SRCanceled)'-and Status -ne '$($SRClosed)'"



$actionlog1 = "log1"
$actionlog2 = $null
$actionlog3 = "log3"
$test = @{
    ActionLog = $actionlog1, $actionlo2, $actionlog3    

}