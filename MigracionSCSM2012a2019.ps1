[Threading.Thread]::CurrentThread.CurrentCulture = 'es-ES'
#region modulos
Import-Module .\MigracionSCSM\MigracionSCSM.psd1  -Force
Import-Module smlets
#endregion 

#region parametros
$servidorOrigen = "s4-dixx-ssm01"
$servidorDestino = "s1-dixx-ssm04"
$basePath = "C:\temp\migracion\reqExport"
$logPath ="C:\temp\migracion\logs\logs_migracion.txt"
$logPathSoloIncidentesPadre = "C:\temp\migracion\logs\logs_migracion_IncPadres.txt"
$logPathSoloIncidentesHijos = "C:\temp\migracion\logs\logs_migracion_IncHijos.txt"
$directorio = "C:\temp\migracion\reqExport"
#enregion

#region clases
$SRClassOrigen = Get-SCSMClass -Name System.WorkItem.ServiceRequest$ -ComputerName $servidorOrigen
$IRclassOrigen = Get-SCSMclass -name System.Workitem.Incident$ -ComputerName $servidorOrigen 
$IRclassDestino = Get-SCSMclass -name System.Workitem.Incident$ -ComputerName $servidorDestino 
$SRclassDestino = Get-SCSMClass -Name system.workitem.servicerequest$ -ComputerName $servidorDestino
$UserClass = Get-SCSMClass -name System.Domain.User$ -ComputerName $servidorOrigen # Get SCSM User class object
#endregion 

#region relaciones
$relAffectedUser = Get-SCSMRelationshipClass -Name System.WorkItemAffectedUser -ComputerName $servidorOrigen # Get SCSM relationship Affected User
$AssignedToRel = get-scsmrelationshipclass -name System.WorkItemAssignedToUser$ -ComputerName $servidorOrigen
$WorkItemAffectedUserRel = Get-SCSMRelationshipClass System.WorkItemAffectedUser$ -ComputerName $servidorOrigen
$AssignedToUserRel = Get-scsmrelationshipclass -name System.WorkItemAssignedToUser$ -ComputerName $servidorOrigen
$WorkItemAffectedUserRel = Get-SCSMRelationshipClass System.WorkItemAffectedUser$ -ComputerName $servidorOrigen
$rel_padre_hijo = Get-SCSMRelationshipClass -Name System.WorkItemHasParentWorkItem -ComputerName $servidorOrigen 


$padreHijoRel = Get-SCSMRelationshipClass -Name System.WorkItemRelatesToWorkItem -ComputerName $servidorDestino

#endregion

#region funciones
function write-log {
    param (
        $wi_id,
        $logName
    )

    $fechaActual = Get-Date -Format "yyyy-MM-dd hh:mm:ss"
    switch($logName)
    {
        padres {$outputPath =  $logPathSoloIncidentesPadre ;
               $mensaje = -join($fechaActual," - " , $wi_id, " - " ,"$($wi_id)" )           
                }
        hijos {}
        general {$outputPath =  $logPath ; 
                 $mensaje = -join($fechaActual," - " , $wi_id, " - " ,"Procesando $($wi_id)" )         
        }

    }
 
    write-host $mensaje  -ForegroundColor Yellow
    $mensaje | out-file $outputPath -Append
}
#enregion
#region main


    #region incidentes
    #objtengo todos los indicentes no cerrados y no resueltos
    $objIR = get-MigracionSCSM_Requerimientos -clase ir -status "irNoCloseNoResolved" -servidor $servidorOrigen 


        #region incidentes padres
        $testIsparent = $objIR |? {$_.IsParent -eq  "True"} 
        #endregion

        $testIsparent[0] | ForEach-Object {
            $wi = $_

            write-log $wi.id "general"
            write-log $wi.id "padres"

            $AffectedUser = Get-SCSMRelatedObject  -Relationship $WorkItemAffectedUserRel -SMObject $wi -ComputerName $servidorOrigen
            $AssignedToUser = Get-SCSMRelatedObject  -Relationship $AssignedToUserRel -SMObject $wi -ComputerName $servidorOrigen

            
           $childWis_relObj = Get-SCSMRelationshipObject -ByTarget $wi -ComputerName $servidorOrigen -Relationship $rel_padre_hijo  #| Where-Object { $_.RelationshipId -eq 'da3123d1-2b52-a281-6f42-33d0c1f06ab4'}

        }
        
        $childWis_relObj | ForEach-Object {


         write-host $_.TargetObject.ClassName
         Write-Host $_.RelationshipId

        }
        # #region incidentes hijos
        # $hijos = $objIR | ? {$_.status.displayname -match "depende"}

        # $hijos[0] | ForEach-Object {
        #     #cambio de varable por legibilidad del codigo
        #     $wi = $_
        #     $origenID = $wi.Id
        #     #logueo la actividad
        #     write-log $wi.id
            
        #     #region obtener usuarios
        #     $AffectedUser = Get-SCSMRelatedObject  -Relationship $WorkItemAffectedUserRel -SMObject $wi -ComputerName $servidorOrigen
        #     $AssignedToUser = Get-SCSMRelatedObject  -Relationship $AssignedToUserRel -SMObject $wi  -ComputerName $servidorOrigen
            

        #     #endregion

        #     get-MigracionSCSM_AttachReq -wi $wi.id -OutputFolder $basePath -servidor $servidorOrigen 

        #     #region actionlog
        #         #te lo debo
        #     #endregion


        # }


        # #endregion



        # #region incidentes Generales
        # $objetosFinales =  $objIR | ? {$_.IsParent -ne  "True" -and $_.status.displayname -notmatch "depende"}

        # #endregion


    #endregion


    #region Solicitudes


    #endregion
#endregion