$start_runtime = Get-Date
$ErrorActionPreference = "Stop"
# https://www.stefanroth.net/2014/09/01/scsm-adding-activities-using-sma-powershell-workflow/
# https://community.cireson.com/discussion/3486/can-you-create-a-service-request-from-another-service-request-activity
[Threading.Thread]::CurrentThread.CurrentCulture = 'es-ES'
import-module SMlets

#region Usuario workflow (revisar)

$ServiceUser = "trabajo\appSCSM2019ProdWFL"
#$ServiceUser = "trabajo\appSCSM2019QaWFL"
#d1q4E3EO.,Vv1l@

#Revisar: sacar comentario a cred

<#
$cred = Get-Credential -credential $ServiceUser

if ( ($cred).length -eq "0"){
write-host "falta service user" -ForegroundColor Redx
break
}

#>

#endregion


#region constantes
$servidorOrigen = "scsm.ministerio.trabajo.gov.ar"
$servidorDestino = "s1-dixx-ssm04"
#$servidorDestino = "s1-hixx-ssm01"
$basePath = "C:\temp\migracion\reqExport\"
$logPath ="C:\temp\migracion\logs\logs_migracionSR.txt"
$pathFunciones = "D:\Trabajo\Github Repos\SCSM\work"
#$pathFunciones = "E:\Trabajo\Github Repos\SCSM\work"
$Registros_procesados_path = "C:\temp\migracion\logs\sr_procesados.txt"
#endregion


#region importar funciones
. $pathFunciones\get-requerimientos.ps1
. $pathFunciones\add-actionLogEntryV2.ps1
. $pathFunciones\get-AttachReqV2.ps1
. $pathFunciones\UploadAttachReqv2.ps1

Function Get-LocalTime($UTCTime)
{
$TZ = [System.TimeZoneInfo]::Local
$LocalTime = [System.TimeZoneInfo]::ConvertTimeFromUtc($UTCTime, $TZ)
Return $LocalTime
}


#endregion

#region clases
$SRClassOrigen = Get-SCSMClass -Name System.WorkItem.ServiceRequest$ -ComputerName $servidorOrigen
$IRclassOrigen = Get-SCSMclass -name System.Workitem.Incident$ -ComputerName $servidorOrigen 
$IRclassDestino = Get-SCSMclass -name System.Workitem.Incident$ -ComputerName $servidorDestino 
$SRclassDestino = Get-SCSMClass -Name system.workitem.servicerequest$ -ComputerName $servidorDestino
$UserClass = Get-SCSMClass -name System.Domain.User$ -ComputerName $servidorOrigen # Get SCSM User class object
$relAffectedUser = Get-SCSMRelationshipClass -Name System.WorkItemAffectedUser -ComputerName $servidorOrigen # Get SCSM relationship Affected User
$MaClassDestino = Get-SCSMClass -Name System.WorkItem.Activity.ManualActivity$  -ComputerName $servidorDestino
$MAclass = Get-SCSMClass -Name System.WorkItem.Activity.ManualActivity.Extended  -ComputerName $servidorDestino
#endregion
#region relaciones

$AssignedToRel = get-scsmrelationshipclass -name System.WorkItemAssignedToUser$ -ComputerName $servidorOrigen
$WorkItemAffectedUserRel = Get-SCSMRelationshipClass System.WorkItemAffectedUser$ -ComputerName $servidorOrigen
$AssignedToUserRel = Get-scsmrelationshipclass -name System.WorkItemAssignedToUser$ -ComputerName $servidorOrigen
$WorkItemAffectedUserRel = Get-SCSMRelationshipClass System.WorkItemAffectedUser$ -ComputerName $servidorOrigen
$WorkItemContainsActivityRelOrigen = Get-scsmrelationshipclass -name System.WorkItemContainsActivity$  -ComputerName $servidorOrigen
$WorkItemContainsActivityRelDestino = Get-SCSMRelationshipClass -Name System.WorkItemContainsActivity$ -ComputerName $servidorDestino
$manualActivitiesRel = Get-SCSMRelationshipClass -Name System.WorkItemContainsActivity$  -ComputerName $servidorOrigen

$RequestedByUserRel = Get-SCSMRelationshipClass  -Name System.WorkitemRequestedbyUser -ComputerName $servidorDestino 
$WorkitemHasParentWorkitemRel = Get-SCSMRelationshipClass -Name system.workitemhasparentworkitem -ComputerName $ServidorDestino
$WorkItemContainsActivityRel = Get-SCSMRelationshipClass -Name System.WorkItemContainsActivity -ComputerName $ServidorDestino


#Get-SCSMRelationshipClass  -ComputerName $servidorOrigen   | select * | Out-GridView

#endregion
#region projection

$serviceRequestTypeProjectionOrigen = Get-SCSMTypeProjection -name System.WorkItem.ServiceRequestProjection$   -ComputerName $servidorOrigen 

$serviceRequestTypeProjectionDestino = Get-SCSMTypeProjection -name System.WorkItem.ServiceRequestProjection$  -ComputerName $servidorDestino

$ActivityTypeProjectionDestino = Get-SCSMTypeProjection -name System.WorkItem.Activity.ManualActivityProjection -ComputerName $servidorOrigen 


 #Get-SCSMTypeProjection  -ComputerName $servidorOrigen | select * | out-gridview
#endregion

#region enumeraciones
$Enums_Servidor_Destino = get-scsmenumeration -ComputerName $servidorDestino
#endregion

#remuevo el log de procesados si existe
if (Test-Path $Registros_procesados_path) {
    $Registros_procesados = Get-Content -Path $Registros_procesados_path
    #Remove-Item $Registros_procesados_path -Force
}


$objSR = get-requerimientos -clase sr -status srNoCompletedCancelClosed -servidor $servidorOrigen 
$objSR = $objSR | Where-Object { $_.id -notin $Registros_procesados } #filtramos aquellos registros que ya fueron copiados correctamente.
#$objSR  =  $objSR | ? {$_.id -eq "SR2121312"} 
#$objSR  =  $objSR | ? {$_.id -eq "SR549291"} 


$SRtotal = $objSR.count

$curent_count = 0
#$objSR | ? {$_.id -eq "SR549291"} | ForEach-Object {
$objSR  | ForEach-Object {
$curent_count+=1
Write-Host "Procesando $($curent_count) / $($SRtotal )"
#region creacion SR
$wi = $_
$wi.id

#obtengo AffectedUser, createdby , AssignedTo y comentarios, me ahorro de traer las relaciones por cada uno de los mencionados     
$serviceRequestProjection = Get-SCSMObjectProjection -ProjectionName $serviceRequestTypeProjectionOrigen.name -filter “ID -eq $($wi.id)” -ComputerName $servidorOrigen 

$RequiredBy  =  Get-SCSMRelatedObject  -Relationship $RequestedByUserRel  -SMObject $wi -ComputerName $servidorOrigen


$SRstatus  = ($Enums_Servidor_Destino  |  ? {$_.name -eq  $wi.Status.name}).displayname

$SRproperties = @{
    Id             = $wi.Id #"SR{0}"
    Title          = $wi.title  
    Description    = $wi.Description
    Urgency        = $wi.Urgency.DisplayName
    priority       = $wi.priority.DisplayName
    Source         = "Portal de autogestión"
    Status         = $SRstatus
    area           = "Pendiente de categorización"
    createdDate =  $wi.CreatedDate
    _Wi = $serviceRequestProjection._WI
  
     
}

  #Es posible que existan SR's Sin grupo de soporte asignado, si existen testeamos que el valor este en el destino y agregamos al diccionario
  if ($wi.SupportGroup.displayname.length -gt 0) {

    $SupportGroup  = ( $Enums_Servidor_Destino  |? {$_.Identifier -match "Trabajo.Solicitudes.Listas.GrupodeSoporte"} | ? {$_.displayname -match $wi.SupportGroup.displayname}  ).name
    $SRproperties["SupportGroup"] =  $SupportGroup
}

#region comentarios
$arrayComentarios = @()
$todosLosComentarios = $serviceRequestProjection.AnalystCommentLog

if ($todosLosComentarios.length -ne 0){

    $todosLosComentarios | ForEach-Object{

    switch ($_.ClassName) {
  
        "System.WorkItem.TroubleTicket.AnalystCommentLog" {$CommentClassName = "AnalystComment"}
        "System.WorkItem.TroubleTicket.UserCommentLog" {$CommentClassName = "EndUserComment"}
    }
         $arrayComentarios += Add-ActionLogEntry -ClassName "System.WorkItem.ServiceRequest" -Action $CommentClassName -Comment $_.comment -EnteredDate  $_.EnteredDate -EnteredBy $_.EnteredBy -IsPrivate $_.IsPrivate -server $servidorDestino

    }

}else{

    $log = "No posee comentarios $($wi.id) con origen en $($servidororigen) -> destino $($servidorDestino) - Requerimiento: $($new_SR.Object.name)"
    write-host $log  -ForegroundColor Yellow
}


$arrayActionLog = @()
$todosLosActionLog = $serviceRequestProjection.ActionLog

if ($todosLosActionLog.length -gt 0){

    $todosLosActionLog | ForEach-Object{

         $arrayActionLog += Add-OperationalActionLogEntry -actionlog $_ -servidor $servidorDestino

    }

}


#     $SRProjectionComment = @{__CLASS = "System.WorkItem.ServiceRequest";
#                  __SEED =   $new_SR.Object
                            
#                  AnalystCommentLog =   $arrayComentarios
              
#                 }
# $new_SRComment = New-SCSMObjectProjection -Type System.WorkItem.ServiceRequestProjection -Projection $SRProjectionComment -PassThru  -ComputerName $servidorDestino # -Credential $cred


#endregion



     $SRProjection = @{__CLASS = "System.WorkItem.ServiceRequest";
                 __OBJECT =   $SRproperties 

                 AffectedUser =  $serviceRequestProjection.AffectedUser
                 CreatedBy = $serviceRequestProjection.CreatedBy
                 AnalystCommentLog =   $arrayComentarios
                 AssignedTo = $serviceRequestProjection.AssignedTo
                 ActionLog = $arrayActionLog
                }
 
try{


$new_SR =  New-SCSMObjectProjection -Type System.WorkItem.ServiceRequestProjection -Projection $SRProjection  -ComputerName $servidorDestino -NoCommit # -Credential $cred 

write-host "se creo el $($new_SR.Object.name)" -ForegroundColor Yellow
}catch{
 $Error[0].Exception 
 $Error[0].CategoryInfo 
 $exception = $Error[0].Exception 



}


#endregion

#reqgion requiredBy
 #New-SCSMRelationshipObject -RelationShip $relAffectedUser -Source $newReq -Target $irAffectedUser -Bulk -ComputerName $servidorDestino

 if ($RequiredBy.length -gt 0){
 #New-SCSMRelationshipObject -RelationShip $RequestedByUserRel -Source $new_SR.Object -Target $RequiredBy -Bulk -ComputerName $servidorDestino
 $new_SR.Add($RequiredBy, $RequestedByUserRel.Target)
 }
#endregion

$new_SR.Commit()


#region Actividades

$ManualActivities = Get-SCSMRelatedObject -Relationship $manualActivitiesRel -SMObject $wi -ComputerName $servidorOrigen
     
$ManualActivities | ForEach-Object {

    $ma = $_

    #$ma | select *

    $MaAssignedToUser = Get-SCSMRelatedObject  -Relationship $AssignedToUserRel -SMObject $ma -ComputerName $servidorOrigen 
    
    $userAnalist = Get-SCSMObject -Class $UserClass -Filter "Username -eq $MaAssignedToUser" -ComputerName $servidorOrigen 


    $MAstatus  = ( $Enums_Servidor_Destino |  ? {$_.name -eq  $ma.Status.name}).displayname


    $ManualActivityProperties = @{
        Id             = $ma.id #"MA{0}"  #$ma.id
        Title          = $ma.Title
        Description    = $ma.Description
        Status         = $MAstatus  
        SequenceId   =  $ma.SequenceId
    }

    #Es posible que existan MA's Sin grupo de soporte asignado, si existen testeamos que el valor este en el destino y agregamos al diccionario
    if ($ma._TierQueue.displayname.length -gt 0) {

        $SupportGroup  = ( $Enums_Servidor_Destino | ? {$_.displayname -match $ma._TierQueue.displayname} | ? {$_.Identifier -match "actividades"} ).name
        $ManualActivityProperties["SupportGroup"] =  $SupportGroup
    }
    



        # Relate the new Manual Activity with the Service Request
  $Projection = @{__CLASS = "System.WorkItem.Activity.ManualActivity";
                    __OBJECT = $ManualActivityProperties
                    
                        
                    ActivityAssignedTo = $MaAssignedToUser;
              
                    ParentWorkItem = $new_SR.Object
                }


#Hago un nuevo objecto de projeccion que automaticamente aplica lo solicit-ado, podria usar -nocommit para uqe sea mas claro la ejecucion. O no.

    $newMa= New-SCSMObjectProjection -Type System.WorkItem.Activity.ManualActivityProjection  -Projection $Projection -ComputerName $servidorDestino  -PassThru
 
    $log = "se creo la $($newMa.Object.DisplayName) -> destino $($servidorDestino) - Requerimiento: $($new_SR.Object.name)"
    write-host $log  -ForegroundColor Yellow


  

            
}


#endregion






#region obtener adjuntos

#descarga los adjuntos en la ruta $basePath con una carpeta con el nombre de $wi.id, ejemplo : c:\temp\ir0001 ,solo si contiene archivos adjuntos

$descarga_realizada = get-AttachReq -wi $wi.id -OutputFolder $basePath -servidor $servidorOrigen 

# #endregion

#region upload archivos Adjuntos
if ($descarga_realizada -eq $true){
   write-host "Preparando para subir archivos" -ForegroundColor Cyan

    $classObj = ($wi.id).substring(0, 2)

    #ruta de la carpeta con nombre del requerimiento
    $FullDirPath = $basePath + $wi.id + "\";
    #obtengo el listado de los archivos descargados
    #$AttachmentEntries = [IO.Directory]::GetFiles($FullDirPath); 

   # $AttachmentArray = $AttachmentEntries.count;
    $existen_archivos = Test-Path ($FullDirPath + "*") 
                    if ( $existen_archivos -eq $true){
              
                    #       foreach($SingleAttachment in $AttachmentEntries) {
                        
                    #                      $AttachmentSingleName = split-path $SingleAttachment -leaf
                     

                    # if ( ( Get-SCSMObject -Class $srClassOrigen -filter "Id -eq $($wi.id)" -ComputerName $servidorOrigen | where {$_.FileAttachment -like $AttachmentSingleName} ) -eq $NULL){

                                Insert-Attachment -SCSMID $new_SR.Object.Name -Directory $FullDirPath -tipoClase $classObj -server $servidorDestino
                                                                                                                    
                                $log = "Subiendo $($AttachmentSingleName) de la carpeta: $($SingleAttachment) -> subido ServiceRequest with ID: $($new_SR.Object.Name) "
                    
                                write-host $log -ForegroundColor Green
                    
                                
                    # }
                                          


                        #   }#finFor
                      
                      }else{
                         $log = "NO hay archivos para adjuntar"
                                                       
                        write-host $log -ForegroundColor Green
                                                       
                        -join($wi.id, "-" ,$log ) | out-file $logPath -Append
                      
                      }#finIF

}
#endregion


#Agrego el WI procesado
Add-Content -Path $Registros_procesados_path -Value $wi.Id
} 


   








$end_runtime = Get-Date


$total_runtime = $end_runtime - $start_runtime

# Display the total run time in seconds
Write-Host "Total run time: $($total_runtime.TotalMinutes) Minutes."
