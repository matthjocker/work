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
#region importar funciones
$pathFunciones = "E:\trabajo\migrarIR\"
. $pathFunciones\get-requerimientos.ps1
. $pathFunciones\add-actionLogEntryV2.ps1
. $pathFunciones\get-AttachReqV2.ps1
. $pathFunciones\UploadAttachReqv2.ps1
#endregion
#region constantes
$servidorOrigen = "scsm.ministerio.trabajo.gov.ar"
$servidorDestino = "s1-dixx-ssm04"
#$servidorDestino = "s1-hixx-ssm01"
$basePath = "C:\temp\reqExport\"
$logPath ="E:\trabajo\migrarSR\logs\logs_migracionSR.txt"
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
$AssignedToUserRel = Get-scsmrelationshipclass -name System.WorkItemAssignedToUser$ -ComputerName $servidorOrigen
$WorkItemAffectedUserRel = Get-SCSMRelationshipClass System.WorkItemAffectedUser$ -ComputerName $servidorOrigen
$WorkItemContainsActivityRelOrigen = Get-scsmrelationshipclass -name System.WorkItemContainsActivity$  -ComputerName $servidorOrigen
$WorkItemContainsActivityRelDestino = Get-SCSMRelationshipClass -Name System.WorkItemContainsActivity$ -ComputerName $servidorDestino
$manualActivitiesRel = Get-SCSMRelationshipClass -Name System.WorkItemContainsActivity$  -ComputerName $servidorOrigen
#Get-SCSMRelationshipClass  -ComputerName $servidorOrigen   | select * | Out-GridView

#endregion
#region projection

$serviceRequestTypeProjectionOrigen = Get-SCSMTypeProjection -name System.WorkItem.ServiceRequestProjection$  -ComputerName $servidorOrigen 

$serviceRequestTypeProjectionDestino = Get-SCSMTypeProjection -name System.WorkItem.ServiceRequestProjection$  -ComputerName $servidorDestino

$ActivityTypeProjectionDestino = Get-SCSMTypeProjection -name System.WorkItem.Activity.ManualActivityProjection -ComputerName $servidorOrigen 

#endregion

$objSR = get-requerimientos -clase sr -status srNoCompletedCancelClosed -servidor $servidorOrigen 
$objSR.count

$objSR | ? {$_.id -eq "SR549291"} | ForEach-Object {
#region creacion SR
$wi = $_
$wi.id

#obtengo AffectedUser, createdby , AssignedTo y comentarios, me ahorro de traer las relaciones por cada uno de los mencionados     
$serviceRequestProjection = Get-SCSMObjectProjection -ProjectionName $serviceRequestTypeProjectionOrigen.name -filter “ID -eq $($wi.id)” -ComputerName $servidorOrigen 

$SupportGroup  = ( get-scsmenumeration -ComputerName $servidorDestino |? {$_.Identifier -match "Trabajo.Solicitudes.Listas.GrupodeSoporte"} | ? {$_.displayname -match $wi.SupportGroup.displayname}  ).name

$SRstatus  = ( get-scsmenumeration -ComputerName $servidorDestino |  ? {$_.name -eq  $wi.Status.name}).displayname

$SRproperties = @{
    Id             = "SR{0}"
    Title          = $wi.title  
    Description    = $wi.Description
    Urgency        = $wi.Urgency.DisplayName
    priority       = $wi.priority.DisplayName
    Source         = "Portal de autogestión"
    Status         = $SRstatus
    SupportGroup   = $SupportGroup
    area           = "Pendiente de categorización"
    createdDate = $wi.CreatedDate
    _Wi = $serviceRequestProjection._WI
     
}
     $SRProjection = @{__CLASS = "System.WorkItem.ServiceRequest";
                 __OBJECT =   $SRproperties 

                 #AffectedUser =  $serviceRequestProjection.AffectedUser
                 #CreatedBy = $serviceRequestProjection.CreatedBy
                # AnalystCommentLog =   $arrayComentarios
                # AssignedTo = $serviceRequestProjection.AssignedTo
                }
 
try{

$new_SR = New-SCSMObjectProjection -Type System.WorkItem.ServiceRequestProjection -Projection $SRProjection -PassThru  -ComputerName $servidorDestino # -Credential $cred
write-host "se creo el $($new_SR.Object.name)" -ForegroundColor Yellow
}catch{
 $Error[0].Exception 
 $Error[0].CategoryInfo 
 $exception = $Error[0].Exception 



}


#endregion

#region comentarios
$arrayComentarios = @()
$todosLosComentarios = $serviceRequestProjection.AnalystCommentLog

if ($todosLosComentarios.length -ne 0){

    $todosLosComentarios | ForEach-Object{

    switch ($_.ClassName)
    {
  
        "System.WorkItem.TroubleTicket.AnalystCommentLog" {$CommentClassName = "AnalystComment"}
        "System.WorkItem.TroubleTicket.UserCommentLog" {$CommentClassName = "EndUserComment"}
    }
         $arrayComentarios += Add-ActionLogEntry -WIObject $new_SR.Object -Action $CommentClassName -Comment $_.comment -EnteredDate $_.EnteredDate -EnteredBy $_.EnteredBy -IsPrivate $_.IsPrivate -server $servidorDestino

    }

}else{

    $log = "No posee comentarios $($wi.id) con origen en $($servidororigen) -> destino $($servidorDestino) - Requerimiento: $($new_SR.Object.name)"
    write-host $log  -ForegroundColor Yellow


}



    $SRProjectionComment = @{__CLASS = "System.WorkItem.ServiceRequest";
                 __SEED =   $new_SR.Object
                            
                 AnalystCommentLog =   $arrayComentarios
              
                }
$new_SRComment = New-SCSMObjectProjection -Type System.WorkItem.ServiceRequestProjection -Projection $SRProjectionComment -PassThru  -ComputerName $servidorDestino # -Credential $cred


#endregion

#region obtener adjuntos

#descarga los adjuntos en la ruta $basePath con una carpeta con el nombre de $wi.id, ejemplo : c:\temp\ir0001 ,solo si contiene archivos adjuntos

get-AttachReq -wi $wi.id -OutputFolder $basePath -servidor $servidorOrigen 

#endregion

#region upload archivos Adjuntos

   write-host "Preparando para subir archivos" -ForegroundColor Cyan

    $classObj = ($wi.id).substring(0, 2)

    #ruta de la carpeta con nombre del requerimiento
    $FullDirPath = $basePath + $wi.id + "\";
    #obtengo el listado de los archivos descargados
    $AttachmentEntries = [IO.Directory]::GetFiles($FullDirPath); 

    $AttachmentArray = $AttachmentEntries.count;

                     if ($AttachmentArray -ne $NULL){
              
                          foreach($SingleAttachment in $AttachmentEntries) {
                        
                                         $AttachmentSingleName = split-path $SingleAttachment -leaf
                     

                                                        if ( ( Get-SCSMObject -Class $srClassOrigen -filter "Id -eq $($wi.id)" -ComputerName $servidorOrigen | where {$_.FileAttachment -like $AttachmentSingleName} ) -eq $NULL){

                                                                  Insert-Attachment -SCSMID $new_SR.Object.Name -Directory $SingleAttachment -tipoClase $classObj -server $servidorDestino
                                                                                                                                                     
                                                                  $log = "Subiendo $($AttachmentSingleName) de la carpeta: $($SingleAttachment) -> subido ServiceRequest with ID: $($new_SR.Object.Name) "
                                                       
                                                                  write-host $log -ForegroundColor Green
                                                       
                                                                  
                                                        }
                                          


                          }#finFor
                      
                      }else{
                         $log = "NO hay archivos para adjuntar"
                                                       
                        write-host $log -ForegroundColor Green
                                                       
                        -join($wi.id, "-" ,$log ) | out-file $logPath -Append
                      
                      }#finIF


#endregion

#region Actividades

$ManualActivities = Get-SCSMRelatedObject -Relationship $manualActivitiesRel -SMObject $wi -ComputerName $servidorOrigen
     
$ManualActivities | ForEach-Object {

    $ma = $_

    #$ma | select *

    $MaAssignedToUser = ( Get-SCSMRelatedObject  -Relationship $AssignedToUserRel -SMObject $ma -ComputerName $servidorOrigen  ).username
    
    $userAnalist = Get-SCSMObject -Class $UserClass -Filter "Username -eq $MaAssignedToUser" -ComputerName $servidorOrigen 

    $SupportGroup  = ( get-scsmenumeration -ComputerName $servidorDestino| ? {$_.displayname -match $ma._TierQueue.displayname} | ? {$_.Identifier -match "actividades"}).name

    $MAstatus  = ( get-scsmenumeration -ComputerName $servidorDestino |  ? {$_.name -eq  $ma.Status.name}).displayname


    $ManualActivityProperties = @{
        Id             = "MA{0}"  #$ma.id
        Title          = $ma.Title
        Description    = $ma.Description
        Status         = $MAstatus  
        #AssignedTo     = $userAnalist.DisplayName 
        SupportGroup   = $SupportGroup
        SequenceId   =  $ma.SequenceId
    }




        # Relate the new Manual Activity with the Service Request
  $Projection = @{__CLASS = "System.WorkItem.Activity.ManualActivity";
                    __OBJECT = $ManualActivityProperties
                    
                        
                    ActivityAssignedTo = $MaAssignedToUser;
              
                    ParentWorkItem = $new_SR.Object
                }


#Hago un nuevo objecto de projeccion que automaticamente aplica lo solicit-ado, podria usar -nocommit para uqe sea mas claro la ejecucion. O no.

    $newMa= New-SCSMObjectProjection -Type System.WorkItem.Activity.ManualActivityProjection  -Projection $Projection -ComputerName $servidorDestino -PassThru
 
    $log = "se creo la $($newMa.Object.DisplayName) -> destino $($servidorDestino) - Requerimiento: $($new_SR.Object.name)"
    write-host $log  -ForegroundColor Yellow



}


#endregion

} 


   









