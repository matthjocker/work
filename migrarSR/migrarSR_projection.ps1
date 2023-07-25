# https://www.stefanroth.net/2014/09/01/scsm-adding-activities-using-sma-powershell-workflow/

# https://community.cireson.com/discussion/3486/can-you-create-a-service-request-from-another-service-request-activity


import-module SMlets

[Threading.Thread]::CurrentThread.CurrentCulture = 'es-ES'


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

. $pathFunciones\get-actionLogFullv2.ps1

. $pathFunciones\add-actionLogEntryV2.ps1

. $pathFunciones\get-AttachReqV2.ps1


. $pathFunciones\UploadAttachReqv2.ps1

#endregion

#region constantes

$servidorOrigen = "scsm.ministerio.trabajo.gov.ar"
$servidorDestino = "s1-dixx-ssm04"
#$servidorDestino = "s1-hixx-ssm01"

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

#region relaciones

$AssignedToRel = get-scsmrelationshipclass -name System.WorkItemAssignedToUser$ -ComputerName $servidorOrigen
$WorkItemAffectedUserRel = Get-SCSMRelationshipClass System.WorkItemAffectedUser$ -ComputerName $servidorOrigen
$AssignedToUserRel = Get-scsmrelationshipclass -name System.WorkItemAssignedToUser$ -ComputerName $servidorOrigen
$WorkItemAffectedUserRel = Get-SCSMRelationshipClass System.WorkItemAffectedUser$ -ComputerName $servidorOrigen

$WorkItemContainsActivityRelOrigen =Get-scsmrelationshipclass -name System.WorkItemContainsActivity$  -ComputerName $servidorOrigen
$WorkItemContainsActivityRelDestino = Get-SCSMRelationshipClass -Name "System.WorkItemContainsActivity" -ComputerName $servidorDestino
$createdByRelClass = Get-SCSMRelationshipClass -Name System.WorkItemCreatedByUser$  -ComputerName $servidorOrigen

#endregion
#endregion



$objSR = get-requrimientos -clase sr -status srNoCompletedCancelClosed -servidor $servidorOrigen 

$objSR.count

$serviceRequestTypeProjectionOrigen = Get-SCSMTypeProjection -name System.WorkItem.ServiceRequestProjection$  -ComputerName $servidorOrigen 

$serviceRequestTypeProjectionDestino = Get-SCSMTypeProjection -name System.WorkItem.ServiceRequestProjection$  -ComputerName $servidorDestino

#$serviceRequestTypeProjectionDestino | ? {$_.name -eq "AnalystCommentLog"}


$objSR| ForEach-Object {

$wi = $_

$wi.id

$serviceRequestProjection = Get-SCSMObjectProjection -ProjectionName $serviceRequestTypeProjectionOrigen.name -filter “ID -eq $($wi.id)” -ComputerName $servidorOrigen 

# $serviceRequestProjection.AnalystCommentLog.values usercomment | select * ActionLog | gm

# $wi | select *

$AffectedUser = Get-SCSMRelatedObject  -Relationship $WorkItemAffectedUserRel -SMObject $_ -ComputerName $servidorOrigen

$AssignedToUser = Get-SCSMRelatedObject  -Relationship $AssignedToUserRel -SMObject $wi -ComputerName $servidorOrigen

$createdBy = Get-SCSMRelatedObject  -Relationship $createdByRelClass -SMObject $_ -ComputerName $servidorOrigen

#$username =  $AffectedUser.UserName 

#$analist = $AssignedToUser.DisplayName

$Username = "SCSM_Usuario_Prueba"

$AffectedUser = Get-SCSMObject -Class $UserClass -Filter "Username -eq $username" -ComputerName $servidorOrigen 

$userAnalist = Get-SCSMObject -Class $UserClass -Filter "Username -eq $analist" -ComputerName $servidorOrigen 

Get-SCSMEnumeration -ComputerName $servidorOrigen  $wi | select * | Out-GridView

$SupportGroup  = ( get-scsmenumeration -ComputerName $servidorDestino| ? {$_.displayname -match $wi.SupportGroup.displayname } | ? {$_.Identifier -match "solicitudes"}).name

$clasificacion = ( get-scsmenumeration -ComputerName $servidorDestino |  ? {$_.displayname -eq  "Pendiente de categorización"} | ? {$_.Identifier -match "Trabajo.Lista.AreaSolictudServicio"}).name

$SRstatus  = ( get-scsmenumeration -ComputerName $servidorDestino |  ? {$_.name -eq  $wi.Status.name}).displayname

#$wi | select * | Out-GridView

#get-scsmenumeration -ComputerName $servidorDestino | select * | Out-GridView

# Prepare Incident properties
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

}
# Create 



 $SRProjection = @{__CLASS = "System.WorkItem.ServiceRequest";
                 __OBJECT =   $SRproperties 

                }

 
#Creamos la proyección utilizando el cmdlet New-SCSMObjectProjection

$new_SR = New-SCSMObjectProjection -Type System.WorkItem.ServiceRequestProjection -Projection $SRProjection -PassThru  -ComputerName $servidorDestino  #-Credential $cred

write-host "se creo el $($new_SR.Object)" -ForegroundColor Yellow

#region comentarios

$serviceRequestProjection.AnalystCommentLog | ForEach-Object{

  switch ($_.ClassName)
        {
  
            "System.WorkItem.TroubleTicket.AnalystCommentLog" {$CommentClassName = "AnalystComment"}
           "System.WorkItem.TroubleTicket.UserCommentLog" {$CommentClassName = "EndUserComment"}
        }

Add-ActionLogEntry -WIObject $wi -Action $CommentClassName -Comment $_.comment -EnteredDate $_.EnteredDate -EnteredBy $_.EnteredBy -IsPrivate $_.IsPrivate -server $servidorDestino

}
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
        Id             = "MA{0}"
        Title          = $ma.Title
        Description    = $ma.Description
        Status         = $MAstatus   # Puedes establecer el estado que desees para la actividad
        AssignedTo     = $userAnalist.DisplayName # Nombre del usuario o grupo al que se asigna la actividad
        SupportGroup   = $SupportGroup # Nombre del grupo de soporte para la actividad
        SequenceId   =  $ma.SequenceId   
    }


       # $newMA = New-SCSMObject -Class $MaClassDestino -PropertyHashtable $ManualActivityProperties -PassThru  -ComputerName $servidorDestino -NoCommit

        # Relate the new Manual Activity with the Service Request
  $Projection = @{__CLASS = "System.WorkItem.Activity.ManualActivity";
                __OBJECT =   $ManualActivityProperties 

                ActivityAssignedTo = $MaAssignedToUser;
              
                ParentWorkItem = $new_SR.Object
                }


#Hago un nuevo objecto de projeccion que automaticamente aplica lo solicit-ado, podria usar -nocommit para uqe sea mas claro la ejecucion. O no.

    $project_test = New-SCSMObjectProjection -Type System.WorkItem.Activity.ManualActivityProjection  -Projection $Projection -ComputerName $servidorDestino -PassThru
 
    write-host "se creo la $($project_test.Object.DisplayName)" -ForegroundColor Yellow

}
#endregion

} 


   









