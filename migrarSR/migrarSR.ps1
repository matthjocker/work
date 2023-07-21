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



#region relaciones

$AssignedToRel = get-scsmrelationshipclass -name System.WorkItemAssignedToUser$ -ComputerName $servidorOrigen
$WorkItemAffectedUserRel = Get-SCSMRelationshipClass System.WorkItemAffectedUser$ -ComputerName $servidorOrigen
$AssignedToUserRel = Get-scsmrelationshipclass -name System.WorkItemAssignedToUser$ -ComputerName $servidorOrigen
$WorkItemAffectedUserRel = Get-SCSMRelationshipClass System.WorkItemAffectedUser$ -ComputerName $servidorOrigen

$manualActivitiesRel =Get-scsmrelationshipclass -name System.WorkItemContainsActivity$  -ComputerName $servidorOrigen


#endregion = 
#endregion







#get-scsmenumeration -ComputerName $servidorDestino | select * | Out-GridView    ? {$_.displayname -match $wi.status} 



#revisar: cambiar "IR{0}" por $wi.id




$objSR = get-requrimientos -clase sr -status srNoCompletedCancelClosed -servidor $servidorOrigen 

$objSR.count

$objSR | select -First 1| ForEach-Object {

$wi = $_






# Obtener las actividades manuales asociadas a la solicitud
   
   



# $wi | select *

$AffectedUser = Get-SCSMRelatedObject  -Relationship $WorkItemAffectedUserRel -SMObject $_ -ComputerName $servidorOrigen

$AssignedToUser = Get-SCSMRelatedObject  -Relationship $AssignedToUserRel -SMObject $_ -ComputerName $servidorOrigen

#$username =  $AffectedUser.UserName 

#$analist = $AssignedToUser.DisplayName

$Username = "SCSM_Usuario_Prueba"

$AffectedUser = Get-SCSMObject -Class $UserClass -Filter "Username -eq $username" -ComputerName $servidorOrigen 

$userAnalist = Get-SCSMObject -Class $UserClass -Filter "Username -eq $analist" -ComputerName $servidorOrigen 


$SupportGroup  = ( get-scsmenumeration -ComputerName $servidorDestino| ? {$_.displayname -match $wi.SupportGroup.displayname } | ? {$_.Identifier -match "solicitudes"}).name



$clasificacion = ( get-scsmenumeration -ComputerName $servidorDestino |  ? {$_.displayname -eq  "Pendiente de categorización"} | ? {$_.Identifier -match "Trabajo.Lista.AreaSolictudServicio"}).name

$SRstatus  = ( get-scsmenumeration -ComputerName $servidorDestino |  ? {$_.name -eq  $wi.Status.name}).displayname




if ($wi.Urgency.DisplayName -eq "No Aplica"){

$Urgency = "Baja"

}else{

$Urgency  = $wi.Urgency.DisplayName
}

$wi | select * | Out-GridView

#get-scsmenumeration -ComputerName $servidorDestino | select * | Out-GridView

# Prepare Incident properties
$SRproperties = @{
    Id             = "SR{0}"
    Title          = $wi.title
    Description    = $wi.Description
    Urgency        = $Urgency 
    priority       = $wi.priority.DisplayName
    Source         = "Portal de autogestión"
    Status         = $SRstatus
  # Classification = "Pendiente de categorización"
    SupportGroup   = $SupportGroup
    area           = "Pendiente de categorización"
}
# Create Incident object

 $newSR = New-SCSMObject -Class $SRclassDestino -PropertyHashtable $SRproperties -PassThru  -ComputerName $servidorDestino #-Credential $cred
 $newSR.DisplayName

 
 #crear nueva relacioncon el nuevo sr

$ManualActivities = Get-SCSMRelatedObject -Relationship $manualActivitiesRel -SMObject $wi -ComputerName $servidorOrigen
     
$ManualActivities[1] | ForEach-Object {

    $ma = $_

    $MaAssignedToUser = (Get-SCSMRelatedObject  -Relationship $AssignedToUserRel -SMObject $ma -ComputerName $servidorOrigen).username
    
   
    
    $userAnalist = Get-SCSMObject -Class $UserClass -Filter "Username -eq $MaAssignedToUser" -ComputerName $servidorOrigen 

    $SupportGroup  = ( get-scsmenumeration -ComputerName $servidorDestino| ? {$_.displayname -match $ma._TierQueue.displayname} | ? {$_.Identifier -match "actividades"}).name

    $MAstatus  = ( get-scsmenumeration -ComputerName $servidorDestino |  ? {$_.name -eq  $ma.Status.name}).displayname
     

     $ManualActivityProperties = @{
        Id             = "MA{0}"
        Title          = $ma.Title
        Description    = $ma.Description
        Status         = $MAstatus   # Puedes establecer el estado que desees para la actividad
        #AssignedTo     = $userAnalist.DisplayName # Nombre del usuario o grupo al que se asigna la actividad
        SupportGroup   = $SupportGroup # Nombre del grupo de soporte para la actividad
    }

    # Crear la nueva Manual Activity
    $newMA = New-SCSMObject -Class $MaClassDestino -PropertyHashtable $ManualActivityProperties -PassThru  -ComputerName $servidorDestino #-Credential $cred


    $existingRelationships = Get-SCSMRelationshipObject -ByTarget $newMA   -ComputerName $servidorDestino   -Relationship (Get-SCSMRelationshipClass -Name System.WorkItemRelatesToActivity$ -ComputerName $servidorDestino)



    if ($newMA -and $newSR) {  

       New-SCSMRelationshipObject -Relationship $manualActivitiesRel -Source $newSR -Target $newMA - -ComputerName $servidorDestino 


    } 

   
}





# Set Affected User
if ($AffectedUser -and $newSR) {

    #New-SCSMRelationshipObject -RelationShip $relAffectedUser -Source $newSR -Target $AffectedUser -Bulk  -ComputerName $servidorDestino #-Credential $cred
}


if ($maAssignedToUserObj -and $newMA) {

    New-SCSMRelationshipObject -RelationShip $AssignedToUserRel -Source $newMA -Target $userAnalist -Bulk
}




}