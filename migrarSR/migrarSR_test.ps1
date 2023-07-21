import-module SMlets

#region importar funciones

$pathFunciones = "E:\trabajo\migrarIR\"

. $pathFunciones\get-requerimientos.ps1


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
#endregion

#region relaciones

$AssignedToRel = get-scsmrelationshipclass -name System.WorkItemAssignedToUser$ -ComputerName $servidorOrigen
$WorkItemAffectedUserRel = Get-SCSMRelationshipClass System.WorkItemAffectedUser$ -ComputerName $servidorOrigen
$AssignedToUserRel = Get-scsmrelationshipclass -name System.WorkItemAssignedToUser$ -ComputerName $servidorOrigen
$WorkItemAffectedUserRel = Get-SCSMRelationshipClass System.WorkItemAffectedUser$ -ComputerName $servidorOrigen

$manualActivitiesRel =Get-scsmrelationshipclass -name System.WorkItemContainsActivity$  -ComputerName $servidorOrigen


#endregion

$objSR = get-requrimientos -clase sr -status srNoCompletedCancelClosed -servidor $servidorOrigen 

$objSR | select -First 1| ForEach-Object {
#region crear SR

$wi = $_

$AffectedUser = Get-SCSMRelatedObject  -Relationship $WorkItemAffectedUserRel -SMObject $_ -ComputerName $servidorOrigen

$AssignedToUser = Get-SCSMRelatedObject  -Relationship $AssignedToUserRel -SMObject $_ -ComputerName $servidorOrigen

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



$SRproperties = @{
    Id             = "SR{0}"
    Title          = $wi.title
    Description    = $wi.Description
    Urgency        = $Urgency 
    priority       = $wi.priority.DisplayName
    Source         = "Portal de autogestión"
    Status         = $SRstatus
 
    SupportGroup   = $SupportGroup
    area           = "Pendiente de categorización"
}
# crear SR

 $newSR = New-SCSMObject -Class $SRclassDestino -PropertyHashtable $SRproperties -PassThru  -ComputerName $servidorDestino #-Credential $cred
 $newSR.DisplayName
#endregion

  


#region Obtener Ma de la SR origen

$ManualActivities = Get-SCSMRelatedObject -Relationship $manualActivitiesRel -SMObject $wi -ComputerName $servidorOrigen
     
$ManualActivities | ForEach-Object {

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

    # Crea la nueva Manual Activity
 
     $newMA = New-SCSMObject -Class $MaClassDestino -PropertyHashtable $MAproperties -PassThru -ComputerName $servidorDestino

    # Relate the new Manual Activity with the Service Request
    $relationshipClass = Get-SCSMRelationshipClass -Name "System.WorkItemContainsActivity" -ComputerName $servidorDestino
    New-SCSMRelationshipObject -Relationship $relationshipClass -Source $newSR -Target $newMA  -ComputerName $servidorDestino
    
   
}

#endregion


#region ayuda
        $sr =  Get-SCSMObject -Class $SRclassDestino -Filter "id -eq $newSR" -ComputerName $servidorDestino
        $ma = Get-SCSMObject -Class $MaClassDestino -Filter "id -eq $newMA" -ComputerName $servidorDestino

       # Get-SCSMRelationshipClass -ComputerName $servidorDestino | select * | Out-GridView

       $existingRelationship = Get-SCSMRelationshipObject -Relationship $relationshipClass -Source $sr  -Target $ma -ComputerName $servidorDestino
     
       $existingRelationship =  Get-SCSMRelationshipObject -BySource $sr -ComputerName $servidorDestino
#endregion

}