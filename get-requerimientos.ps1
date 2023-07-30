
function get-requerimientos{
param(
  [Parameter(Mandatory=$true)]
  [String]$servidor,

  [Parameter(Mandatory=$true)]
  [ValidateSet("IR", "SR", "MA")]
  [string[]]
  $clase,

  [Parameter(Mandatory=$true)]
  [ValidateSet("iractive", "irprogress", "irpending", "irpadre","irNoCloseNoResolved","srInProgress", "srNoCompletedCancelClosed")]
  [string[]]
  $status
  

)

#menos cerrados o resuelto

    Import-Module SMLETS
    #$GLOBAL:smdefaultcomputer = $servidor

    Function Crear-Objeto {
  
       Param (
       $obj, 
       $AffectedUser,
       $AssignedToUser,
       $AffectedItems

        )



    $Description = $obj.Description -replace "`r`n", ". " -replace "'", " "

  
    $IRProp = [ordered]@{


    'Id' = $obj.Id;
    'Status'  =$obj.Status.displayname;
    'Title' = $obj.title;
    'Description' = $Description;
    'Classification' =$obj.Classification.displayname;
    'Source' =$obj.source.displayname;
    'Impact' =$obj.Impact.displayname;
    'Urgency' =$obj.Urgency.displayname;
    'priority' =$obj.priority;
    'TierQueue' =$obj.TierQueue.displayName;
    'AssignedToUser' = $AssignedToUser.username; 
    'usuarioAfectado' = $AffectedUser
    '_wi' = $obj._wi;
    'Fecha de creacion'= $obj.CreatedDate;
    <#
    #'ItemsAfectados'=$AffectedItems -join ";" ;
    #'ItemsAfectados2'=$Item 
    'Fecha de creacion'= get-date ($obj.CreatedDate) -format  "dd-MM-yyy";
    #'Tipo de Usuario Afectado'=$AffectedUser.Notes; 
    #'Fecha de creacion'= get-date ($obj.CreatedDate) -format  "dd-MM-yyy";
    #'Fecha de ultima modificacion'=get-date ($obj.LastModified) -format  "dd-MM-yyy";

    #
    #'Resuelto por'=$TroubleTicketResolvedByUser.Username;
    #'Work item'=$obj.WorkItem; #ver
    
    #'fecha_Asignacion' = $fecha;
    #'Servicios afectados'= $Servicio; #ver

    #'ELEMENTOS DE TRABAJO'=$RelatesToWorkItem.id -join  ";";

    #'FechaResolucion' = $FechaResolucion;   

    #'Creado por'=$CreatedByUser.UserName;

    #'Cerrado por'= $CerradoPor.Username;
    #'Categoría de resolucion'=$obj.ResolutionCategory.displayname;
    #'Descripcion de resolucion'=$resolucion;
    #'Site'=$AffectedUser.StreetAddress;
    #'Depto.  De Usuario Afectado'=$AffectedUser.Department; 
    #'Compania de Usuario Afectado'=$AffectedUser.Company;
    #>

    } 


    $IRCustomObjects = New-Object -TypeName Psobject -Property  $IRProp
    
    return $IRCustomObjects



    }


   # $servidor = "s1-dixx-ssm04.ministerio.trabajo.gov.ar"
    
    #region clases


    $IRClass = Get-SCSMClass -Name System.WorkItem.Incident$ -ComputerName $servidor
    $SRClass = Get-SCSMClass -Name System.WorkItem.ServiceRequest$ -ComputerName $servidor
    $MAclass = Get-SCSMClass -Name ClassExtension_5ba907c1_f06b_484f_9c37_7a69eb51f2b8$ -ComputerName $servidor
    $WorkItemContainsActivityRelClass = Get-SCSMRelationshipClass -name System.WorkItemContainsActivity$ -ComputerName $servidor
    #endregion

   

    #region rlaciones
    $WorkItemAffectedUserRel = Get-SCSMRelationshipClass System.WorkItemAffectedUser$ -ComputerName $servidor
    $AssignedToUserRel = Get-scsmrelationshipclass -name System.WorkItemAssignedToUser$ -ComputerName $servidor
    $ImpactsServiceRel = Get-SCSMRelationshipClass System.WorkItemImpactsService  -ComputerName $servidor #######System.WorkItemImpactsService
    $AffectedItemsRel = Get-SCSMRelationshipClass System.WorkItemAboutConfigItem$ -ComputerName $servidor
    $WorkItemHasParentWorkItemRel = Get-SCSMRelationshipClass System.WorkItemHasParentWorkItem$ -ComputerName $servidor
    $WorkItemRelatesToConfigItemRel = Get-SCSMRelationshipClass System.WorkItemRelatesToConfigItem$ -ComputerName $servidor
    $WorkItemRelatesToWorkItemRel =  Get-SCSMRelationshipClass  System.WorkItemRelatesToWorkItem$ -ComputerName $servidor
    $CreatedByUserRel = Get-SCSMRelationshipClass System.WorkItemCreatedByUser$ -ComputerName $servidor
    $TroubleTicketResolvedByUserRel = Get-SCSMRelationshipClass System.WorkItem.TroubleTicketResolvedByUser$ -ComputerName $servidor
    $comentariosRel = Get-SCSMRelationshipClass -Name System.WorkItemHasCommentLog$	-ComputerName $servidor

    $ConfigItemOwnedByUserRel = Get-SCSMRelationshipClass -name "System.ConfigItemOwnedByUser$" -ComputerName $servidor
    $relAffectedUser = Get-SCSMRelationshipClass -Name System.WorkItemAffectedUser -ComputerName $servidor
    $UserClass = Get-SCSMClass -name System.Domain.User$ -ComputerName $servidor # Get SCSM User class object

    #relaciones de SR
    $RequestedByUserRel = Get-SCSMRelationshipClass -name System.WorkItemRequestedByUser$ -ComputerName $servidor
    $ClosedByUserRel = Get-SCSMRelationshipClass  -name System.WorkItemClosedByUser$ -ComputerName $servidor
    $OfertasRel = Get-SCSMRelationshipClass -Name System.WorkItemRelatesToRequestOffering$ -ComputerName $servidor
    $WorkItemContainsActivityRel = Get-scsmrelationshipclass -name System.WorkItemContainsActivity  -ComputerName $servidor

    #relaciones de MA 

    $CreatedByUserRel = Get-SCSMRelationshipClass System.WorkItemCreatedByUser$ -ComputerName $servidor


    #endregion
    
    $IrActive = (Get-SCSMEnumeration  -Name IncidentStatusEnum.Active$ -ComputerName $servidor).id
    $IrProgress = (Get-SCSMEnumeration -Name Enum.58ab29d637184e3989cff0092999d468 -ComputerName $servidor).id 
    $IrPending = (Get-SCSMEnumeration -Name IncidentStatusEnum.Active.Pending -ComputerName $servidor).id 
    $IrPadre = (Get-SCSMEnumeration -Name Enum.b8704074109441a48140755cc165c7a1 -ComputerName $servidor).id;
    $Irclose = (Get-SCSMEnumeration -Name IncidentStatusEnum.Closed$ -ComputerName $servidor).id;
    $IrResolved = (Get-SCSMEnumeration -Name IncidentStatusEnum.Resolved$ -ComputerName $servidor).id;



  
    $SRNew =(Get-SCSMEnumeration -Name ServiceRequestStatusEnum.New -ComputerName $servidor).id
    $SRClosed =(Get-SCSMEnumeration -Name ServiceRequestStatusEnum.Closed -ComputerName $servidor).id
    $SRCompleted =(Get-SCSMEnumeration -Name ServiceRequestStatusEnum.Completed -ComputerName $servidor).id
    $SRFailed =(Get-SCSMEnumeration -Name ServiceRequestStatusEnum.Failed -ComputerName $servidor).id
    $SRCanceled =(Get-SCSMEnumeration -Name ServiceRequestStatusEnum.Canceled -ComputerName $servidor).id
    $SROnHold =(Get-SCSMEnumeration -Name ServiceRequestStatusEnum.OnHold -ComputerName $servidor).id
    $SRInProgress =(Get-SCSMEnumeration -Name ServiceRequestStatusEnum.InProgress -ComputerName $servidor).id


<#
#-----------test----borrar---->
   # Get-SCSMEnumeration | Out-GridView

        $clase = "sr"

        $class = switch ($clase)
    {
        IR {$IRClass}
        SR {$SRClass}
        MA {$MAclass}
       
    }



$status = "srNoCompletedCancelClosed"

#-----------test--borrar------>

    #$Username = "meaguirre"
    #$UserClass = Get-SCSMClass -name System.Domain.User$ # Get SCSM User class object
    #$userAnalist = Get-SCSMObject -Class $UserClass -Filter "Username -eq $Username"

#>  



    $estado = switch ($status)
    {
        iractive { $filtro = "Status -eq '{0}'" -f $IrActive}
        irprogress {$filtro = "Status -eq '{0}'" -f $IrProgress}
        irpending {$filtro = "Status -eq '{0}'" -f $IrPending}
        irpadre { $filtro = "Status -eq '{0}'" -f $IrPadre}
        srInProgress {$filtro = "Status -eq '{0}'" -f $SRInProgress}
        irNoCloseNoResolved{ $filtro = "Status -ne '{0}' -and Status -ne '{1}'" -f $IrResolved, $Irclose }
        srNoCompletedCancelClosed{ $filtro = "Status -ne '{0}' -and Status -ne '{1}'-and Status -ne '{2}'" -f $SRCompleted, $SRCanceled, $SRClosed}
    }

 
    #sr menos completado, cancelado, cerrado

     $class = switch ($clase)
    {
        IR {$IRClass}
        SR {$SRClass}
        MA {$MAclass}
       
    }
 

    $requerimiento = Get-SCSMObject -Class $class -Filter $filtro -ComputerName $servidor
   
  <#
    $obj = @()

    #$IRObjects | ? {$_.id -eq "IR2062995"}| sort-object -Property Id | ForEach-Object -Process {

    $requerimiento | sort-object -Property Id | ForEach-Object -Process {

    #$TroubleTicketResolvedByUser = Get-SCSMRelatedObject -SMObject $_ -Relationship $TroubleTicketResolvedByUserRel

    $AffectedItems = Get-SCSMRelatedObject  -Relationship $AffectedItemsRel -SMObject $_ -ComputerName $servidor

    $AffectedUser = Get-SCSMRelatedObject  -Relationship $WorkItemAffectedUserRel -SMObject $_ -ComputerName $servidor

    $AssignedToUser = Get-SCSMRelatedObject  -Relationship $AssignedToUserRel -SMObject $_ -ComputerName $servidor

    #$ConfigItemOwnedByUser = Get-SCSMRelatedObject  -Relationship $ConfigItemOwnedByUserRel -SMObject $_

    # $fecha  =  (get-HistoryUserassign $_ $AssignedToUser).fechaAsig

    $obj += Crear-Objeto $_ $AffectedUser $AssignedToUser $AffectedItems

    }#fin for-object

    #$objIR  | Export-Clixml -Path c:\temp\Incidentes_Test.xml


return $obj
#>

 return $requerimiento

}


#get-requrimientos -clase IR -servidor s1-dixx-ssm04 -status irNoCloseNoResolved