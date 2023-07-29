
# param(
#     [parameter(Position=0,Mandatory=$false)][boolean]$BeQuiet=$true,
#     [parameter(Position=1,Mandatory=$false)][string]$URL  
# )
# Then call the import-module cmdlet like this:

# import-module .\myModule.psm1 -ArgumentList $True,'http://www.microsoft.com'


Import-Module smlets

    #region Funciones genericas
  
    function get-MigracionSCSM_Requerimientos{
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
            
            #region clases          
            $IRClass = Get-SCSMClass -Name System.WorkItem.Incident$ -ComputerName $servidor
            $SRClass = Get-SCSMClass -Name System.WorkItem.ServiceRequest$ -ComputerName $servidor
            $MAclass = Get-SCSMClass -Name ClassExtension_5ba907c1_f06b_484f_9c37_7a69eb51f2b8$ -ComputerName $servidor
            $WorkItemContainsActivityRelClass = Get-SCSMRelationshipClass -name System.WorkItemContainsActivity$ -ComputerName $servidor
            #endregion
        
           
        
            #region rlaciones
            $WorkItemAffectedUserRel = Get-SCSMRelationshipClass System.WorkItemAffectedUser$ -ComputerName $servidor
            $AssignedToUserRel = Get-scsmrelationshipclass -name System.WorkItemAssignedToUser$ -ComputerName $servidor
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
            $WorkItemContainsActivityRel =Get-scsmrelationshipclass -name System.WorkItemContainsActivity -ComputerName $servidor
        
            #relaciones de MA 
        
            $CreatedByUserRel = Get-SCSMRelationshipClass System.WorkItemCreatedByUser$ -ComputerName $servidor
        
        
            #endregion
            
            #region enumeraciones
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
            #endregion
        
            $filtro = switch ($status)
            {
                iractive { "Status -eq '{0}'" -f $IrActive ; break}
                irprogress {"Status -eq '{0}'" -f $IrProgress ; break}
                irpending { "Status -eq '{0}'" -f $IrPending ; break}
                irpadre { "Status -eq '{0}'" -f $IrPadre ; break}
                srInProgress {"Status -eq '{0}'" -f $SRInProgress ; break}
                irNoCloseNoResolved{  "Status -ne '{0}' -and Status -ne '{1}'" -f $IrResolved, $Irclose  ; break}
                srNoCompletedCancelClosed{ "Status -ne '{0}' -and Status -ne '{1}'-and Status -ne '{2}'" -f $SRCompleted, $SRCanceled, $SRClosed ; break}
            }
               
            #sr menos completado, cancelado, cerrado
        
             $class = switch ($clase)
            {
                IR {$IRClass}
                SR {$SRClass}
                MA {$MAclass}
               
            }
         
        
            $requerimiento = Get-SCSMObject -Class $class -Filter $filtro -ComputerName $servidor
                
         return $requerimiento
        
        }
        

    function get-MigracionSCSM_AttachReq {
        param(
            [Parameter(Mandatory=$true)]
            [String]$servidor,
        
            [Parameter(Mandatory=$true)]
            [string[]]
            $wi,
        
            [Parameter(Mandatory=$true)]
            [string]$OutputFolder
            
        
        )
        
        
        $srClass = Get-SCSMClass -Name system.workitem.servicerequest$ -ComputerName $servidor   
        $irClass = Get-SCSMClass -name System.Workitem.Incident$ -ComputerName $servidor         
        $AttachedFileRel = get-scsmrelationshipclass System.WorkItemHasFileAttachment$ -ComputerName $servidor
        
        
        $id = $wi.Trim()
        
        $classObj = $wi.substring(0, 2)
        
            switch ( $classObj )
                    {           
                        sr {  $requerimiento = Get-SCSMObject -Class $srClass -filter "id -eq $id" -ComputerName $servidor}
                        ir {  $requerimiento = Get-SCSMObject -Class $irClass -filter "id -eq $id" -ComputerName $servidor }
            
                    }
        
        #get file relations, then return a collection of files. 
        $archivos_Adjuntos = ( Get-SCSMRelationshipObject -Bysource $requerimiento -ComputerName $servidor| ? {$_.relationshipid -eq $AttachedFileRel.id -and $_.IsDeleted -eq $false} | % {$_.TargetObject})
        
        
        if($archivos_Adjuntos){           
            #crea una carpeta con el nuemero de requerimiento
                $finalFolder = -join($OutputFolder + $id)          
                write-host $finalFolder -ForegroundColor Magenta      
            if (Test-Path -Path $finalFolder) {
                Write-Host "Ya existe carpeta" -ForegroundColor Yellow
        } else {
                New-Item -Path $finalFolder -ItemType Directory
                    Write-Host 'carpeta creada:'  $finalFolder -ForegroundColor Yellow
        }

        
        foreach ($File in $archivos_Adjuntos) {
            Try {
                #File byte buffer
                $buffer = new-object byte[] -ArgumentList 4096
            
                $archivo = ($File.Values | ? {$_.type -match "DisplayName"}).value 
            
                $Content =  ($File.Values | ? {$_.type -match "Content"}).value # as Microsoft.EnterpriseManagement.Common.ServerBinaryStream
            
                
        
            $archivoDestino =  $finalFolder + "\" + $archivo
            
            $stream = new-object System.IO.FileStream($archivoDestino,[System.IO.FileMode]'Create',[System.IO.FileAccess]'Write')
            
            #Loop through the server content stream, copying bytes into the buffer, 4k at a time, until there are no more bytes
            $ReadBytes = 0
        
            Write-Host "Descargando Archivos" -ForegroundColor Green
            do {
                if ($ReadBytes -Ne 0) {
                    
                    $Stream.Write($buffer,0,$ReadBytes)
                }
                $ReadBytes = $Content.Read($Buffer, 0,4096)
        
            } until ($ReadBytes -eq 0)
        
                $archivo
            
                #clean up
                $Stream.Close()
                $content.Close()
            }
            Catch {
                # IO error, Handle it
            }
        }
        
        }else{
        Write-Host "No contiene archivos adjuntos, No se crea Carpeta" -ForegroundColor Green
        
        }
        
        }

    #endregion

#endregion
