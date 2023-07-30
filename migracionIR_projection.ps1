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

#region funciones
$pathFunciones = "E:\trabajo\migrarIR\"

. $pathFunciones\get-requerimientos.ps1

. $pathFunciones\add-actionLogEntryV2.ps1

. $pathFunciones\get-AttachReqV2.ps1

. $pathFunciones\UploadAttachReqv2.ps1

function Get-hijosIncidentes{


 param (
  [Parameter(Mandatory = $True)]
  $wi,

  $server
 )


 $childWIs_obj = @()
 
 $childWIs_relobj = Get-SCSMRelationshipObject -ByTarget $wi -ComputerName $server| where{ $_.RelationshipId -eq 'da3123d1-2b52-a281-6f42-33d0c1f06ab4'}

 ForEach ($childWI_relobj in $childWis_relObj)
 {
   $childWI_id = $childWI_relobj.SourceObject.id.guid
   $childWI_obj = Get-SCSMObject -id $childWI_id -ComputerName $servidorOrigen
   If ($childWI_obj.ClassName -eq 'System.WorkItem.Incident')
   {
    $childWIs_obj += $childWI_obj
   }
 }
 if ($childWIs_obj.length -gt 0){
 return $childWIs_obj
 }else{
 write-host "sin hijos" -ForegroundColor Yellow
 }

 }

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
    hijos { 
            $outputPath =  $logPathSoloIncidentesPadre ;
            $mensaje = -join($fechaActual," - " , $wi_id, " - " ,"$($wi_id)" )  
    }
    general {$outputPath =  $logPath ; 
            $mensaje = -join($fechaActual," - " , $wi_id, " - " ,"Procesando $($wi_id)" )         
    }
    Adjunto {$outputPath =  $logPath ; 
            $mensaje = -join($fechaActual," - " , $wi_id, " - " ,"Procesando $($wi_id)" )         
    }

    comentario {$outputPath =  $logPath ; 
            $mensaje = -join($fechaActual," - " , $wi_id, " - " ,"Procesando $($wi_id)" )         
    }
}
 
write-host $mensaje  -ForegroundColor Yellow
$mensaje | out-file $outputPath -Append
}
#endregion

#region constantes
$servidorOrigen = "scsm.ministerio.trabajo.gov.ar"
$servidorDestino = "s1-dixx-ssm04"
#$servidorDestino = "s1-hixx-ssm01"

$basePath = "C:\temp\reqExport\"

$logPath ="E:\trabajo\migrarIR\logs\logs_migracion.txt"

$incidentTypeProjectionOrigen = Get-SCSMTypeProjection -name System.WorkItem.IncidentPortalProjection -ComputerName $servidorOrigen 

$incidentTypeProjectionDestino = Get-SCSMTypeProjection -name System.WorkItem.IncidentPortalProjection  -ComputerName $servidorDestino
#Get-SCSMTypeProjection -ComputerName $servidorDestino  | select typeprojection  | Out-GridView
#endregion

#region Relaciones
$relAffectedUser = Get-SCSMRelationshipClass -Name System.WorkItemAffectedUser -ComputerName $servidorOrigen # Get SCSM relationship Affected User
$AssignedToRel = get-scsmrelationshipclass -name System.WorkItemAssignedToUser$ -ComputerName $servidorOrigen
$WorkItemAffectedUserRel = Get-SCSMRelationshipClass System.WorkItemAffectedUser$ -ComputerName $servidorOrigen
$AssignedToUserRel = Get-scsmrelationshipclass -name System.WorkItemAssignedToUser$ -ComputerName $servidorOrigen
$WorkItemAffectedUserRel = Get-SCSMRelationshipClass System.WorkItemAffectedUser$ -ComputerName $servidorOrigen

$padreHijoRelDestino = Get-SCSMRelationshipClass -Name System.WorkItemRelatesToWorkItem -ComputerName $servidorDestino

$padreHijorelOrigen  = Get-SCSMRelationshipClass -Name System.WorkItemHasParentWorkItem -ComputerName $servidorOrigen 
#endregion

#region obtener IR activos - NoCloseNoResolved

#obtener IR activos
$objIR = get-requerimientos -clase ir -status "irNoCloseNoResolved" -servidor $servidorOrigen 

$objIR.count

#obtener pedidos que no sean padres y no sean hijos
$objetosFinales =  $objIR | ? {$_.IsParent -ne  "True" -and $_.status.displayname -notmatch "depende"}

$objetosFinales.count
#endregion


#region main

$objetosFinales | ? {$_.id -eq "IR2251653"} | ForEach-Object {

#region crear incidente en el servidor remoto
$wi = $_

#obtengo AffectedUser, createdby , AssignedTo y comentarios, mea ahorro de traer las relaciones por cada uno de los mencionados           
$incidentRequestProjection = Get-SCSMObjectProjection -ProjectionName $incidentTypeProjectionOrigen.name -filter “ID -eq $($wi.id)” -ComputerName $servidorOrigen 

$grupoSoporte = @{
    "Desarrollo - Juicios" = "Desarrollo"
    "IT Producción" = "Producción"
    "Networking" = "Redes"
    "Plataforma/Soft base" = "Plataforma"
    "Seguridad" = "Seguridad Informática"
    "Software de 3ros" = "Software de Terceros"
    "Soporte funcional" = "Soporte Funcional de Aplicativos"
    "Soluciones Tecnológicas" = "Servicios Tecnológicos"
}

$fixArea = $grupoSoporte[$wi.TierQueue.displayname]
if ($fixArea) {
    Write-Output "se corrigió $($fixArea)"
    $TierQueue  = (get-scsmenumeration -ComputerName $servidorDestino| ? {$_.displayname -match $fixArea} | ? {$_.Identifier -match "incidente"}).name
} else {
    $TierQueue  = (get-scsmenumeration -ComputerName $servidorDestino| ? {$_.displayname -match $wi.TierQueue.displayname} | ? {$_.Identifier -match "incidente"}).name
}

if ($wi.Status.displayname -ne "En progreso" ){
    $status  = ( get-scsmenumeration -ComputerName $servidorDestino |  ? {$_.name -eq  $wi.Status.name}).displayname
    }else{
    $status  = (get-scsmenumeration -ComputerName $servidorDestino|  ? {$_.name -eq  "IncidentStatusEnum.Active"}).displayname
}

$clasificacion = "Pendiente de categorización"

#revisar: cambiar "IR{0}" por $wi.id

$irproperties = @{
    Id             = "IR{0}" #$wi.id
    Title          = $wi.Title
    Description    = $wi.Description
    Urgency        = $wi.Urgency
    Impact         = $wi.Impact
    Source         = $wi.Source
    Status         = $status 
    Classification = $clasificacion 
    TierQueue      = $TierQueue 
    _WI = $wi._WI
    createdDate = $wi.CreatedDate
    
   
     }

     #SE COMENTA LAS 4 PROPIEDADES PARA TESTEO, FUNCIONA TODO MENOS AnalystCommentLog
     $IRProjection = @{__CLASS = "System.Workitem.Incident";
                 __OBJECT =   $irproperties 
                 #AffectedUser =  $incidentRequestProjection.AffectedUser
                 #CreatedByUser = $incidentRequestProjection.CreatedBy
                 #AssignedTo = $incidentRequestProjection.AssignedTo 
                 #AnalystCommentLog =   $actionLogComment
                }
 
try{
$new_IR = New-SCSMObjectProjection -Type System.WorkItem.IncidentPortalProjection -Projection $IRProjection -PassThru  -ComputerName $servidorDestino # -Credential $cred

    $new_IR.Object


    write-log $new_IR.Object 
    $log = "se migro incidente ID $($wi.id) con origen en $($servidororigen) -> destino $($servidorDestino) - Requerimiento: $($new_IR.Object)"
    write-host $log -ForegroundColor Yellow
  

}catch{
     $Error[0].Exception 
     $Error[0].CategoryInfo 
     $exception = $Error[0].Exception 
}
#endregion

#region comentarios
$todosLosComentarios = $incidentRequestProjection.AnalystComments

if ($todosLosComentarios.length -ne 0){
   $todosLosComentarios | ForEach-Object{
   $objcomment = $_

    #necesito el objeto $new_IR.Object creado para indicar en el __SEED

    # Generate a new GUID for the Action Log entry
    $NewGUID = ([guid]::NewGuid()).ToString()
    # Create the object projection with properties
    $IRProjectionComment = @{__CLASS = "System.WorkItem.Incident";
                    __SEED = $new_IR.Object;
                    "AnalystComments" = @{__CLASS = $CommentClass;
                                        __OBJECT = @{Id = $NewGUID;
                                            DisplayName = $NewGUID;
                                            ActionType = $ActionType;
                                            Comment = $objcomment.comment;
                                            #Title = "$($ActionEnum.DisplayName)";
                                            EnteredBy  = $objcomment.EnteredBy;
                                            EnteredDate = $objcomment.EnteredDate
                                            IsPrivate = $objcomment.IsPrivate 
                                        }
                    }
    }

<# No funciona
$actionLogComment =@()
 $IRProjectionComment = @{__CLASS = "System.Workitem.Incident";
                 __SEED =   $new_IR.Object
                  AnalystComments  =   $IRProjectionComment 
                
                }
#>

$new_IRcomment = New-SCSMObjectProjection -Type System.WorkItem.IncidentPortalProjection -Projection $IRProjectionComment -PassThru  -ComputerName $servidorDestino # -Credential $cred

$log = "se agregó los comentarios de analistas y usuarios en el actionLog ID $($wi.id) con origen en $($servidororigen) -> destino $($servidorDestino) - Requerimiento: $($new_IR.Object)"
write-host $log  -ForegroundColor Yellow

-join($wi.id, "-" ,$log ) | out-file $logPath -Append

}


}else{

    $log = "No posee comentarios $($wi.id) con origen en $($servidororigen) -> destino $($servidorDestino) - Requerimiento: $($new_IR.Object.name)"
    write-host $log  -ForegroundColor Yellow
    -join($wi.id, "-" ,$log ) | out-file $logPath -Append

}

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

    $AttachmentEntries = [IO.Directory]::GetFiles($FullDirPath); 

    $AttachmentArray = $AttachmentEntries.count;

                     if ($AttachmentArray -ne $NULL){
              
                          foreach($SingleAttachment in $AttachmentEntries) {
                                                        
                                         $AttachmentSingleName = split-path $SingleAttachment -leaf

                                                 if ($classObj -eq "ir"){

                                                         if ((Get-SCSMObject -Class $IRclassOrigen -filter "Id -eq $($wi.id)" -ComputerName $servidorOrigen | where {$_.FileAttachment -like $AttachmentSingleName}) -eq $NULL){ 
                                                          
                                                             
                                                                  #Insert-Attachment -SCSMID $new_IR.Object.Id -Directory $SingleAttachment -tipoClase $classObj -server $servidorDestino
                                                             
                                                              
                                                              $log = "$AttachmentSingleName from Folder $SingleAttachment -> subido al  incidente con ID: $new_IR.Object.Id"

                                                               write-host $log   -ForegroundColor DarkYellow
                                                               -join($wi.id, "-" ,$log ) | out-file $logPath -Append

                                                         }         
                                                 }              
                                        


                          }#finFor
                      
                      }#finIF


#endregion

}

#endregion

#region padres

$Parents = $objIR |? {$_.IsParent -eq  "True"} 
$Parents[1] | ForEach-Object {

#region crear incidente en el servidor remoto
$wi = $_

#obtengo AffectedUser, createdby , AssignedTo y comentarios, mea ahorro de traer las relaciones por cada uno de los mencionados   , NO PUDE HACER UN -AND           
$incidentRequestProjection = Get-SCSMObjectProjection -ProjectionName $incidentTypeProjectionOrigen.name -filter “ID -eq $($wi.id)” -ComputerName $servidorOrigen 

$grupoSoporte = @{
    "Desarrollo - Juicios" = "Desarrollo"
    "IT Producción" = "Producción"
    "Networking" = "Redes"
    "Plataforma/Soft base" = "Plataforma"
    "Seguridad" = "Seguridad Informática"
    "Software de 3ros" = "Software de Terceros"
    "Soporte funcional" = "Soporte Funcional de Aplicativos"
    "Soluciones Tecnológicas" = "Servicios Tecnológicos"
}

$fixArea = $grupoSoporte[$wi.TierQueue.displayname]

if ($fixArea) {
    Write-Output "se corrigió $($fixArea)"
        $TierQueue  = (get-scsmenumeration -ComputerName $servidorDestino| ? {$_.displayname -match $fixArea} | ? {$_.Identifier -match "incidente"}).name
    } else {
        $TierQueue  = (get-scsmenumeration -ComputerName $servidorDestino| ? {$_.displayname -match $wi.TierQueue.displayname} | ? {$_.Identifier -match "incidente"}).name
    }

if ($wi.Status.displayname -ne "En progreso" ){
        $status  = ( get-scsmenumeration -ComputerName $servidorDestino |  ? {$_.name -eq  $wi.Status.name}).displayname
    }else{
        $status  = (get-scsmenumeration -ComputerName $servidorDestino|  ? {$_.name -eq  "IncidentStatusEnum.Active"}).displayname
    }
$clasificacion = "Pendiente de categorización"

#revisar: cambiar "IR{0}" por $wi.id

$irproperties = @{
    Id             = "IR{0}" #$wi.id
    Title          = $wi.Title
    Description    = $wi.Description
    Urgency        = $wi.Urgency
    Impact         = $wi.Impact
    Source         = $wi.Source
    Status         = $status 
    Classification = $clasificacion 
    TierQueue      = $TierQueue 
    _WI = $wi._WI
    createdDate = $wi.CreatedDate
    
   
     }

     #SE COMENTA LAS 4 PROPIEDADES PARA TESTEO, FUNCIONA TODO MENOS AnalystCommentLog
     $IRProjection = @{__CLASS = "System.Workitem.Incident";
                 __OBJECT =   $irproperties 
                 #AffectedUser =  $incidentRequestProjection.AffectedUser
                 #CreatedByUser = $incidentRequestProjection.CreatedBy
                 #AssignedTo = $incidentRequestProjection.AssignedTo 
                 #AnalystCommentLog =   $actionLogComment
                }
 

$new_ParentIR = New-SCSMObjectProjection -Type System.WorkItem.IncidentPortalProjection -Projection $IRProjection -PassThru  -ComputerName $servidorDestino # -Credential $cred

$new_ParentIR.Object
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

    $AttachmentEntries = [IO.Directory]::GetFiles($FullDirPath); 

    $AttachmentArray = $AttachmentEntries.count;

                     if ($AttachmentArray -ne $NULL){
              
                          foreach($SingleAttachment in $AttachmentEntries) {
                                                        
                                         $AttachmentSingleName = split-path $SingleAttachment -leaf

                                                 if ($classObj -eq "ir"){

                                                         if ((Get-SCSMObject -Class $IRclassOrigen -filter "Id -eq $($wi.id)" -ComputerName $servidorOrigen | where {$_.FileAttachment -like $AttachmentSingleName}) -eq $NULL){ 
                                                          
                                                             
                                                                  Insert-Attachment -SCSMID $new_IR.Object.Id -Directory $SingleAttachment -tipoClase $classObj -server $servidorDestino
                                                             
                                                              
                                                              $log = "$AttachmentSingleName from Folder $SingleAttachment -> subido al  incidente con ID: $new_IR.Object.Id"

                                                               write-host $log   -ForegroundColor DarkYellow
                                                               -join($wi.id, "-" ,$log ) | out-file $logPath -Append

                                                         }         
                                                 }              
                                        


                          }#finFor
                      
                      }#finIF


#endregion upload archivos
  

#region crear hijos  y relacionarlo con el padre         
$childrens = Get-hijosIncidentes $wi $servidorOrigen

    $childrens| ForEach-Object{

#region crear incidente en el servidor remoto
        $wi = $_
        
        #obtengo AffectedUser, createdby , AssignedTo y comentarios, me ahorro de traer las relaciones por cada uno de los mencionados , NO PUDE HACER UN -AND          
        $incidentRequestProjection = Get-SCSMObjectProjection -ProjectionName $incidentTypeProjectionOrigen.name -filter “ID -eq $($wi.id)” -ComputerName $servidorOrigen 

        $grupoSoporte = @{
            "Desarrollo - Juicios" = "Desarrollo"
            "IT Producción" = "Producción"
            "Networking" = "Redes"
            "Plataforma/Soft base" = "Plataforma"
            "Seguridad" = "Seguridad Informática"
            "Software de 3ros" = "Software de Terceros"
            "Soporte funcional" = "Soporte Funcional de Aplicativos"
            "Soluciones Tecnológicas" = "Servicios Tecnológicos"
        }

        $fixArea = $grupoSoporte[$wi.TierQueue.displayname]

        if ($fixArea) {
            Write-Output "se corrigió $($fixArea)"
                $TierQueue  = (get-scsmenumeration -ComputerName $servidorDestino| ? {$_.displayname -match $fixArea} | ? {$_.Identifier -match "incidente"}).name
            } else {
                $TierQueue  = (get-scsmenumeration -ComputerName $servidorDestino| ? {$_.displayname -match $wi.TierQueue.displayname} | ? {$_.Identifier -match "incidente"}).name
            }

        if ($wi.Status.displayname -ne "En progreso" ){
                $status  = ( get-scsmenumeration -ComputerName $servidorDestino | ? {$_.displayname -eq  $wi.Status.displayname }).name
            }else{
                $status  = (get-scsmenumeration -ComputerName $servidorDestino|  ? {$_.name -eq  "IncidentStatusEnum.Active"}).name
            }
        $clasificacion = "Pendiente de categorización"

        #revisar: cambiar "IR{0}" por $wi.id

        $irproperties = @{
            Id             = "IR{0}" #$wi.id
            Title          = $wi.Title
            Description    = $wi.Description
            Urgency        = $wi.Urgency
            Impact         = $wi.Impact
            Source         = $wi.Source
            Status         = $status 
            Classification = $clasificacion 
            TierQueue      = $TierQueue 
            _WI = $wi._WI
            createdDate = $wi.CreatedDate
    
   
         }

             #SE COMENTA LAS 4 PROPIEDADES PARA TESTEO, FUNCIONA TODO MENOS AnalystCommentLog
             $IRProjection = @{__CLASS = "System.Workitem.Incident";
                         __OBJECT =   $irproperties 
                         #AffectedUser =  $incidentRequestProjection.AffectedUser
                         #CreatedByUser = $incidentRequestProjection.CreatedBy
                         #AssignedTo = $incidentRequestProjection.AssignedTo 
                         #AnalystCommentLog =   $actionLogComment
                        }
 

                        $new_childIR = New-SCSMObjectProjection -Type System.WorkItem.IncidentPortalProjection -Projection $IRProjection -PassThru  -ComputerName $servidorDestino # -Credential $cred

                        $new_childIR.Object

                        New-SCSMRelationshipObject -RelationShip $padreHijoRelDestino -Source $new_ParentIR.Object -Target $new_childIR.Object -Bulk -ComputerName $servidorDestino
                       
                        $log = "se agrega relacion de padre $($new_ParentIR.Object) a hijo  $($new_childIR.Object) -> destino $($servidorDestino) - Requerimiento: $($new_ParentIR.Object)"

                        write-host   $log
#endregion creacion remota
 #region obtener adjuntos

#descarga los adjuntos en la ruta $basePath con una carpeta con el nombre de $wi.id, ejemplo : c:\temp\ir0001 ,solo si contiene archivos adjuntos

get-AttachReq -wi $wi.id -OutputFolder $basePath -servidor $servidorOrigen 

#endregion
#region upload archivos Adjuntos

   write-host "Preparando para subir archivos" -ForegroundColor Cyan

    $classObj = ($wi.id).substring(0, 2)

    #ruta de la carpeta con nombre del requerimiento
    $FullDirPath = $basePath + $wi.id + "\";

    $AttachmentEntries = [IO.Directory]::GetFiles($FullDirPath); 

    $AttachmentArray = $AttachmentEntries.count;

                     if ($AttachmentArray -ne $NULL){
              
                          foreach($SingleAttachment in $AttachmentEntries) {
                                                        
                                         $AttachmentSingleName = split-path $SingleAttachment -leaf

                                                 if ($classObj -eq "ir"){

                                                         if ((Get-SCSMObject -Class $IRclassOrigen -filter "Id -eq $($wi.id)" -ComputerName $servidorOrigen | where {$_.FileAttachment -like $AttachmentSingleName}) -eq $NULL){ 
                                                          
                                                             
                                                                  Insert-Attachment -SCSMID $new_IR.Object.Id -Directory $SingleAttachment -tipoClase $classObj -server $servidorDestino
                                                             
                                                              
                                                              $log = "$AttachmentSingleName from Folder $SingleAttachment -> subido al  incidente con ID: $new_IR.Object.Id"

                                                               write-host $log   -ForegroundColor DarkYellow
                                                               -join($wi.id, "-" ,$log ) | out-file $logPath -Append

                                                         }         
                                                 }              
                                        


                          }#finFor
                      
                      }#finIF


#endregion
    

              
            
    }
 #endregion
         
}#end for

#endregion



     
