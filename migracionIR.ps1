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

. $pathFunciones\get-AttachReqV2.ps1


. $pathFunciones\get-actionLogFullv2.ps1


. $pathFunciones\UploadAttachReqv2.ps1

. $pathFunciones\add-actionLogEntryV2.ps1



#endregion

#region constantes
$servidorOrigen = "scsm.ministerio.trabajo.gov.ar"
$servidorDestino = "s1-dixx-ssm04"
#$servidorDestino = "s1-hixx-ssm01"

$SRClassOrigen = Get-SCSMClass -Name System.WorkItem.ServiceRequest$ -ComputerName $servidorOrigen
$IRclassOrigen = Get-SCSMclass -name System.Workitem.Incident$ -ComputerName $servidorOrigen 

$IRclassDestino = Get-SCSMclass -name System.Workitem.Incident$ -ComputerName $servidorDestino 
$SRclassDestino = Get-SCSMClass -Name system.workitem.servicerequest$ -ComputerName $servidorDestino

$UserClass = Get-SCSMClass -name System.Domain.User$ -ComputerName $servidorOrigen # Get SCSM User class object

$basePath = "C:\temp\reqExport\"

$logPath ="E:\trabajo\migrarIR\logs\logs_migracion.txt"

#endregion

#region Relaciones
$relAffectedUser = Get-SCSMRelationshipClass -Name System.WorkItemAffectedUser -ComputerName $servidorOrigen # Get SCSM relationship Affected User
$AssignedToRel = get-scsmrelationshipclass -name System.WorkItemAssignedToUser$ -ComputerName $servidorOrigen
$WorkItemAffectedUserRel = Get-SCSMRelationshipClass System.WorkItemAffectedUser$ -ComputerName $servidorOrigen
$AssignedToUserRel = Get-scsmrelationshipclass -name System.WorkItemAssignedToUser$ -ComputerName $servidorOrigen
$WorkItemAffectedUserRel = Get-SCSMRelationshipClass System.WorkItemAffectedUser$ -ComputerName $servidorOrigen
$padreHijoRel = Get-SCSMRelationshipClass -Name System.WorkItemRelatesToWorkItem -ComputerName $servidorDestino

#endregion

#region obtener IR activos

#obtener IR activos
$objIR = get-requrimientos -clase ir -status "irNoCloseNoResolved" -servidor $servidorOrigen 

$objIR.count

#obtener pedidos que no sean padres y no sean hijos
$objetosFinales =  $objIR | ? {$_.IsParent -ne  "True" -and $_.status.displayname -notmatch "depende"}

$objetosFinales.count
#endregion

#region main

$objetosFinales[1..5] | ForEach-Object {

$wi = $_

$log ="procesando $wi"
write-host $log -ForegroundColor Yellow
-join($wi.id, " - " ,$log ) | out-file $logPath -Append

#region obtener datos de usuario y analista
$AffectedUser = Get-SCSMRelatedObject  -Relationship $WorkItemAffectedUserRel -SMObject $_ -ComputerName $servidorOrigen

$AssignedToUser = Get-SCSMRelatedObject  -Relationship $AssignedToUserRel -SMObject $_ -ComputerName $servidorOrigen

$username =  $AffectedUser.UserName 

$analist = $AssignedToUser.DisplayName

$irAffectedUser = Get-SCSMObject -Class $UserClass -Filter "Username -eq $username" -ComputerName $servidorOrigen 

$userAnalist = Get-SCSMObject -Class $UserClass -Filter "Username -eq $analist" -ComputerName $servidorOrigen 
#endregion


#region obtener adjuntos

#descarga los adjuntos en la ruta $basePath con una carpeta con el nombre de $wi.id, ejemplo : c:\temp\ir0001 ,solo si contiene archivos adjuntos

get-AttachReq -wi $wi.id -OutputFolder $basePath -servidor $servidorOrigen 

#endregion


#region obtener actionlog



$actionLog = get-actionLogFull $wi $servidorOrigen #| out-file  $actionLogPath

$onlyComent = @()

$actionLog | ForEach-Object {

    if (($_.comentario).length -ne "0" ){

     $onlyComent  += $_
    }
}


$actLog = [system.String]::Join(" ", $onlyComent)

#endregion

#region crear incidente en el servidor remoto

$TierQueue  = (get-scsmenumeration -ComputerName $servidorDestino| ? {$_.displayname -match $wi.TierQueue.displayname} | ? {$_.Identifier -match "incidente"}).name

$status  = ( get-scsmenumeration -ComputerName $servidorDestino |  ? {$_.name -eq  $wi.Status.name}).displayname

$clasificacion = "Pendiente de categorización"

#revisar: cambiar "IR{0}" por $wi.id

$properties = @{
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
   # "Gestión de Requerimientos de TI y Analítica de Datos" #
    
   
}

try{
$newReq = New-SCSMObject -Class $IRclassDestino -PropertyHashtable $properties -PassThru -ComputerName $servidorDestino #-Credential $cred
}catch{
 $Error[0].Exception 
 $Error[0].CategoryInfo 
 $exception = $Error[0].Exception 

 -join($wi.id, $Error[0].Exception ) | out-file $logPath -Append

}
#endregion


#region agregar log
$log = "se migro incidente ID $($wi.id) con origen en $($servidororigen) -> destino $($servidorDestino) - Requerimiento: $($newReq)"
write-host $log -ForegroundColor Yellow
-join($wi.id, "-" ,$log ) | out-file $logPath -Append


if ($actLog -and $newReq ){
Add-ActionLogEntry -WIObject $newReq  -Action "AnalystComment" -Comment $actLog  -EnteredBy $ServiceUser -IsPrivate $false -server $servidorDestino

$log = "se agregó los comentarios de analistas y usuarios en el actionLog ID $($wi.id) con origen en $($servidororigen) -> destino $($servidorDestino) - Requerimiento: $($newReq)"
write-host $log  -ForegroundColor Yellow

-join($wi.id, "-" ,$log ) | out-file $logPath -Append

}

#endregion


#region Relacionar al Affected User

if ($irAffectedUser -and $newReq) {

     New-SCSMRelationshipObject -RelationShip $relAffectedUser -Source $newReq -Target $irAffectedUser -Bulk -ComputerName $servidorDestino

     if (!($userAnalist)){
     $log = "no existe  analista asignado"
     write-host $log -ForegroundColor DarkMagenta
     -join($wi.id, "-" ,$log ) | out-file $logPath -Append
    
     }else {
      New-SCSMRelationshipObject -Relationship $AssignedToRel -Source $newReq -Target $userAnalist -Bulk -ComputerName $servidorDestino
      $log = "se asignó al analista $($userAnalist)"
      write-host $log -ForegroundColor Yellow
      -join($wi.id, "-" ,$log ) | out-file $logPath -Append
     }
    
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
                                                          
                                                             
                                                                   Insert-Attachment -SCSMID $newReq.id -Directory $SingleAttachment -tipoClase $classObj -server $servidorDestino
                                                             
                                                              
                                                              $log = "$AttachmentSingleName from Folder $SingleAttachment -> subido al  incidente con ID: $newReq.id "

                                                               write-host $log   -ForegroundColor DarkYellow
                                                               -join($wi.id, "-" ,$log ) | out-file $logPath -Append

                                                         }         
                                                 }              
                                        


                          }#finFor
                      
                      }#finIF


#endregion
    



}



}

#endregion