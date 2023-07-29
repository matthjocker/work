import-module SMlets
#import-module E:\trabajo\migrarIR\modulos\add-actionLogEntry.psm1

[Threading.Thread]::CurrentThread.CurrentCulture = 'es-ES'

#$ServiceUser = "trabajo\appSCSM2019ProdWFL"
#d1q4E3EO.,Vv1l@
$ServiceUser = "trabajo\appSCSM2019QaWFL"
#8o8g05MU5IV/kdk

$ServiceUser = "trabajo\SCSM2019DesaWFL"
#87-,49\@tH+4aH-



#Revisar: sacar comentario a cred

<#
$cred = Get-Credential -credential $ServiceUser

if ( ($cred).length -eq "0"){
write-host "falta service user" -ForegroundColor Redx
break
}

#>

#$objIR  = "" Import-Clixml C:\temp\Incidentes_Test.xml

$pathFunciones = "E:\trabajo\migrarIR\"



. $pathFunciones\get-requerimientos.ps1

. $pathFunciones\get-AttachReqV2.ps1


. $pathFunciones\get-actionLogFullv2.ps1


. $pathFunciones\UploadAttachReqv2.ps1

. $pathFunciones\add-actionLogEntryV2.ps1


Function Get-hijosIncidentes{

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


 Function get-padreIncident {
param(
$wi,
$server

)


$Padre_relObj = Get-SCSMRelationshipObject -BySource $wi -ComputerName $server | where{ $_.RelationshipId -eq 'da3123d1-2b52-a281-6f42-33d0c1f06ab4'} 

#$Padre_relObj.TargetObject.id.guid 

$padreGuid  = $Padre_relObj.TargetObject.id.guid  

$Padre_obj = Get-SCSMObject -id $padreGuid -ComputerName $server

return $Padre_obj
}





#$GLOBAL:smdefaultcomputer = "s1-hixx-ssm01.ministerio.trabajo.gov.ar" 
# --------------------
$servidorOrigen = "scsm.ministerio.trabajo.gov.ar"
$servidorDestino = "s1-dixx-ssm04"
#$servidorDestino = "s1-hixx-ssm01"


$SRClassOrigen = Get-SCSMClass -Name System.WorkItem.ServiceRequest$ -ComputerName $servidorOrigen
$IRclassOrigen = Get-SCSMclass -name System.Workitem.Incident$ -ComputerName $servidorOrigen 


$IRclassDestino = Get-SCSMclass -name System.Workitem.Incident$ -ComputerName $servidorDestino 
$SRclassDestino = Get-SCSMClass -Name system.workitem.servicerequest$ -ComputerName $servidorDestino

$UserClass = Get-SCSMClass -name System.Domain.User$ -ComputerName $servidorOrigen # Get SCSM User class object

$relAffectedUser = Get-SCSMRelationshipClass -Name System.WorkItemAffectedUser -ComputerName $servidorOrigen # Get SCSM relationship Affected User
$AssignedToRel = get-scsmrelationshipclass -name System.WorkItemAssignedToUser$ -ComputerName $servidorOrigen
$WorkItemAffectedUserRel = Get-SCSMRelationshipClass System.WorkItemAffectedUser$ -ComputerName $servidorOrigen
$AssignedToUserRel = Get-scsmrelationshipclass -name System.WorkItemAssignedToUser$ -ComputerName $servidorOrigen
$WorkItemAffectedUserRel = Get-SCSMRelationshipClass System.WorkItemAffectedUser$ -ComputerName $servidorOrigen


$padreHijoRel = Get-SCSMRelationshipClass -Name System.WorkItemRelatesToWorkItem -ComputerName $servidorDestino

#$IncidentPrimaryOwnerRelClass = Get-SCSMRelationshipClass -Name System.WorkItem.IncidentPrimaryOwner$ -ComputerName $servidorOrigen
#$Username = "SCSM_Usuario_Prueba2"
#$nombre = "Carlos Braian Humeres"

$basePath = "C:\temp\reqExport\"

$logPath ="E:\trabajo\migrarIR\logs\logs_migracion.txt"
$logPathSoloIncidentesHijos = "E:\trabajo\migrarIR\logs\logs_migracion_IncHijos.txt"

#revisar: lote por entre fechas?

$objIR = get-requrimientos -clase ir -status "irNoCloseNoResolved" -servidor $servidorOrigen 

$hijos = $objIR | ? {$_.status.displayname -match "depende"}

$hijos.Count
$hijos | select id | Out-GridView




$directorio = "C:\temp\reqExport"



#$directoriosCreados = (Get-ChildItem -Path $directorio -Directory | Where-Object { $_.LastWriteTime.Date -eq (Get-Date).Date }).name
#$directoriosCreados.count
#$directoriosCreados | Out-GridView



#$crearIncidentesHijos | select * | Out-GridView


$hijos| ForEach-Object {

$wi = $_


$log ="procesando $wi"
write-host $log -ForegroundColor Yellow
-join($wi.id, " - " ,$log ) | out-file $logPath -Append


# write-host $log -ForegroundColor Yellow
# $wi.id |out-file $logPathSoloIncidentesHijos -Append



$AffectedUser = Get-SCSMRelatedObject  -Relationship $WorkItemAffectedUserRel -SMObject $_ -ComputerName $servidorOrigen

$AssignedToUser = Get-SCSMRelatedObject  -Relationship $AssignedToUserRel -SMObject $_ -ComputerName $servidorOrigen

$username =  $AffectedUser.UserName 

$analist = $AssignedToUser.DisplayName

$origenID = $_.Id

$irAffectedUser = Get-SCSMObject -Class $UserClass -Filter "Username -eq $username" -ComputerName $servidorOrigen 

$userAnalist = Get-SCSMObject -Class $UserClass -Filter "Username -eq $analist" -ComputerName $servidorOrigen 



#region obtener adjuntos

#descarga los adjuntos en la ruta $basePath con una carpeta con el nombre de $wi.id, ejemplo : c:\temp\ir0001 , solo si contiene archivos adjuntos

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

#($actionLog[0].comentario).length

$actLog = [system.String]::Join(" ", $onlyComent)

#endregion



#$clasificacion = (get-scsmenumeration -ComputerName $servidorDestino| ? {$_.displayname -match $wi.Classification.displayname} |  ? {$_.Identifier -match $class}).displayname #get-enum $wi.Classification.displayname $servidorDestino
$TierQueue  = (get-scsmenumeration -ComputerName $servidorDestino| ? {$_.displayname -match $wi.TierQueue.displayname} | ? {$_.Identifier -match "incidente"}).name


$status  = ( get-scsmenumeration -ComputerName $servidorDestino |  ? {$_.name -eq  $wi.Status.name}).displayname

$clasificacion = "Pendiente de categorización"

#revisar: cambiar "IR{0}" por $wi.id

$properties = @{
    Id             = "IR{0}" #$wi.id #
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

$log = "se migro incidente ID $($wi.id) con origen en $($servidororigen) -> destino $($servidorDestino) - Requerimiento: $($newReq)"

write-host $log -ForegroundColor Yellow

-join($wi.id, "-" ,$log ) | out-file $logPath -Append


#region cracion action log

if ($actLog -and $newReq ){

    Add-ActionLogEntry -WIObject $newReq  -Action "AnalystComment" -Comment $actLog  -EnteredBy $ServiceUser -IsPrivate $false -server $servidorDestino

    $log = "se agregó los comentarios de analistas y usuarios en el actionLog ID $($wi.id) con origen en $($servidororigen) -> destino $($servidorDestino) - Requerimiento: $($newReq)"

    write-host $log

    -join($wi.id, "-" ,$log ) | out-file $logPath -Append

}else{

    $log = "Error al importar action log - Requerimiento: $($newReq)"

    write-host $log
    -join($wi.id, "-" ,$log ) | out-file $logPath -Append

}

#endregion

#region Relacionar padre e hijo


$incidentePadre = $_


$childWis_relObj = Get-SCSMRelationshipObject -ByTarget $incidentePadre -ComputerName $servidorOrigen | where{ $_.RelationshipId -eq 'da3123d1-2b52-a281-6f42-33d0c1f06ab4'} 

$incidentesHIjos = Get-hijosIncidentes $_ $servidorOrigen

$incidentesHIjos| ForEach-Object {

    $child = $_

    New-SCSMRelationshipObject -RelationShip $padreHijorel -Source $incidentePadre -Target $child -Bulk -ComputerName $servidorDestino

    $log = "se agrega relacion de padre $($incidentePadre) a hijo  $($child) -> destino $($servidorDestino) - Requerimiento: $($newReq)"

    write-host $log

    -join($wi.id, "-" ,$log ) | out-file $logPath -Append

}

#endregion

# Setear Affected User
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
    

     write-host "-------------------------------------------------------------------------------------"
     

#region upload archivos Adjuntos

   write-host "Preparando para subir archivos" -ForegroundColor Cyan

    $classObj = ($wi.id).substring(0, 2)

    #ruta de la carpeta con nombre del requerimiento
    $FullDirPath = $basePath + $wi.id + "\";

    $AttachmentEntries = [IO.Directory]::GetFiles($FullDirPath); 

    $AttachmentArray = $AttachmentEntries.count;

                     if ($AttachmentArray -ne $NULL){
              
                          foreach($SingleAttachment in $AttachmentEntries) {
                        
                            #obtener extension
                            #$extn = [IO.Path]::GetExtension($SingleAttachment)
                                 
                                         $AttachmentSingleName = split-path $SingleAttachment -leaf


                                                 if ($classObj -eq "sr"){

                                                        if ((Get-SCSMObject -Class $srClassOrigen -filter "Id -eq $wi.id" | where {$_.FileAttachment -like $AttachmentSingleName}) -eq $NULL){

                                                                            

                                                                 #   Insert-Attachment -SCSMID $newReq.id -Directory $SingleAttachment -tipoClase $classObj -server $servidorDestino
                                                                 
                            
                                                                            
                                                                $log = "Subiendo $($AttachmentSingleName) de la carpeta: $($SingleAttachment) -> subido ServiceRequest with ID: $($newReq.id) "
                                                       
                                                                   write-host $log -ForegroundColor Green
                                                       
                                                                    -join($wi.id, "-" ,$log ) | out-file $logPath -Append
                                                        }
                                          
                                                 }



                                                 if ($classObj -eq "ir"){

                                                         if ((Get-SCSMObject -Class $IRclassOrigen -filter "Id -eq $($wi.id)" -ComputerName $servidorOrigen | where {$_.FileAttachment -like $AttachmentSingleName}) -eq $NULL){ 
                                                          
                                                             
                                                                 #  Insert-Attachment -SCSMID $newReq.id -Directory $SingleAttachment -tipoClase $classObj -server $servidorDestino
                                                             
                                                              
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



$directoriosCreados = (Get-ChildItem -Path $directorio -Directory | Where-Object { $_.LastWriteTime.Date -eq (Get-Date).Date }).name
$directoriosCreados.count