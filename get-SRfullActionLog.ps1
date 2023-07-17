

Import-Module SMlets 

[Threading.Thread]::CurrentThread.CurrentCulture = 'es-ES'

$GLOBAL:smdefaultcomputer = "s1-dixx-ssm04"

$SRClass = Get-SCSMClass -Name System.WorkItem.ServiceRequest$
$IRclass=Get-SCSMclass -name System.Workitem.Incident$ # Get SCSM Incident class object

$EndUserAnalystCommentRel = Get-SCSMRelationshipClass "System.WorkItemHasCommentLog"
$ActionCommentRel = Get-SCSMRelationshipClass "System.WorkItemHasActionLog"

 Function Crear-Objeto{
    Param (
    $obj,
    $commentBy,
    $UserComment,
    $EnteredDate  
    )

# Figure out your current offset from UTC


$ReqProp = [ordered]@{

#$EnteredDate = "16:26:48"
'Identificardor' =$obj.id;
'AgregadoPor' = $commentBy;
'comentario' = $UserComment ;
#'Fecha de creacion'= get-date ($obj.CreatedDate) -format  "dd-MM-yyyy";#Fecha Creado
'enteredDate'= $EnteredDate
#'Fecha de ultima modificacion'=get-d)ate ($obj.LastModified) -format  "dd-MM-yyyy";
#'Titulo'= $obj.Title;
#'Descripcion'= $obj.Description;
#'Source'= $obj.Source.displayname;


} 

$CustomObjects = New-Object -TypeName Psobject -Property  $ReqProp
    
return $CustomObjects
  
}

function get-SRfullActionLog{

param ($obj)   

    $ObjetoReq = @()
     
         $obj | ForEach-Object {

             $requerimiento = $_

             $AffectedUser = Get-SCSMRelatedObject -Relationship $WorkItemAffectedUserRel -SMObject $_

              $EndUserAnalystComments = Get-SCSMRelatedObject -SMObject $_ -Relationship $EndUserAnalystCommentRel
        
              $EndUserAnalystComments | select enteredby, comment, lastModified ,enteredDate | Sort-Object -Property lastModified


             $ObjetoReq += Crear-Objeto $requerimiento  $_.EnteredBy $_.Comment  $_.EnteredDate 

        }
        return  $ObjetoReq 
}




 $wi = "SR7919"

 $class = switch ($wi.substring(0, 2))

    {
        IR {$IRClass}
        SR {$SRClass}
        MA {$MAclass}
       
    }

$wiObject = Get-SCSMObject -Class $class -Filter "Id -eq $wi"

get-SRfullActionLog $wiObject 

