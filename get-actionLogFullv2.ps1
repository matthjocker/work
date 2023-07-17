
Import-Module SMlets 

[Threading.Thread]::CurrentThread.CurrentCulture = 'es-ES'
function get-actionLogFull {
param(


  [Parameter(Mandatory=$true)]
  $objElement,

  [Parameter(Mandatory=$true)]
  [string]$server
  

)

#$SRClass = Get-SCSMClass -Name System.WorkItem.ServiceRequest$ -ComputerName $server
#$IRclass =Get-SCSMclass -name System.Workitem.Incident$ -ComputerName $servidor # Get SCSM Incident class object

$SREndUserAnalystCommentRel = Get-SCSMRelationshipClass "System.WorkItemHasCommentLog" -ComputerName $server
$relIncidentActionLog = Get-SCSMRelationshipClass -Name System.WorkItem.TroubleTicketHasActionLogg$  -ComputerName $server

#$IRActionCommentRel = Get-SCSMRelationshipClass "System.WorkItemHasActionLog" -ComputerName $server





    $ObjetoReq = @()
     
   

         $objElement | ForEach-Object {

             $requerimiento = $_

            # write-host "$($_.clasname)" -ForegroundColor Yellow


             $AffectedUser = Get-SCSMRelatedObject -Relationship $WorkItemAffectedUserRel -SMObject $_ -ComputerName $server

              switch ($_.ClassName)
        {
            "System.WorkItem.Incident" {$CommentClassName = $IRrelIncidentActionLog }
            "System.WorkItem.ServiceRequest" {$CommentClassName =  $SREndUserAnalystCommentRel}

        }

   
            
            if ($_.ClassName -eq "System.WorkItem.Incident"){
        
             $AnalystComments = Get-SCSMRelatedObject -SMObject  $_ -Relationship $CommentClassName  -ComputerName $server

                 $AnalystComments | ForEach-Object {
        
                    $ObjetoReq += Crear-Objeto $requerimiento  $_.EnteredBy $_.Comment  $_.EnteredDate 


                  }
             

              
              }else{
                  $EndUserAnalystComments = Get-SCSMRelatedObject -SMObject $_ -Relationship $CommentClassName  -ComputerName $server

                     $EndUserAnalystComments | ForEach-Object {
        
                    $ObjetoReq += Crear-Objeto $requerimiento  $_.EnteredBy $_.Comment  $_.EnteredDate 
                  }
        
             # $EndUserAnalystComments | select enteredby, comment, lastModified ,enteredDate | Sort-Object -Property lastModified
             
             }



        }
        return  $ObjetoReq 



}

 Function Crear-Objeto{
    Param (
    $req,
    $commentBy,
    $UserComment,
    $EnteredDate  
    )



$ReqProp = [ordered]@{

#$EnteredDate = "16:26:48"
'Identificardor' =$req.id;
'AgregadoPor' = $commentBy;
'comentario' = $UserComment ;
#'Fecha de creacion'= get-date ($obj.CreatedDate) -format  "dd-MM-yyyy";#Fecha Creado
'enteredDate' = $EnteredDate
#'Fecha de ultima modificacion'=get-date ($obj.LastModified) -format  "dd-MM-yyyy";
#'Titulo'= $obj.Title;
#'Descripcion'= $obj.Description;
#'Source'= $obj.Source.displayname;


} 

$CustomObjects = New-Object -TypeName Psobject -Property  $ReqProp
    
return $CustomObjects
  
}




<#

 $wi = "IR7326"

 $class = switch ($wi.substring(0, 2))

    {
        IR {$IRClass}
        SR {$SRClass}
        MA {$MAclass}
       
    }


$req = "IR7326"

$elemento = Get-SCSMObject -Class $irClass -Filter "Id -eq $req " -ComputerName "s1-dixx-ssm04"

get-actionLogFull $elemento "s1-dixx-ssm04"


#$fullLog  | select * | Out-GridView

    #$AnalystComments | select * | Out-GridView
#  
    
    
   #>