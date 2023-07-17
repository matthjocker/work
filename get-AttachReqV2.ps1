#$CRID = "CR1306"
Import-Module SMLets 

function get-AttachReq {
param(
  [Parameter(Mandatory=$true)]
  [String]$servidor,

  [Parameter(Mandatory=$true)]
  [string[]]
  $wi,

  [Parameter(Mandatory=$true)]
  [string]$OutputFolder
  

)

#setup static variables for later use
# $servidor = "s1-dixx-ssm04"
# $OutputFolder =  "C:\temp\reqExport"
  
  
$srClass = Get-SCSMClass -Name system.workitem.servicerequest$ -ComputerName $servidor

$irClass = Get-SCSMClass -name System.Workitem.Incident$ -ComputerName $servidor


#$fileClass = Get-SCSMClass System.FileAttachment$

#$ContainsActivity = get-scsmrelationshipclass System.WorkItemContainsActivity$

$AttachedFileRel = get-scsmrelationshipclass System.WorkItemHasFileAttachment$ -ComputerName $servidor

#$OutputFolder = "C:\temp\reqExport\"



#id = "SR7157"
#$id = "SR7846"

#$id = "SR7670"

#Get targeted CR

# $classObj = "IR7437"

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

 # get-AttachReq -wi "IR7437" -OutputFolder "C:\temp\reqExport\" -servidor "s1-dixx-ssm04"