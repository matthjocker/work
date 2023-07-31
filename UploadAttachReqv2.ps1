import-module SMlets

Function Insert-Attachment{

    [CmdletBinding( SupportsShouldProcess=$false )]
    Param(
        [Parameter( Mandatory = $true )]
        [string]$SCSMID,
        [Parameter( Mandatory = $true )]
        [string]$Directory,
        [Parameter( Mandatory = $true )]
        [string]$server,
        [Parameter( Mandatory = $true )]
        [string]$tipoClase,
        [string]$EnteredBy = $cred.UserName
      
    )
    #Get management group
    $ManagementGroup = New-Object Microsoft.EnterpriseManagement.EnterpriseManagementGroup $server
    #Init classes
    $FileAttachmentRel = Get-SCSMRelationshipClass -Name System.WorkItemHasFileAttachment$ -ComputerName $server
    $FileAttachmentClass = Get-SCSMClass -Name System.FileAttachment$  -ComputerName $server
    $ActionLogClass = Get-SCSMClass -Name System.WorkItem.TroubleTicket.ActionLog$  -ComputerName $server
    $ActionLogRel = Get-SCSMRelationshipClass -Name System.WorkItemHasActionLog$  -ComputerName $server

    #Get the file listing for the directory
    $AllFiles = Get-ChildItem $Directory

    #Check how many files were in the directory
    #Also check for any empty files?
    Write-Host  $AllFiles.Count
    Foreach ( $FileObject in $AllFiles )
    {
        try{
        Write-Host "pase"
        #Create a filestream 
        $FileMode = [System.IO.FileMode]::Open
        $fRead = New-Object System.IO.FileStream $FileObject.FullName, $FileMode

        #Create file object to be inserted
        $NewFileAttach = New-Object Microsoft.EnterpriseManagement.Common.CreatableEnterpriseManagementObject($ManagementGroup, $FileAttachmentClass)
        #Populate properties with info
        $SCSMGUID_Attachment = [Guid]::NewGuid().ToString()
        $NewFileAttach.Item($FileAttachmentClass, "Id").Value = $SCSMGUID_Attachment
        $NewFileAttach.Item($FileAttachmentClass, "DisplayName").Value = $FileObject.Name
        $NewFileAttach.Item($FileAttachmentClass, "Description").Value = $FileObject.Name
        $NewFileAttach.Item($FileAttachmentClass, "Extension").Value = $FileObject.Extension
        $NewFileAttach.Item($FileAttachmentClass, "Size").Value = $FileObject.Length
        $NewFileAttach.Item($FileAttachmentClass, "AddedDate").Value = [DateTime]::Now.ToUniversalTime()
        $NewFileAttach.Item($FileAttachmentClass, "Content").Value = $fRead

        #cambia la projection segun la clase del requerimiento
        
        switch ( $tipoClase )
            {
     
                ir {  $ProjectionType = Get-SCSMTypeProjection -Name System.WorkItem.IncidentPortalProjection$ -ComputerName $server  }
                sr {  $ProjectionType = Get-SCSMTypeProjection -Name System.WorkItem.ServiceRequestProjection$ -ComputerName $server   }
  
            }

           # $SCSMID = "ir2714"
           # $server = "s1-hixx-ssm01"
          
        $Projection = Get-SCSMObjectProjection -ProjectionName $ProjectionType.Name -Filter "ID -eq $SCSMID" -ComputerName $server 
        
        #Attach file object to Service Manager
        $Projection.__base.Add($NewFileAttach, $FileAttachmentRel.Target)
        $Projection.__base.Commit()

        # $SCSMGUID_ActionLog = [Guid]::NewGuid().ToString()
        # $MP = Get-SCManagementPack -Name "System.WorkItem.Library" -ComputerName $server
        # $ActionType = "System.WorkItem.ActionLogEnum.FileAttached"
        # $NewLog = New-Object Microsoft.EnterpriseManagement.Common.CreatableEnterpriseManagementObject($ManagementGroup, $ActionLogClass)

        # $NewLog.Item( $ActionLogClass, "Id").Value = $SCSMGUID_ActionLog
        # $NewLog.Item( $ActionLogClass, "DisplayName").Value = $SCSMGUID_ActionLog
        # $NewLog.Item( $ActionLogClass, "ActionType").Value = $MP.GetEnumerations().GetItem($ActionType)
        # $NewLog.Item( $ActionLogClass, "Title").Value = "Attached File"
        # $NewLog.Item( $ActionLogClass, "EnteredBy").Value = $EnteredBy
        # $NewLog.Item( $ActionLogClass, "Description").Value = $FileObject.Name
        # $NewLog.Item( $ActionLogClass, "EnteredDate").Value = (Get-Date).ToUniversalTime()

        # #Insert comment to action log
        # $Projection.__base.Add($NewLog, $ActionLogRel.Target)
        # $Projection.__base.Commit()

        #Cleanup
        $fRead.Close();
    }  catch {
        $fRead.Close();
    }
    }
}


<#
$iRupload = "ir7437"

$iRdestino = "ir2714"

$basePath = "C:\temp\reqExport\"

  $FullDirPath = $basePath + $iRupload + "\";

    $AttachmentEntries = [IO.Directory]::GetFiles($FullDirPath); 
    $classObj = "ir"
    $servidorDestino = "s1-hixx-ssm01"

Insert-Attachment -SCSMID $iRdestino -Directory $AttachmentEntries[0] -tipoClase $classObj -server $servidorDestino
 #>                                                            