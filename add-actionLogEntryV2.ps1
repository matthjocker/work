Import-Module SMLETS

function Add-ActionLogEntry {
    param (
        [parameter(Mandatory=$true, Position=0)]
        $ClassName,
        [parameter(Mandatory=$true, Position=1)]
        [ValidateSet("Assign","AnalystComment","Closed","Escalated","EmailSent","EndUserComment","FileAttached","FileDeleted","Reactivate","Resolved","TemplateApplied")]
        [string] $Action,
        [parameter(Mandatory=$true, Position=2)]
        [string] $Comment,
        [parameter(Mandatory=$true, Position=3)]
        [string] $EnteredBy,
        [parameter(Mandatory=$false, Position=4)]
        [Nullable[boolean]] $IsPrivate = $false,
        [parameter(Mandatory=$false, Position=5)]
        [DateTime] $EnteredDate,
        [Parameter( Mandatory = $true , Position=6)]
        [string]$server
    )


    #Choose the Action Log Entry to be created. Depending on the Action Log being used, the $propDescriptionComment Property could be either Comment or Description
    switch ($Action)
    {
        Assign {$CommentClass = "System.WorkItem.TroubleTicket.ActionLog"; $ActionType = "System.WorkItem.ActionLogEnum.RecordAssigned"; $ActionEnum = get-scsmenumeration $ActionType; $propDescriptionComment = "Description"}
        AnalystComment {$CommentClass = "System.WorkItem.TroubleTicket.AnalystCommentLog"; $propDescriptionComment = "Comment"}
        Closed {$CommentClass = "System.WorkItem.TroubleTicket.ActionLog"; $ActionType = "System.WorkItem.ActionLogEnum.RecordClosed"; $ActionEnum = get-scsmenumeration $ActionType; $propDescriptionComment = "Description"}
        Escalated {$CommentClass = "System.WorkItem.TroubleTicket.ActionLog"; $ActionType = "System.WorkItem.ActionLogEnum.RecordEscalated"; $ActionEnum = get-scsmenumeration $ActionType; $propDescriptionComment = "Description"}
        EmailSent {$CommentClass = "System.WorkItem.TroubleTicket.ActionLog"; $ActionType = "System.WorkItem.ActionLogEnum.EmailSent"; $ActionEnum = get-scsmenumeration $ActionType; $propDescriptionComment = "Description"}
        EndUserComment {$CommentClass = "System.WorkItem.TroubleTicket.UserCommentLog"; $propDescriptionComment = "Comment"}
        FileAttached {$CommentClass = "System.WorkItem.TroubleTicket.ActionLog"; $ActionType = "System.WorkItem.ActionLogEnum.FileAttached"; $ActionEnum = get-scsmenumeration $ActionType; $propDescriptionComment = "Description"}
        FileDeleted {$CommentClass = "System.WorkItem.TroubleTicket.ActionLog"; $ActionType = "System.WorkItem.ActionLogEnum.FileDeleted"; $ActionEnum = get-scsmenumeration $ActionType; $propDescriptionComment = "Description"}
        Reactivate {$CommentClass = "System.WorkItem.TroubleTicket.ActionLog"; $ActionType = "System.WorkItem.ActionLogEnum.RecordReopened"; $ActionEnum = get-scsmenumeration $ActionType; $propDescriptionComment = "Description"}
        Resolved {$CommentClass = "System.WorkItem.TroubleTicket.ActionLog"; $ActionType = "System.WorkItem.ActionLogEnum.RecordResolved"; $ActionEnum = get-scsmenumeration $ActionType; $propDescriptionComment = "Description"}
        TemplateApplied {$CommentClass = "System.WorkItem.TroubleTicket.ActionLog"; $ActionType = "System.WorkItem.ActionLogEnum.TemplateApplied"; $ActionEnum = get-scsmenumeration $ActionType; $propDescriptionComment = "Description"}
    }
    #Alias on Type Projection for Service Requests and Problem and are singular, whereas Incident and Change Request are plural. Update $CommentClassName
    if (($ClassName -eq "System.WorkItem.Problem") -or ($ClassName -eq "System.WorkItem.ServiceRequest")) {$CommentClassName = "ActionLog"} else {$CommentClassName = "ActionLogs"}
    #Analyst and End User Comments Classes have different Names based on the Work Item class
    if ($Action -eq "AnalystComment")
    {
        switch ($ClassName)
        {
            "System.WorkItem.Incident" {$CommentClassName = "AnalystComments"}
            "System.WorkItem.ServiceRequest" {$CommentClassName = "AnalystCommentLog"}
            "System.WorkItem.Problem" {$CommentClassName = "Comment"}
            "System.WorkItem.ChangeRequest" {$CommentClassName = "AnalystComments"}
        }
    }
    if ($Action -eq "EndUserComment")
    {
        switch ($ClassName)
        {
            "System.WorkItem.Incident" {$CommentClassName = "UserComments"}
            "System.WorkItem.ServiceRequest" {$CommentClassName = "EndUserCommentLog"}
            "System.WorkItem.Problem" {$CommentClass = "System.WorkItem.TroubleTicket.AnalystCommentLog"; $CommentClassName = "Comment"}
            "System.WorkItem.ChangeRequest" {$CommentClassName = "UserComments"}
        }
    }
    # Generate a new GUID for the Action Log entry
    $NewGUID = ([guid]::NewGuid()).ToString()
    # Create the object projection with properties
    $Projection = @{__CLASS = $CommentClass;
                    __OBJECT = @{Id = $NewGUID;
                        DisplayName = $NewGUID;
                        ActionType = $ActionType;
                        $propDescriptionComment = $Comment;
                        Title = "$($ActionEnum.DisplayName)";
                        EnteredBy  = $EnteredBy;
                        EnteredDate = $EnteredDate
                        IsPrivate = $IsPrivate;
                    }
                    
    }


    #Create the projection based on the work item class
    <#
    switch ($WIObject.ClassName)
    {
        "System.WorkItem.Incident" {New-SCSMObjectProjection -Type "System.WorkItem.IncidentPortalProjection$" -Projection $Projection  -ComputerName $server }
        "System.WorkItem.ServiceRequest" {New-SCSMObjectProjection -Type "System.WorkItem.ServiceRequestProjection$" -Projection $Projection  -ComputerName $server }
        "System.WorkItem.Problem" {New-SCSMObjectProjection -Type "System.WorkItem.Problem.ProjectionType$" -Projection $Projection  -ComputerName $server }
        "System.WorkItem.ChangeRequest" {New-SCSMObjectProjection -Type "Cireson.ChangeRequest.ViewModel$" -Projection $Projection  -ComputerName $server }
    }

    #>


      switch ($ClassName)
        {
            "System.WorkItem.Incident" {return $Projection}
            "System.WorkItem.ServiceRequest" { return $Projection}
           
        }
   
}


function Add-OperationalActionLogEntry {
    param (
        [parameter(Mandatory=$true, Position=0)]
        $ActionLog,
        [parameter(Mandatory=$true, Position=1)]
        $servidor
    )


    $NewGUID = ([guid]::NewGuid()).ToString()
    $ActionType = $ActionLog.ActionType;
    $ActionEnum = get-scsmenumeration $ActionType -ComputerName $servidor
    $Projection = @{__CLASS = "System.WorkItem.TroubleTicket.ActionLog";
                        __OBJECT = @{Id = $NewGUID;
                            DisplayName = $NewGUID;
                            ActionType =  $ActionType
                            Description = $ActionLog.Description;
                            Title = "$($ActionEnum.DisplayName)";
                            EnteredBy  = $ActionLog.EnteredBy;
                            EnteredDate = $ActionLog.EnteredDate
                            IsPrivate = $ActionLog.IsPrivate;
                        }
    
}
    return $Projection
}
<#

Import-Module SMLETS


$server = "s1-hixx-ssm01"

$IRClass = Get-SCSMClass -Name System.WorkItem.Incident$ -ComputerName $server


$ir = Get-SCSMObject -Class $IRClass -Filter "id -eq IR3074" -ComputerName $server

Add-ActionLogEntry -WIObject $ir -Action "AnalystComment" -Comment "Insertando comentario de prueba desde powershell - 2" -EnteredBy "meaguirre" -IsPrivate $false -server $server

#>