function Move-Items {
  [OutputType([System.Object])]
  Param (
    [Parameter(Mandatory = $true)]
    $Origin,
    [Parameter(Mandatory = $true)]
    $Destiny,
    [Parameter(Mandatory = $true)]
    [string] $ListName,
    [Parameter(Mandatory = $true)]
    [AllowNull()]
    $Fields,
    [Parameter(Mandatory = $true)]
    [AllowNull()]
    [string] $Query,
    [Parameter(Mandatory = $false)]
    $Users,
    [Parameter(Mandatory = $false)]
    $Attachments,
    [Parameter(Mandatory = $false)]
    [Switch]$CopyAttachments
    
  )
  begin { }
  process {
    $ois = Get-PnPListItem -List $ListName -Query $Query -PageSize 100 -Connection $Origin
    $nis = Get-PnPListItem -List $ListName -Query $Query -PageSize 100 -Connection $Destiny

    $oListItems = New-Object System.Collections.ArrayList
    foreach($oi in $ois) {
      $null = $oListItems.Add($oi)
    }

    $nListItems = New-Object System.Collections.ArrayList
    foreach($ni in $nis) {
      $null = $nListItems.Add($ni)
    }

    # Copy Items
    for ($i = 0; $i -lt $oListItems.Count; $i++) {
      $oListItem = $oListItems[$i]
      $nListItem = $nListItems[$i]

      $itemVal = @{}
      foreach ($field in $Fields) {
        if ($oListItem.FieldValues[$field]) {
          switch (($oListItem.FieldValues[$field].GetType()).Name) {
            'FieldUserValue' {
              $itemVal[$field] = $oListItem.FieldValues[$field].Email
            }
            'FieldLookupValue' {
              $itemVal[$field] = $oListItem.FieldValues[$field].LookupId
            }
            default {
              switch ($field) {
                'CCEmployees' {
                  if ($null -ne $oListItem.FieldValues['CCEmployees']) {
                    $employees = New-Object System.Collections.ArrayList
                    
                    $employeesEntries =  $oListItem.FieldValues['CCEmployees'] | ConvertFrom-Json
                    foreach ($entry in $employeesEntries) {
                      if ($entry.CCEmployeeId) {
                        $entry.CCEmployeeId = $Users["$($entry.CCEmployeeId)"]
                      }
                      $null = $employees.Add($entry)
                    }
                  }
                  
                  $itemVal['CCEmployees'] = ConvertTo-Json -InputObject $employees -Depth 5 -Compress
                }
                'CCSupervision' {
                  if ($null -ne $oListItem.FieldValues['CCSupervision']) {
                    $supervisions = New-Object System.Collections.ArrayList

                    $supervisionEntries =  $oListItem.FieldValues['CCSupervision'] | ConvertFrom-Json
                    foreach ($entry in $supervisionEntries) {
                      if ($entry.CCSupervisorId) {
                        $entry.CCSupervisorId = $Users["$($entry.CCSupervisorId)"]
                      }
                      if ($entry.def_author) {
                        $entry.def_author = $Users["$($entry.def_author)"]
                      }
                      $null = $supervisions.Add($entry)
                    }
                  }

                  $itemVal['CCSupervision'] = ConvertTo-Json -InputObject $supervisions -Depth 5 -Compress
                }
                'CCSupervisionResults' {
                  if ($null -ne $oListItem.FieldValues['CCSupervisionResults']) {
                    $supervisionResults = New-Object System.Collections.ArrayList

                    $supervisionResultsEntries =  $oListItem.FieldValues['CCSupervisionResults'] | ConvertFrom-Json
                    foreach ($entry in $supervisionResultsEntries) {
                      if ($entry.CCSupervisorId) {
                        $entry.CCSupervisorId = $Users["$($entry.CCSupervisorId)"]
                      }
                      if ($entry.def_author) {
                        $entry.def_author = $Users["$($entry.def_author)"]
                      }
                      $null = $supervisionResults.Add($entry)
                    }
                  }

                  $itemVal['CCSupervisionResults'] = ConvertTo-Json -InputObject $supervisionResults -Depth 5 -Compress
                }
                'CCAttachments' {
                  if ($null -ne $oListItem.FieldValues['CCAttachments']) {
                    $attachs = New-Object System.Collections.ArrayList
                    
                    $attachmentsEntries =  $oListItem.FieldValues['CCAttachments'] | ConvertFrom-Json
                    foreach ($entry in $attachmentsEntries) {
                      if ($entry.def_id) {
                        $entry.def_id = $Attachments[$entry.def_id]
                      }
                      $null = $attachs.Add($entry)
                    }
                  }
                  
                  $itemVal['CCAttachments'] = ConvertTo-Json -InputObject $attachs -Depth 5 -Compress
                }
                default {
                  $itemVal[$field] = $oListItem.FieldValues[$field]
                }
              }
            }
          }
        }
      }

      if($null -ne $nListItem) {
        $null = Set-PnPListItem -List $ListName -Identity $nListItem.Id -Values $itemVal -Connection $Destiny
      } else {
        $nListItem = Add-PnPListItem -List $ListName -Values $itemVal -Connection $Destiny
        $null = $nListItems.Add($nListItem)
      }

      if ($CopyAttachments) {
        Add-Attachment -Origin $Origin -Destiny $Destiny -OItem $oListItem -DItem $nListItem
      }
    }

    if ($nListItems.Count -gt $oListItems.Count) {
      for ($j = $oListItems.Count; $j -lt $nListItems.Count; $j++) {
        $nListItem = $nListItems[$j]
        Remove-PnPListItem -List $NewList -Identity $nListItem.Id -Force -Recycle -Connection $Destiny
      }
    }

    $output = @{}

    for ($i = 0; $i -lt $oListItems.Count; $i++) {
      $oListItem = $oListItems[$i]
      $nListItem = $nListItems[$i]

      $output.Add($oListItem.Id, $nListItem.Id)
    }

    Write-Output $output
  }
  end { }
}

function Add-Attachment {
  Param (
    [Parameter(Mandatory = $true)]
    $Origin,
    [Parameter(Mandatory = $true)]
    $Destiny,
    [Parameter(Mandatory = $true)]
    $OItem,
    [Parameter(Mandatory = $true)]
    $DItem
  )
  begin {}
  process {
    $originAttachments = Get-PnPProperty -ClientObject $OItem -Property AttachmentFiles -Connection $Origin
    $destinationAttachments = Get-PnPProperty -ClientObject $DItem -Property AttachmentFiles -Connection $Destiny

    # Copy Attachments
    foreach($attachment in $originAttachments) {
      $destinationAttachment = $destinationAttachments | Where-Object { $_.FileName -eq $attachment.FileName }
      if ($null -eq $destinationAttachment) {
        # Add Attachment
        $file = Get-PnPFile -Url $attachment.ServerRelativeUrl -AsFileObject -Connection $Origin
        $stream = $file.OpenBinaryStream()
        Invoke-PnPQuery -Connection $Origin
        
        $attachmentCreation = New-Object Microsoft.SharePoint.Client.AttachmentCreationInformation
        $attachmentCreation.ContentStream = $stream.Value
        $attachmentCreation.FileName = $attachment.FileName
        $null = $DItem.AttachmentFiles.Add($attachmentCreation)
        
        Invoke-PnPQuery -Connection $Destiny
      }
    }

    $attachment = $null
    $destinationAttachments = Get-PnPProperty -ClientObject $DItem -Property AttachmentFiles -Connection $Destiny

    # Remove Attachments
    foreach($attachment in $destinationAttachments) {
      if ($attachment.FileName) {
        $originAttachment = $originAttachments | Where-Object { $_.FileName -eq $attachment.FileName }
        if ($null -eq $originAttachment) {
          # Remove Attachment
          Remove-PnPFile -ServerRelativeUrl $attachment.ServerRelativeUrl -Force -Connection $Destiny
        }
      }
    }
  }
  end { }
}

function Move-CC {
  Param (
    [Parameter(Mandatory = $true)]
    [String] $OriginUrl,
    [Parameter(Mandatory = $true)]
    [String] $DestinyUrl
  )
  begin {
    Try {
      $origin = Connect-PnPOnline -url $OriginUrl -UseWebLogin -ReturnConnection
      $destiny = Connect-PnPOnline -url $DestinyUrl -UseWebLogin -ReturnConnection
    }
    Catch {
      Write-Error -Exception $_.Exception -Message $_.Exception.Message
      Exit
    }

    Set-Variable attachmentsList -value 'Attachments'
    Set-Variable completedReportsList -value 'Completed Reports'
    Set-Variable investigationsTypesList -value 'Investigation Types'
    Set-Variable attachmentsFields -value 'Title', 'CCReport', 'CCDescription', 'CCAttachmentPrecedence', 'CCContributor',
    'CCContributorType', 'Created', 'Author'
    Set-Variable completedReportsFields -value 'Title', 'CCID', 'CCCaseNumber', 'CCReportLinked', 'CCReportLinkedTo', 'CCReportType', 'CCDateTime',
    'CCStreetAddress', 'CCCardinal', 'CCStreetName', 'CCStreetType', 'CCApartment', 'CCCity', 'CCState', 'CCZIPCode', 'CCDistrict',
    'CCIncidentNarrative', 'CCLocationType', 'CCOtherLocation', 'CCDesiredOutcome', 'CCResult', 'CCOfficerCompletingReport', 'CCCaseReportUrl', 'CCCaseNumberLinked',
    'CCEmployees', 'CCPersons', 'CCSupervision', 'CCSupervisionResults', 'CCAttachments', 'CCComplaintInvestigators', 'CCResultsReportType', 'CCResultsCaseNumber',
    'CCResultsCaseNumberLinked', 'CCResultsCaseNumberLinkedTo', 'CCResultsSubjectsRaces', 'CCResultsSubjectsGenders', 'CCResultsThirdParty', 'CCResultsKnownEmployees',
    'CCResultsUnknownEmployees', 'CCResultsEmployees', 'CCResultsEmployeesDistricts', 'CCResultsEmployeesUnits', 'CCResultsEmployeesRaces', 'CCResultsSupervisors',
    'CCResultsImproperUOF', 'CCResultsInvestigationType', 'CCResultsFileNumber', 'CCResultsFormalReportCompleted', 'CCResultsInvestigators', 'CCResultsInvestigationResult',
    'CCResultsInternalInvestigation', 'CCResultsCompletedBy', 'CCReporterType', 'Author', 'Created'
    Set-Variable investigationsTypesFields -value 'Title', 'Created', 'Author'
  }
  process {
    # Map Users
    Write-Progress -Activity 'Moving C&C Users' -Id 1 -Status "Moving Users..." -PercentComplete 0
    $webOrigin = Get-PnPWeb -Includes SiteUsers -Connection $origin
    $users = @{}
    
    foreach ($user in $webOrigin.SiteUsers)
    {
      if ($user.Email) {
        $u = New-PnPUser -LoginName $user.LoginName -Connection $destiny
        
        $users.Add("$($user.ID)", $u.ID)
      }
    }
    Write-Progress -Activity 'Moving C&C Users' -Id 1 -Status "Moving Users..." -PercentComplete 100 -Completed

    $query = "<View></View>"
    Write-Progress -Activity 'Moving C&C Configuration Lists' -Id 2 -Status "Moving Investigation Types..." -PercentComplete 0
    Move-Items -Origin $origin -Destiny $destiny -ListName $investigationsTypesList -Fields $investigationsTypesFields -Query $query -Users $users
    Write-Progress -Activity 'Moving C&C Configuration Lists' -Id 2 -Status "Moving Investigation Types..." -PercentComplete 100 -Completed

    Write-Progress -Activity 'Moving C&C Completed Reports' -Id 3 -Status 'Getting Completed Reports...' -PercentComplete 30
    $reports = Get-PnPListItem -List $completedReportsList -PageSize 100 -Connection $origin
    $rpp = 0;
    foreach ($report in $reports) {
      $reportId = [int]$report.FieldValues['CCID']
      Write-Progress -Activity 'Moving C&C Completed Reports' -Id 3 -Status "Processing Report $($reportId)" -PercentComplete ($rpp / $reports.Count * 100)
      
      $query = "<View><Query><Where><Eq><FieldRef Name='CCContributor'/><Value Type='Text'>$($reportId)</Value></Eq></Where></Query></View>"
      Write-Progress -Activity 'Report Entities' -Id 4 -ParentId 3 -Status "Moving Attachments..." -PercentComplete 30
      $attachments = Move-Items -Origin $origin -Destiny $destiny -ListName $attachmentsList -Fields $attachmentsFields -Query $query -CopyAttachments
      
      Write-Progress -Activity 'Report Entities' -Id 4 -ParentId 3 -Status "Moving Report..." -PercentComplete 70
      $query = "<View><Query><Where><Eq><FieldRef Name='CCID'/><Value Type='Text'>$($reportId)</Value></Eq></Where></Query></View>"
      Move-Items -Origin $origin -Destiny $destiny -ListName $completedReportsList -Fields $completedReportsFields -Query $query -Users $users -Attachments $attachments

      Write-Progress -Activity 'Report Entities' -Id 4 -ParentId 3 -Status "Completed" -Completed
      $rpp = $rpp + 1;
    }
    Write-Progress -Activity 'Moving C&C Reports' -Id 3 -Status 'Completed' -Completed
  }
  end { }
}

# Enter RTR url
$ccEndpointOne = Read-Host "Origin Url:"
$ccEndpointTwo = Read-Host "Destiny Url:"
Move-CC -OriginUrl $ccEndpointOne -DestinyUrl $ccEndpointTwo