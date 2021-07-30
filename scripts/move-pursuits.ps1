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
        if ($null -ne $oListItem.FieldValues[$field]) {
          switch (($oListItem.FieldValues[$field].GetType()).Name) {
            'FieldUserValue' {
              $itemVal[$field] = $oListItem.FieldValues[$field].Email
            }
            'FieldLookupValue' {
              $itemVal[$field] = $oListItem.FieldValues[$field].LookupId
            }
            default {
              switch ($field) {
                'PursuitSecondaryUnit' {
                  if ($null -ne $oListItem.FieldValues['PursuitSecondaryUnit']) {                    
                    $secondaryUnit =  $oListItem.FieldValues['PursuitSecondaryUnit'] | ConvertFrom-Json

                    if ($secondaryUnit.PursuitDriverName) {
                      $secondaryUnit.PursuitDriverName.ID = $Users[$secondaryUnit.PursuitDriverName.ID]
                    }
                  }
                  
                  $itemVal['PursuitSecondaryUnit'] = ConvertTo-Json -InputObject $secondaryUnit -Depth 5 -Compress
                }
                'PursuitSupervision' {
                  if ($null -ne $oListItem.FieldValues['PursuitSupervision']) {
                    $supervisions = New-Object System.Collections.ArrayList

                    $supervisionEntries =  $oListItem.FieldValues['PursuitSupervision'] | ConvertFrom-Json
                    foreach ($entry in $supervisionEntries) {
                      if ($entry.User) {
                        $entry.User.ID = $Users[$entry.User.ID]
                      }
                      if ($entry.NextSupervisor) {
                        $entry.NextSupervisor.ID = $Users[$entry.NextSupervisor.ID]
                      }
                      $null = $supervisions.Add($entry)
                    }
                  }

                  $itemVal['PursuitSupervision'] = ConvertTo-Json -InputObject $supervisions -Depth 5 -Compress
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

function Move-Pursuits {
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

    Set-Variable reportsList -value 'Pursuits'
    Set-Variable reportFields -value 'Title', 'PursuitCaseNumber', 'PursuitReason', 'PursuitTimeStarted', 'PursuitTimeEnded', 'PursuitTotalTime', 'PursuitTermination',
    'PursuitTerminationOther', 'PursuitLength', 'PursuitTDDRelated', 'PursuitRoadBlock', 'PursuitExactRoute', 'PursuitSticksDeployed', 'PursuitSticksEffective', 'PursuitStartLocation', 'PursuitEndLocation',
    'PursuitTypeArea', 'PursuitTrafficConditions', 'PursuitRoadSurface', 'PursuitRoadConditions', 'PursuitWeatherConditions', 'PursuitDriverName', 'PursuitDispatchNotified', 'PursuitSupervisorNotified',
    'PursuitAuthorized', 'PursuitMonitoring', 'PursuitVehicleNumber', 'PursuitEmergencyLights', 'PursuitSirenActivated', 'PursuitTopSpeed', 'PursuitDriverInjured', 'PursuitDamagedVehicles',
    'PursuitDamagesDetails', 'PursuitPassengerName', 'PursuitPassengerInjured', 'PursuitPassengerInfo', 'PursuitSecondaryUnit', 'PursuitAgenciesInvolved',
    'PursuitSuspectName', 'PursuitSuspectAge', 'PursuitSuspectImpaired', 'PursuitSuspectVehicle', 'PursuitSuspectVehicleStolen', 'PursuitSuspectApprehended',
    'PursuitCharges', 'PursuitSuspectInCrash', 'PursuitSuspectInjured', 'PursuitSuspectPassengerInfo', 'PursuitOtherCriminalActivity', 'PursuitWitnesses',
    'PursuitNotes', 'PursuitCrashCaseNumber', 'PursuitSupervisor', 'PursuitSupervision', 'PursuitStatus', 'PursuitFollowedCrash', 'Author', 'Created'
  }
  process {
    # Map Users
    Write-Progress -Activity 'Moving Users' -Id 1 -Status "Moving Users..." -PercentComplete 0
    $webOrigin = Get-PnPWeb -Includes SiteUsers -Connection $origin
    $users = @{}
    
    foreach ($user in $webOrigin.SiteUsers)
    {
      if ($user.Email) {
        $u = New-PnPUser -LoginName $user.LoginName -Connection $destiny
        
        $users.Add($user.ID, $u.ID)
      }
    }
    Write-Progress -Activity 'Moving Users' -Id 1 -Status "Moving Users..." -PercentComplete 100 -Completed

    Write-Progress -Activity 'Moving Reports' -Id 3 -Status 'Getting Pursuits...' -PercentComplete 30
    $reports = Get-PnPListItem -List $reportsList -PageSize 100 -Connection $origin -Query "<View><Query><Where><Eq><FieldRef Name='PursuitStatus'/><Value Type='Text'>Completed</Value></Eq></Where></Query></View>"
    $rpp = 0;
    foreach ($report in $reports) {
      $reportId = [int]$report.FieldValues['ID']
      Write-Progress -Activity 'Moving Reports' -Id 3 -Status "Processing Report $($reportId)" -PercentComplete ($rpp / $reports.Count * 100)
      
      $query = "<View><Query><Where><Eq><FieldRef Name='ID'/><Value Type='Number'>$($reportId)</Value></Eq></Where></Query></View>"
      Move-Items -Origin $origin -Destiny $destiny -ListName $reportsList -Fields $reportFields -Query $query -Users $users -CopyAttachments

      $rpp = $rpp + 1;
    }
    Write-Progress -Activity 'Moving Reports' -Id 3 -Status 'Completed' -Completed
  }
  end { }
}

# Enter RTR url
$purEndpointOne = Read-Host "Origin Url:"
$purEndpointTwo = Read-Host "Destiny Url:"
Move-Pursuits -OriginUrl $purEndpointOne -DestinyUrl $purEndpointTwo