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
                'CrashSupervision' {
                  if ($null -ne $oListItem.FieldValues['CrashSupervision']) {
                    $supervisions = New-Object System.Collections.ArrayList

                    $supervisionEntries =  $oListItem.FieldValues['CrashSupervision'] | ConvertFrom-Json
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

                  $itemVal['CrashSupervision'] = ConvertTo-Json -InputObject $supervisions -Depth 5 -Compress
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

function Move-Crashes {
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

    Set-Variable reportsList -value 'Crashes'
    Set-Variable reportFields -value 'Title', 'CrashCaseNumber', 'CrashPursuitCase', 'CrashDateOfIncident', 'CrashSpeedAtImpact', 'CrashLocation', 'CrashTypeOfArea',
    'CrashTrafficConditions', 'CrashRoadSurface', 'CrashRoadConditions', 'CrashWeatherConditions', 'CrashDriverName', 'CrashDriverStatus', 'CrashDriverInjured',
    'CrashDriverReceivedTreatment', 'CrashPassengersInjured', 'CrashContactInfoPassengers', 'CrashPassengersTreatment', 'CrashAgenciesInvolved', 'CrashOtherAgencies',
    'CrashOfficersOtherAgenciesInjure', 'CrashOfficersOtherAgTreatment', 'CrashBystandersInjured', 'CrashBystandersReceivedTreatment', 'CrashWitnesses',
    'CrashPropertyDamaged', 'CrashDescribePropertyDamaged', 'CrashCostPrivatePropertyDamaged', 'CrashCostPublicPropertyDamaged', 'CrashCostAgencyPropertyDamaged',
    'CrashPropertyOwnerName', 'CrashPropertyOwnerPhone', 'CrashNotes', 'CrashSupervisor', 'CrashSupervision', 'CrashStatus', 'Author', 'Created'
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

    Write-Progress -Activity 'Moving Reports' -Id 3 -Status 'Getting Crashes...' -PercentComplete 30
    $reports = Get-PnPListItem -List $reportsList -PageSize 100 -Connection $origin -Query "<View><Query><Where><Eq><FieldRef Name='CrashStatus'/><Value Type='Text'>Completed</Value></Eq></Where></Query></View>"
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
$crashEndpointOne = Read-Host "Origin Url:"
$crashEndpointTwo = Read-Host "Destiny Url:"
Move-Crashes -OriginUrl $crashEndpointOne -DestinyUrl $crashEndpointTwo