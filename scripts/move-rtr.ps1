function Move-Files {
  Param (
    [Parameter(Mandatory = $true)]
    $Origin,
    [Parameter(Mandatory = $true)]
    $Destiny,
    [Parameter(Mandatory = $true)]
    [string] $LibraryName
  )
  begin {}
  process {
    $library = Get-PnPList -Identity $LibraryName -Includes RootFolder.Folders -Connection $Origin

    foreach ($folder in $library.RootFolder.Folders) {
      if ($folder.Name -eq 'RTRFinal')
      {
        Get-PnPProperty -ClientObject $folder -Property Files

        # Create folder
        $queryFolder = "<View><Query><Where><Eq><FieldRef Name='Title'/><Value Type='Text'>$($folder.Name)</Value></Eq></Where></Query></View>"

        $destinationFolder = Get-PnPListItem -List $LibraryName -Query $queryFolder -Connection $Destiny
        if ($null -eq $destinationFolder) {
          # Create new folder
          $destinationFolder = Add-PnPFolder -Folder "$($LibraryName)" -Name $folder.Name -Connection $Destiny
        }

        #  Move files
        foreach ($file in $folder.Files) {
          $stream = $file[0].OpenBinaryStream()
          Invoke-PnPQuery -Connection $Origin
          Add-PnPFile -FileName $file.Name -Stream $stream.Value -Folder "$($LibraryName)/$($folder.Name)" -Connection $Destiny
        }
      }
    }
  }
  end {}
}

function Move-Items {
  [OutputType([System.Array])]
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
    [string] $ReportId
  )
  begin {
    Set-Variable query -value "<View Scope='RecursiveAll'><Query><Where><Eq><FieldRef Name='SFReportID'/><Value Type='Text'>$($ReportId)</Value></Eq></Where></Query></View>"
    $users = @{
      '40' = 137
      '43' = 134
      '36' = 133
    }
  }
  process {
    $ois = Get-PnPListItem -List $ListName -Query $query -PageSize 100 -Connection $Origin
    $nis = Get-PnPListItem -List $ListName -Query $query -PageSize 100 -Connection $Destiny

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
              $itemVal[$field] = $oListItem.FieldValues[$field]
            }
          }

          if ($field -eq 'RTRAuthorID' -or $field -eq 'RTRForwardTo') {
            $itemVal[$field] = $users["$($oListItem.FieldValues[$field])"]
          }
        }
      }

      if($null -ne $nListItem) {
        $null = Set-PnPListItem -List $ListName -Identity $nListItem.Id -Values $itemVal -Connection $Destiny
      } else {
        $null = Add-PnPListItem -List $ListName -Values $itemVal -Connection $Destiny
      }
    }

    if ($nListItems.Count -gt $oListItems.Count) {
      for ($j = $oListItems.Count; $j -lt $nListItems.Count; $j++) {
        $nListItem = $nListItems[$j]
        Remove-PnPListItem -List $ListName -Identity $nListItem.Id -Connection $Destiny
      }
    }
  }
  end { }
}

function Move-ItemsInFolder {
  [OutputType([System.Array])]
  Param (
    [Parameter(Mandatory = $true)]
    [AllowNull()]
    $Origin,
    [Parameter(Mandatory = $true)]
    [AllowNull()]
    $Destiny,
    [Parameter(Mandatory = $true)]
    [AllowNull()]
    [string] $ListName,
    [Parameter(Mandatory = $true)]
    [AllowNull()]
    $Fields,
    [Parameter(Mandatory = $true)]
    [AllowNull()]
    [string] $ReportId
  )
  begin {
    Set-Variable query -value "<View Scope='RecursiveAll'><Query><Where><Eq><FieldRef Name='SFReportID'/><Value Type='Text'>$($ReportId)</Value></Eq></Where></Query></View>"
  }
  process {

    $ois = Get-PnPListItem -List $ListName -Query $query -PageSize 100 -Connection $Origin
    $nis = Get-PnPListItem -List $ListName -Query $query -PageSize 100 -Connection $Destiny

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

      $folderUrl = [string] $oListItem.FieldValues['FileDirRef']
      $fName = $folderUrl.Substring($folderUrl.LastIndexOf('/') + 1)

      $null = Set-ListFolder -Origin $Origin -Destiny $Destiny -ListName $ListName -FolderName $fName
      
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
              $itemVal[$field] = $oListItem.FieldValues[$field]
            }
          }
        }
      }

      if($null -ne $nListItem) {
        $null = Set-PnPListItem -List $ListName -Identity $nListItem.Id -Values $itemVal -Connection $Destiny
      } else {
        $null = Add-PnPListItem -List $ListName -Folder $fName -Values $itemVal -Connection $Destiny
      }
    }

    if ($nListItems.Count -gt $oListItems.Count) {
      for ($j = $oListItems.Count; $j -lt $nListItems.Count; $j++) {
        $nListItem = $nListItems[$j]
        Remove-PnPListItem -List $NewList -Identity $nListItem.Id -Connection $Destiny
      }
    }

  }
  end { }
}

function Set-ListFolder {
  Param (
    [Parameter(Mandatory = $true)]
    [AllowNull()]
    $Origin,
    [Parameter(Mandatory = $true)]
    [AllowNull()]
    $Destiny,
    [Parameter(Mandatory = $true)]
    [AllowNull()]
    [string] $ListName,
    [Parameter(Mandatory = $true)]
    [AllowNull()]
    [string] $FolderName
  )
  begin {
    Set-Variable queryFolder -value "<View><Query><Where><Eq><FieldRef Name='Title'/><Value Type='Text'>$($FolderName)</Value></Eq></Where></Query></View>"
  }
  process {
    $folder = Get-PnPListItem -List $ListName -Query $queryFolder -Connection $Destiny
    if ($null -eq $folder) {
	    # Create new folder
      $folder = Add-PnPFolder -Folder "/Lists/$($ListName)" -Name $FolderName -Connection $Destiny
    }
    Write-Output $folder
  }
  end { }
}

function Move-RTR {
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

    Set-Variable reportList -value 'SFRTRReport'
    Set-Variable attachmentsList -value 'SFRTRAttachments'
    Set-Variable chargesList -value 'SFRTRCharges'
    Set-Variable employeesList -value 'SFRTREmployee'
    Set-Variable eventsList -value 'SFRTREvents'
    Set-Variable reviewList -value 'SFRTRReview'
    Set-Variable subjectList -value 'SFRTRSubject'
    Set-Variable witnessList -value 'SFRTRWitness'
    Set-Variable agenciesList -value 'SFRTRAgency'
    Set-Variable reportFields -value 'Title', 'SFReportID', 'SFReportStatus', 'SFRTRTypeOfIncident', 'SFRTRCaseNumber', 'RTRCaseNumber',
    'RTRDateAndTime', 'RTRLocation', 'RTRApartment', 'RTRNarrative', 'RTRDistricts', 'RTRReasonRTR', 'RTRCircumstancesAtScene',
    'RTRLightingConditions', 'RTRPossibleWeatherConditions', 'RTRWeatherConditions', 'RTRForwardTo', 'RTROwner',
    'RTRReviewStack', 'Created', 'Author'
    Set-Variable attachmentsFields -value 'Title', 'RTRDescription', 'RTRFileName', 'RTRPersonName', 'SFReportID',
    'SFTreatment', 'RTRTypeAttachment', 'RTRTypePerson', 'Created', 'Author'
    Set-Variable chargesFields -value 'Title', 'SFRTRCharges', 'SFFullNameSubject', 'SFReportID', 'RTRDetail',
    'Created', 'Author'
    Set-Variable employeesFields -value 'Title', 'RTRDistrict', 'RTRDutyStatus', 'RTRGender', 'RTRHeight', 'RTRWeight',
    'RTRRace', 'RTRRank', 'SFFullNameEmployee', 'SFReportID', 'RTRUnit', 'RTRWorkShift', 'SFLookupEmployee',
    'Created', 'Author'
    Set-Variable eventsFields -value 'Title', 'RTRApproxDistance', 'RTRBodyHits', 'RTRCanineAction', 'RTRCanineBite', 'RTRCanineDeployReason',
    'RTRCanineInjured', 'RTRCEWCatridge', 'RTRCEWSerial', 'RTRDeEscalation', 'RTRDeEscalationTechnique', 'RTRDeployReason',
    'RTRDetail', 'RTRDriveStun', 'RTREffective', 'RTREmployeeInjured', 'RTREventOrder', 'RTREventType', 'RTRFirearmAction', 'RTROfficerStatus',
    'RTRPartialHit', 'RTRProbesCollected', 'RTRReasonCEWNotEffective', 'SFEventSubtype', 'SFFullNameEmployee', 'SFFullNameSubject',
    'SFReportID', 'RTRSubjectInjured', 'RTRTypeOfChemicalAgent', 'RTRTypeOfFirearm', 'RTRTypeOfImmediateDanger', 'RTRTypeOfImpactWeapon', 'RTRTypeOfLessLethalWeapon',
    'RTRTypeOfPhysicalTechnique', 'RTRTypeOfResistance', 'RTRTypeOfSubjectBehavior', 'RTRVictimInjured', 'RTRWeaponDisplayed', 'Created', 'Author'
    Set-Variable reviewFields -value 'Title', 'RTRAction', 'RTRAuthorID', 'RTRAuthorName', 'RTRComment', 'RTRForwardTo', 'RTRLevel',
    'RTRReviewOrder', 'SFReportID', 'Created', 'Author'
    Set-Variable subjectFields -value 'Title', 'RTRAddress', 'RTRAppearance', 'RTRArrested', 'RTRCellphone', 'RTRContactMethod', 'RTRDOB', 'RTREmail',
    'RTRFirstName', 'RTRGender', 'RTRHeight', 'RTRWeight', 'RTRLastName', 'RTRMiddleName', 'RTRPhone', 'RTRRace', 'RTRApartment', 'RTRDistrict',
    'SFReportID', 'SFFullNameSubject', 'Created', 'Author'
    Set-Variable witnessFields -value 'Title', 'SFReportID', 'SFFullNameWitness', 'RTRAddress', 'RTRApartment', 'RTRArrested', 'RTRCellphone',
    'RTRContactMethod', 'RTRDOB', 'RTREmail', 'RTRFirstName', 'RTRGender', 'RTRLastName', 'RTRMiddleName', 'RTRPhone',
    'RTRWitnessType', 'RTRDistrict', 'RTRRace', 'Created', 'Author'
    Set-Variable agencyFields -value 'Title', 'SFReportID', 'RTRName', 'RTRAgencyORI', 'RTRCaseNumber', 'Created', 'Author'
    Set-Variable reportQuery -value "<View Scope='RecursiveAll'><Query><Where><And><Eq><FieldRef Name='FSObjType'/><Value Type='Integer'>0</Value></Eq><Eq><FieldRef Name='SFReportStatus'/><Value Type='Text'>Finished</Value></Eq></And></Where></Query></View>"
  }
  process {
    # Map Users
    $webOrigin = Get-PnPWeb -Includes SiteUsers -Connection $origin
    $users = @{}

    foreach ($user in $webOrigin.SiteUsers)
    {
      if ($user.Email) {
        $u = New-PnPUser -LoginName $user.LoginName -Connection $destiny
        
        $users.Add("$($user.ID)", $u.ID)
      }
    }
    
    Write-Progress -Activity 'Moving RTR Reports' -Id 1 -Status 'Getting Reports...' -PercentComplete 0
    $reports = Get-PnPListItem -List $reportList -Query $reportQuery -PageSize 100 -Connection $origin
    $rpp = 0;
    foreach ($report in $reports) {
      $reportId = $report.FieldValues['SFReportID']
      Write-Progress -Activity 'Moving RTR Reports' -Id 1 -Status "Processing Report $($reportId)" -PercentComplete ($rpp / $reports.Count * 100)
      Move-ItemsInFolder -Origin $origin -Destiny $destiny -ListName $reportList -Fields $reportFields -ReportId $reportId

      Write-Progress -Activity 'Report Entities' -Id 2 -ParentId 1 -Status "Moving Report Subjects..." -PercentComplete 5
      Move-ItemsInFolder -Origin $origin -Destiny $destiny -ListName $subjectList -Fields $subjectFields -ReportId $reportId
      
      Write-Progress -Activity 'Report Entities' -Id 2 -ParentId 1 -Status "Moving Report Employees..." -PercentComplete 20
      Move-ItemsInFolder -Origin $origin -Destiny $destiny -ListName $employeesList -Fields $employeesFields -ReportId $reportId

      Write-Progress -Activity 'Report Entities' -Id 2 -ParentId 1 -Status "Moving Report Agencies..." -PercentComplete 30
      Move-ItemsInFolder -Origin $origin -Destiny $destiny -ListName $agenciesList -Fields $agencyFields -ReportId $reportId
      
      Write-Progress -Activity 'Report Entities' -Id 2 -ParentId 1 -Status "Moving Report Witnesses..." -PercentComplete 40
      Move-ItemsInFolder -Origin $origin -Destiny $destiny -ListName $witnessList -Fields $witnessFields -ReportId $reportId
      
      Write-Progress -Activity 'Report Entities' -Id 2 -ParentId 1 -Status "Moving Report Charges..." -PercentComplete 50
      Move-ItemsInFolder -Origin $origin -Destiny $destiny -ListName $chargesList -Fields $chargesFields -ReportId $reportId
      
      Write-Progress -Activity 'Report Entities' -Id 2 -ParentId 1 -Status "Moving Report Events..." -PercentComplete 70
      Move-ItemsInFolder -Origin $origin -Destiny $destiny -ListName $eventsList -Fields $eventsFields -ReportId $reportId
      
      Write-Progress -Activity 'Report Entities' -Id 2 -ParentId 1 -Status "Moving Report Attachments..." -PercentComplete 80
      Move-ItemsInFolder -Origin $origin -Destiny $destiny -ListName $attachmentsList -Fields $attachmentsFields -ReportId $reportId
      
      Write-Progress -Activity 'Report Entities' -Id 2 -ParentId 1 -Status "Moving Report Reviews..." -PercentComplete 90
      Move-Items -Origin $origin -Destiny $destiny -ListName $reviewList -Fields $reviewFields -ReportId $reportId
      
      Write-Progress -Activity 'Report Entities' -Id 2 -ParentId 1 -Status "Completed" -Completed

      $rpp = $rpp + 1;
    }
    
    Move-Files -Origin $origin -Destiny $destiny -LibraryName 'SFRTRFinalReports'

    Write-Progress -Activity 'Moving RTR Reports' -Id 1 -Status 'Completed' -Completed
  }
  end { }
}

# Enter RTR url
$rtrEndpointOne = Read-Host "Site 1 Url:"
$rtrEndpointTwo = Read-Host "Site 2 Url:"
Move-RTR -OriginUrl $rtrEndpointOne -DestinyUrl $rtrEndpointTwo