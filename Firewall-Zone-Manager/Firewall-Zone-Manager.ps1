param ($zones, $type, $profileType = "any", $interfaceType = "any", [Switch] $delete)

function CreateInboundRule {
    param($ruleName, $textRanges, $description, $profileType, $interfaceType, $start, $end)  
    Write-Host "Creating inbound rule '$ruleName' for ranges $start to $end."
    
    netsh.exe advfirewall firewall add rule name="$rulename" dir=in action=block localip=any remoteip="$textRanges" description="$description" profile="$profileType" interfacetype="$interfaceType"
    
    if (-not $?) {
        Write-Host "Failed to create inbound rule '$rulename'..."
    }
}

function CreateOutboundRule {
    param($ruleName, $textRanges, $description, $profileType, $interfaceType, $start, $end)
    Write-Host "Creating outbound rule '$ruleName' for ranges $start to $end."    
    
    netsh.exe advfirewall firewall add rule name="$rulename" dir=out action=block localip=any remoteip="$textRanges" description="$description" profile="$profileType" interfacetype="$interfaceType"
    
    if (-not $?) {
        Write-Host "Failed to create outbound rule '$rulename'..."
    }
}

function DeleteRules {
    param($zone)

    $existingRules = netsh.exe advfirewall firewall show rule name=all | select-string '^[Rule Name]+:\s+(.+$)' | ForEach-Object { $_.matches[0].groups[1].value } | Get-Unique
  
    if ($existingRules.count -lt 3) {
        Write-Host "Could not get a list of the existing rules." ; 
        
        return;
    } 

    foreach ($rule in $existingRules) {
        if ($rule -match "$zone-\d\d\d") { 
            Write-Host "Deleting inbound and outbound rule '$rule'."
            netsh.exe advfirewall firewall delete rule name=$rule
        }
    }

    Write-Host "Deleted existing rules for zone '$zone'."
}
function AddRules {
    param($zone, $ranges, $type, $profileType, $interfaceType)
    
    $rangesCount = $ranges.count

    # Better performance when adding only 100 ranges per rule
    $maxRangesPerRule = 100
    
    $ruleNumber = 1
    $start = 1
    $end = $maxRangesPerRule

    while ($start -le $rangesCount) {   
        $ruleName = "$zone-$($ruleNumber.ToString().PadLeft(3, '0'))"

        if ($end -gt $rangesCount) {
            $end = $rangesCount 
        } 

        $textRanges = [System.String]::Join(",", $($ranges[$($start - 1)..$($end - 1)])) 

        $description = "Automatically generated rule to blouck zone '$zone'. See '' for more info."

        switch ($type.ToLower()) {
            "inbound" {
                CreateInboundRule -ruleName $ruleName -textRanges $textRanges -description $description -profileType $profileType -interfaceType $interfaceType -start $start -end $end
                
                break;
            }
            "outbound" {
                CreateOutboundRule -ruleName $ruleName -textRanges $textRanges -description $description -profileType $profileType -interfaceType $interfaceType -start $start -end $end
                
                break;
            }
            "both" {
                CreateInboundRule -ruleName $ruleName -textRanges $textRanges -description $description -profileType $profileType -interfaceType $interfaceType -start $start -end $end
                CreateOutboundRule -ruleName $ruleName -textRanges $textRanges -description $description -profileType $profileType -interfaceType $interfaceType -start $start -end $end

                break;
            }
            default {
                CreateInboundRule -ruleName $ruleName -textRanges $textRanges -description $description -profileType $profileType -interfaceType $interfaceType -start $start -end $end
                CreateOutboundRule -ruleName $ruleName -textRanges $textRanges -description $description -profileType $profileType -interfaceType $interfaceType -start $start -end $end
            }
        }

        $ruleNumber++
        $start += $maxRangesPerRule
        $end += $maxRangesPerRule
    }  
}

function ProcessZone {
    param($zoneFile, $type, $profileType = "any", $interfaceType = "any", $delete)    
  
    $zone = [IO.Path]::GetFileNameWithoutExtension($zoneFile)

    Write-Host "Processing zone '$zone'."
   
    DeleteRules -zone $zone

    if ($delete) {
        return;       
    }
   
    $ranges = Get-Content $zoneFile | Where-Object { ($_.trim().length -ne 0) -and ($_ -match '^[0-9a-f]{1,4}[\.\:]') } 
    
    if (!$ranges) {
        Write-Host "No IP addresses to block."

        return;
    }
      
    AddRules -zone $zone -ranges $ranges -type $type -profileType $profileType -interfaceType $interfaceType
}

while (!$zones) {
    Write-Host "Please indicate the zones with the '-zones' parameter, values can be: 'all', only one zone name e.g. 'be' or a comma-seperated list with multiple zone names e.g. 'be,nl,fr,de'. To list available zones type 'list'."; 
    
    $zones = Read-Host -Prompt 'Zones'

    while ($zones -like "list") {
        $files = Get-ChildItem ".\zones\*.zone"

        foreach ($file in $files) {
            $zone = [IO.Path]::GetFileNameWithoutExtension($file)
           
            Write-Host $zone
        }

        $zones = Read-Host -Prompt 'Zones'
    }    
}

while (!$delete -and (!$type -or ($type -notlike "inbound" -and $type -notlike "outbound" -and $type -notlike "both"))) {
    Write-Host "Please indicate the type with the '-type' parameter, values can be: 'inbound', 'outbound' or 'both'."; 
    
    $type = Read-Host -Prompt 'Type'
}

$action = "create"

if ($delete) {
    $action = "delete"
    $type = "both"
}

switch ($type.ToLower()) { 

    "inbound" {
        Write-Host "This script will $action inbound rules for '$zones'."
        break;
    }
    "outbound" {
        Write-Host "This script will $action outbound rules for '$zones'."
        break;
    }
    "both" {
        Write-Host "This script will $action inbound and outbound rules for '$zones'."
        break;
    }
    default {
        Write-Host "This script will $action inbound and outbound rules for '$zones'."        
    }
}    

if ($zones -like "all") {
    $zoneFiles = Get-ChildItem ".\zones\*.zone"

    foreach ($zoneFile in $zoneFiles) {
        ProcessZone -zoneFile $zoneFile -type $type -profileType $profileType -interfaceType $interfaceType -delete $delete
    }    
}
else {
    $zonesArray = $zones.Split(",")

    foreach ($zone in $zonesArray) {
       
        $formattedZone = $zone.ToLower().Trim()
        
        $zoneFile = Get-Item ".\zones\$formattedZone.zone" -ErrorAction SilentlyContinue
        
        if (!$zoneFile) {
            Write-Host "The zone '$formattedZone' does not exist, skipping this one..."            
        }
        else {
            ProcessZone -zoneFile $zoneFile -type $type -profileType $profileType -interfaceType $interfaceType -delete $delete
        }
    }
}