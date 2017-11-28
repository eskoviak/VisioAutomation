try {
    Remove-Module RWSVisioUtil
} catch {
    #Do Nothing - module wasn"t loaded
}

Import-Module .\RWSVisioUtil.psm1

Write-Host("Test 00001:  Test Visio Instance Detection") -ForegroundColor Yellow
$prompt = Read-Host ("(P)rocceed or (S)kip")
if($prompt.ToUpper() -eq "P") {
    Write-Host("TEST 00001a:`nSetup:  Ensure no visio instances are running.  This calls`nGet-VisioInstance with no DocumentName.") -ForegroundColor Blue
    Write-Host("Expected:`"No Visio Object Found`" message") -ForegroundColor Yellow
    $instance = Get-VisioInstance -GetAll
    if($error[0].CategoryInfo.Category -eq [System.Management.Automation.ErrorCategory]::ObjectNotFound) {
        Write-Host("PASS") -ForegroundColor Green
    } else {
        Write-Host("FAIL") -ForegroundColor Red
        exit -1
    }

    Write-Host("Test 00001b:`nSetup:  Start an instance of Visio.  When prompted to create a new drawing, select a blank drawing, CREATE and then close the window (CTRL-W).")  -ForegroundColor Blue
    $prompt = Read-Host -Prompt "Press enter to proceed"
    Write-Host("Expected:  Instance found with 0 documents loaded") -ForegroundColor Yellow
    $instance = Get-VisioInstance -GetAll
    Write-Host("Found a Visio Instance with {0} documents loaded" -f $instance.Documents.Count)
    if($instance.documents.Count -eq 0) {
    Write-Host("PASS") -ForegroundColor Green
    } else {
        Write-Host("FAIL") -ForegroundColor Red
        exit -1
    }

    Write-Host("Test 00001c: `nSetup:  In the running visio instance, open a file.")
    $prompt = Read-Host -Prompt "Enter the name of the visio file which you opened"
    #$instance = Get-VisioInstance -DocumentName $prompt



    
}

