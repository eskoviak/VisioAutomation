try {
    Remove-Module RWSVisioUtil
} catch {
    #Do Nothing - module wasn"t loaded
}

Import-Module .\RWSVisioUtil.psm1

Write-Host("Test 00001:  Test Visio Instance Detection") -ForegroundColor Yellow
$prompt = Read-Host ("(P)rocceed or (S)kip")
if($prompt.ToUpper() -eq "P") {
    Write-Host("TEST 00001a:`nSetup:  Ensure no visio instances are running.  This calls`nGet-VisioInstance with no DocumentName.  Should get
    prompted.  Enter anything, should get `"No Visio Object Found`" message") -ForegroundColor Blue
    $instance = Get-VisioInstance
    
    Write-Host("Test 00001b:`nSetup:  Start an instance of Visio.  When prompted to create a new drawing, select a blank drawing, create and then close the window (CTRL-W).`n") -ForegroundColor Blue
    $prompt = Read-Host -Prompt "Press enter to proceed"
    $instance = Get-VisioInstance -DocumentName DummyString
    Write-Host("Found a Visio Instance with {0} documents loaded" -f $instance.Documents.Count)
    
}

