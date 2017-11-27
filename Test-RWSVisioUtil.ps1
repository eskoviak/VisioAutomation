try {
    Remove-Module RWSVisioUtil
} catch {
    #Do Nothing - module wasn"t loaded
}

Import-Module .\RWSVisioUtil.psm1

Write-Host("TEST 00001:`nSetup:  Ensure no visio instances are running.  This calls`nGet-VisioInstance with no DocumentName.  Should get
prompted.  Enter anything, should get No Instance Loaded message") -ForegroundColor Blue
$instance = Get-VisioInstance