# Functions
Function InstallOfficeExtension() {
    <# .DESCRIPTION Adds a Google Chrome extension to the forced install list. Can be used for forcing installation of any Google Chrome extension. Takes existing extensions into account which might be added by other means, such as GPO and MDM. #>
param(
    [string]$extensionId = 'gbkeegbaiigmenfmjfclcdgdpimamgkj',
    [switch]$info
)
if($info){
    $InformationPreference = "Continue"
}
if(!($extensionId)){
    # Empty Extension
    $result = "No Extension ID"
}
else{
    Write-Information "ExtensionID = $extensionID"
    $extensionId = "$extensionId;https://clients2.google.com/service/update2/crx"
    $regKey = "HKLM:\SOFTWARE\Policies\Google\Chrome\ExtensionInstallForcelist"
    if(!(Test-Path $regKey)){
        New-Item $regKey -Force
        Write-Information "Created Reg Key $regKey"
    }
    # Add Extension to Chrome
    $extensionsList = New-Object System.Collections.ArrayList
    $number = 0
    $noMore = 0
    do{
        $number++
        Write-Information "Pass : $number"
        try{
            $install = Get-ItemProperty $regKey -name $number -ErrorAction Stop
            $extensionObj = [PSCustomObject]@{
                Name = $number
                Value = $install.$number
            }
            $extensionsList.add($extensionObj) | Out-Null
            Write-Information "Extension List Item : $($extensionObj.name) / $($extensionObj.value)"
        }
        catch{
            $noMore = 1
        }
    }
    until($noMore -eq 1)
    $extensionCheck = $extensionsList | Where-Object {$_.Value -eq $extensionId}
    if($extensionCheck){
        $result = "Extension Already Exists"
        Write-Information "Extension Already Exists"
    }else{
        $newExtensionId = $extensionsList[-1].name + 1
        New-ItemProperty HKLM:\SOFTWARE\Policies\Google\Chrome\ExtensionInstallForcelist -PropertyType String -Name $newExtensionId -Value $extensionId
        $result = "Installed"
    }
}
$result
}

Function AssociateFileTypes() {
$protocols = 'htm','html','shtml','svg','xht','xhtml','ftp','http','https'
$filetypes = '.doc','.docx','.ppt','.pptx','.xls','.xlsx'

foreach ($p in $protocols) {
    Write-Output "Assigned $p to Chrome"
    \\10.51.87.13\Tech\SetUserFTA.exe $p ChromeHTML
}

foreach ($f in $filetypes) {
    Write-Output "Assignied $f to Chrome"
    \\10.51.87.13\Tech\SetUserFTA.exe $f ChromeHTML
}
}

# Logic

InstallOfficeExtension
AssociateFileTypes