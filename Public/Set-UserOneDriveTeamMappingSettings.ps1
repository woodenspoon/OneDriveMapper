<#
.SYNOPSIS

Sets the OneDriveMapper Function App name and auth key so that Mount-UserOneDriveTeamMappings can be run more transparently

.DESCRIPTION

This function will set the registry keys OneDriveMapperFunctionApp and OneDriveMapperAuthKey in hive HKLM:\SOFTWARE\WST\OneDriveMapper, so that Mount-UserOneDriveTeamMappings can read them during execution.

#>
Function Set-UserOneDriveTeamMappingSettings {
    [CmdletBinding()]
    Param(
        [ValidatePattern('^[a-zA-Z0-9]+$')][Parameter(Mandatory=$true)][String]$OneDriveMapperFunctionApp,
        [ValidatePattern('^[A-Za-z0-9+/=]|=[^=]|={3,}$')][Parameter(Mandatory=$true)][String]$OneDriveMapperAuthKey
    )

    New-Item -Path 'HKLM:\SOFTWARE\WST' -ErrorAction SilentlyContinue | Out-Null
    New-Item -Path 'HKLM:\SOFTWARE\WST\OneDriveMapper' -ErrorAction SilentlyContinue | Out-Null
    Set-ItemProperty -Path 'HKLM:\SOFTWARE\WST\OneDriveMapper' -Name 'OneDriveMapperFunctionApp' -Value $OneDriveMapperFunctionApp | Out-Null
    Set-ItemProperty -Path 'HKLM:\SOFTWARE\WST\OneDriveMapper' -Name 'OneDriveMapperAuthKey' -Value $OneDriveMapperAuthKey | Out-Null
}
