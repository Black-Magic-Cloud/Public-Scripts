#========================================================================
# Created on:		a long time ago....
# Created by:		Andreas Hähnel
# Organization:		Black Magic Cloud
# Script Version: 	1.0
#========================================================================
# RequiredPermissions and modules:
# Delegated (work or school account)		Not supported.
# Delegated (personal Microsoft account)	Not supported.
# Application								Sites.FullControl.All
#                                           User.Read.All
#========================================================================
# Description:
# this script identifies orphaned users in SPO and removes them.
#========================================================================

#Requires -Version 3
#Requires -Modules AzureAD
#Requires -Modules PnP.PowerShell

param(
        [Parameter(Mandatory)][string]$thumbprint,
        [Parameter(Mandatory)][string]$appID,
        [Parameter(Mandatory)][string]$tenantID,
        [Parameter(Mandatory)][string]$tenantURL,
        [Parameter(Mandatory)][string]$tenant
    )

# connect to aad
Connect-AzureAD -CertificateThumbprint $thumbprint -ApplicationId $appId -TenantId $tenantId

# get all aad user entities
$allAADUsers = Get-AzureADUser -All:$true
$allAADUserUPNs = @()
$allAADUsers | ForEach-Object {$allAADUserUPNs += $_.UserPrincipalName}

# connect to spo admin site to get all sites in tenant
$pnpAdminConnection = Connect-PnPOnline -Url $tenantURL -ClientId $appId -Thumbprint $thumbprint -Tenant $tenant -ReturnConnection
$allTenantSites = Get-PnPTenantSite -IncludeOneDriveSites -Connection $pnpAdminConnection
# iterate through all sites and do magic
$allTenantSites | ForEach-Object {
    $pnpSiteConnection = Connect-PnPOnline -Url $_.Url -ClientId $appId -Thumbprint $thumbprint -Tenant $tenant -ReturnConnection
    Write-Host "connected to $($_.Url)"
    $allSiteEntities = Get-PnPUser -Connection $pnpSiteConnection
    # filter for user entities 
    $allSiteUsers = $allSiteEntities | Where-Object {$_.LoginName -like "i:0#.f|*"}
    $allSiteUsers | ForEach-Object {
        # filter anonymous links
        if ($_.LoginName -notlike "*urn%3aspo%3aanon#*") {
            if ($_.Email -notin $allAADUserUPNs) {
                # remove if not present in aad
                Remove-PnPUser -Identity $_.Id -Force -Connection $pnpSiteConnection
                Write-Host "Removed orphaned user $($_.LoginName)"
            }
        }
    }
}