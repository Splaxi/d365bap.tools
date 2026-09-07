
<#
    .SYNOPSIS
        Get UDE environment models.
        
    .DESCRIPTION
        Gets the models for a specified environment.
        
        Works against any unified environment (UDE, USE and others).
        Queries the msprov_fnomodule table through the Dataverse Web API.
        
    .PARAMETER EnvironmentId
        The id of the environment that you want to work against
        
    .PARAMETER Name
        The name of the model that you are looking for.
        
        Supports wildcard patterns.
        
    .PARAMETER LatestOnly
        Instructs the cmdlet to return only the latest model.
        
        Is based on the modified date.
        
    .PARAMETER AsExcelOutput
        Instructs the function to export the results to an Excel file.
        
    .EXAMPLE
        PS C:\> Get-UdeEnvironmentModel -EnvironmentId "env-123"
        
        This will retrieve all models for the specified environment id.
        
    .EXAMPLE
        PS C:\> Get-UdeEnvironmentModel -EnvironmentId "env-123" -Name "B"
        
        This will retrieve the model named B for the specified environment id.
        
    .EXAMPLE
        PS C:\> Get-UdeEnvironmentModel -EnvironmentId "env-123" -Name "B*" -LatestOnly
        
        This will retrieve only the latest model matching B* for the specified environment id.
        It is based on the modified date.
        
    .EXAMPLE
        PS C:\> Get-UdeEnvironmentModel -EnvironmentId "env-123" -AsExcelOutput
        
        This will retrieve all models for the specified environment id.
        Will output all details into an Excel file, that will auto open on your machine.
        
    .NOTES
        Author: Mötz Jensen (@Splaxi)
#>
function Get-UdeEnvironmentModel {
    [CmdletBinding()]
    [OutputType('System.Object[]')]
    param (

        [Parameter (Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [Alias("PpacEnvId")]
        [string] $EnvironmentId,

        [string] $Name = "*",

        [switch] $LatestOnly,

        [switch] $AsExcelOutput
    )

    begin {
    }

    process {
        if (Test-PSFFunctionInterrupt) { return }

        $envObj = Get-UnifiedEnvironment -EnvironmentId $EnvironmentId -SkipVersionDetails | Select-Object -First 1

        if ($null -eq $envObj) {
            $messageString = "Could not find environment with Id <c='em'>$EnvironmentId</c>. Please verify the Id and try again, or list available environments using <c='em'>Get-UnifiedEnvironment</c>. Consider using wildcards if needed."

            Write-PSFMessage -Level Important -Message $messageString
            Stop-PSFFunction -Message "Stopping because environment was NOT found based on the id." `
                -Exception $([System.Exception]::new($($messageString -replace '<[^>]+>', '')))
            return
        }

        $baseUri = $envObj.PpacEnvUri + "/" #! Very important to have the trailing slash

        $secureToken = (Get-AzAccessToken -ResourceUrl $baseUri -AsSecureString).Token
        $tokenWebApiValue = ConvertFrom-SecureString -AsPlainText -SecureString $secureToken

        $headers = @{
            "Authorization"    = "Bearer $($tokenWebApiValue)"
            "Accept"           = "application/json;odata.metadata=minimal" # minimal || full
            "OData-MaxVersion" = "4.0"
            "OData-Version"    = "4.0"
            "Prefer"           = "odata.include-annotations=*"
        }

        $localUri = $baseUri + "api/data/v9.0/msprov_fnomodules"

        $colModules = Invoke-RestMethod -Uri $localUri `
            -Method Get `
            -Headers $headers | `
            Select-Object -ExpandProperty value | `
            Where-Object { $_.msprov_name -like $Name } | `
            Sort-Object -Property modifiedon -Descending

        if ($LatestOnly) {
            $colModules = $colModules | Select-Object -First 1
        }

        $resCol = @(
            $colModules | Select-PSFObject -TypeName 'D365Bap.Tools.UdeEnvironmentModel' `
                -Property "msprov_name As Name",
            "msprov_fnomoduleid As ModuleId",
            "'_createdby_value@OData.Community.Display.V1.FormattedValue' As CreatedBy",
            "'_modifiedby_value@OData.Community.Display.V1.FormattedValue' As ModifiedBy",
            "'statuscode@OData.Community.Display.V1.FormattedValue' As Status",
            "'statecode@OData.Community.Display.V1.FormattedValue' As State",
            "createdon As Created",
            @{ Name = "CreatedUtc"; Expression = { ([datetime]$_.createdon).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss") } },
            "modifiedon As Modified",
            @{ Name = "ModifiedUtc"; Expression = { ([datetime]$_.modifiedon).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss") } },
            "versionnumber As Version",
            * `
                -ExcludeProperty '@odata.etag'
        )

        if ($AsExcelOutput) {
            $resCol | Export-Excel -WorksheetName "Get-UdeEnvironmentModel"
            return
        }

        $resCol
    }

    end {
    }
}