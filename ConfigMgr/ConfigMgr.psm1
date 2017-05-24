<Function Get-WCMSiteCode {
	[CmdLetBinding()]

	Param(
		[Parameter(Mandatory=$True)]
		[ValidateScript({Test-Connection $_ -Count 1 -Quiet})]
		[String] $ComputerName
	)

	Try {
        $locationProvider = Get-CimInstance -ClassName "SMS_ProviderLocation" -Namespace "Root\SMS" -ComputerName $ComputerName -ErrorAction SilentlyContinue
            
        If( ($locationProvider -eq $null) -or ($locationProvider.SiteCode -eq "")) {
            Throw [System.Management.Automation.ItemNotFoundException] "No valid sitecode found on $($ComputerName)"
        }
        Else {
			Return $locationProvider.SiteCode
        }
    } Catch {
        Write-Error -Exception $_ -Category InvalidResult
    }

	Return $null
}

Function New-WCMCollectionVariable {
	<#.Synopsis
	Add a collection variable
	.DESCRIPTION
	Add a collection variable to an ConfigMgr collection
	.EXAMPLE
	New-WCMCollectionVariable -Name MyCollectionVariable - Value "Example value"
	.EXAMPLE
	New-WCMCollectionVariable -Name MyCollectionVariable - Value "Example hidden value" -HideValueInConsole $True
	#>
    [CmdLetBinding()]
    
    Param(
        [Parameter(Mandatory=$True,Position=1,ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True)]
        [String] $Name,

        [Parameter(Mandatory=$True,Position=2,ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True)]
        [String] $Value,
        
        [ValidateScript({ Test-Connection -ComputerName $_ -Count 1 -Quiet })]
        [String] $ComputerName = ".",

        [ValidateLength(3,3)]
        [String] $Site,

        [Parameter(Mandatory=$True)]
        [String] $CollectionID,
        
        [Parameter(Position=3,ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True)]
        [String] $HideValueInConsole = $False
    )

    Begin {
        If( $ComputerName -ne "." ) {
            Write-Verbose "Using remote computer '$ComputerName'"
        }

        If( $Site -eq "" ) {
            Write-Verbose "No site code was provided, attempt auto detection"
            $Site = Get-WCMSiteCode -ComputerName $ComputerName -ErrorAction Stop
        }
    }

    Process {
        $objCollectionVariable = ([wmiclass] "\\$ComputerName\root\sms\site_$($Site):SMS_CollectionVariable").CreateInstance()
        $objCollectionVariable.Name = $Name
        $objCollectionVariable.Value = $Value
        $objCollectionVariable.IsMasked = [System.Convert]::ToBoolean($HideValueInConsole)
    
        $objCollectionSetting = Get-WmiObject -ComputerName $ComputerName -Class SMS_CollectionSettings -Namespace root\sms\site_$($Site) -Filter "CollectionID = '$collectionID'"

        If( $objCollectionSetting -eq $null ) {
            Write-Verbose "Collection setting object not found for this collection, creating new instance"
            $objCollectionSetting = ([wmiclass] "\\$ComputerName\root\sms\site_$($Site):SMS_CollectionSettings").CreateInstance()
        }
        Else {
            $objCollectionSetting.Get()
        }

        $objCollectionSetting.CollectionVariables += $objCollectionVariable
        #$objCollectionSetting.Put() | Out-Null
        #Write-Output "Sucessfully created collection variable $Name with value $Value"
    }
}