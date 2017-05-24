Function Get-WCMSiteCode {
	<#
	.SYNOPSIS
	Automatic detection for the site code
	.DESCRIPTION
	Gets the site code of the computer provided in -ComputerName
	.EXAMPLE
	Get-WCMSiteCode -ComputerName SCCM01
	#>
	[CmdLetBinding()]

	Param(
		[Parameter(Mandatory=$True,HelpMessage='Provide the computername')]
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
	<#
	.SYNOPSIS
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
        [Parameter(Mandatory=$True,HelpMessage='Provide the name of the variable to create',Position=1,ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True)]
        [String] $Name,

        [Parameter(Mandatory=$True,HelpMessage='Provide the value of the variable to create',Position=2,ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True)]
        [String] $Value,
        
		[Parameter(HelpMessage='Provide the (remote) computername of the site server where de collection variable will be created')]
        [ValidateScript({ Test-Connection -ComputerName $_ -Count 1 -Quiet })]
        [String] $ComputerName = ".",

		[Parameter(HelpMessage='Provide the site code of the site (is blank, auto detection will be attempted)')]
        [ValidateLength(3,3)]
        [String] $Site,

		[Parameter(HelpMessage='Provide the collection identifier of the targeted collection')]
        [Parameter(Mandatory=$True)]
        [String] $CollectionID,
        
		[Parameter(HelpMessage='Set to true if the value should be hidden when viewed in the console')]
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