<Function Get-WCMSiteCode {
	[CmdLetBinding()]

	Param(
		[Parameter(Mandatory=$True)]
		[ValidateScript({Test-Connection $_ -Count 1 -Quiet})]
		[String] $ComputerName
	)

}