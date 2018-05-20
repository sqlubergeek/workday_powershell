#Get Execution FileName and Path and use to create paths for additional processing files
$baseName = ([io.fileinfo]$MyInvocation.InvocationName.ToString()).FullName.Replace(([io.fileinfo]$MyInvocation.InvocationName.ToString()).Extension,"")
$ConfigFileName = $baseName + ".config"
$LogFileName = $baseName + ".log"


$ErrorActionPreference = "Stop"

function Write-Log
{
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true)]
        [ValidateNotNullOrEmpty()]
        [Alias("LogContent")]
        [string]$Message,

        [Parameter(Mandatory=$false)]
        [Alias('LogPath')]
        [string]$Path,
        
        [Parameter(Mandatory=$false)]
        [ValidateSet("Error","Warn","Info")]
        [string]$Level="Info"
        
    )

    Begin
    {
        # Set VerbosePreference to Continue so that verbose messages are displayed.
        $VerbosePreference = 'Continue'
    }
    Process
    {
		# If attempting to write to a log file in a folder/path that doesn't exist create the file including the path.
		if (!(Test-Path $Path)) {
			Write-Verbose "Creating $Path."
			$NewLogFile = New-Item $Path -Force -ItemType File
		}
		# Format Date for our Log File
		$FormattedDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

		# Write message to error, warning, or verbose pipeline and specify $LevelText
		switch ($Level) {
			'Error' {
				#Write Error but continue through the rest of the script to log the actual error
				Write-Error -ErrorAction Continue $Message
				$LevelText = 'ERROR:'
				}
			'Warn' {
				Write-Warning $Message
				$LevelText = 'WARNING:'
				}
			'Info' {
				Write-Verbose $Message
				$LevelText = 'INFO:'
				}
			}
		# Write log entry to $Path
		"$FormattedDate`t$LevelText`t $Message" | Out-File -FilePath $Path -Append
		if ($ErrorActionPreference -eq "Stop" -and $Level -eq "Error")
		{
			throw $Message
			exit
		}

	}
    End
    {
    }
}	

Function Put-WDExternal_Pay_Group{
    [CmdletBinding(DefaultParameterSetName="Help")]
    param(

        [Parameter(Mandatory = $true, ParameterSetName = 'InProgress')]
		[Parameter(Mandatory = $true, ParameterSetName = 'Complete')]
        [ValidateNotNullorEmpty()]
		[String] $PayGroupID,

		[Parameter(Mandatory = $true, ParameterSetName = 'InProgress')]
		[Parameter(Mandatory = $true, ParameterSetName = 'Complete')]
        [ValidateNotNullorEmpty()]
		[String] $PayPeriod,

        [Parameter(Mandatory = $true, ParameterSetName = 'InProgress')]
		[Parameter(Mandatory = $true, ParameterSetName = 'Complete')]
        [ValidateNotNullorEmpty()]
		[System.Management.Automation.PSCredential] $Credentials,

        [Parameter(Mandatory = $true, ParameterSetName = 'InProgress')]
		[Parameter(Mandatory = $true, ParameterSetName = 'Complete')]
		[String] $Uri,

        [Parameter(Mandatory = $true, ParameterSetName = 'InProgress')]
		[Parameter(Mandatory = $true, ParameterSetName = 'Complete')]
        [String] $LogFile,

		[Parameter(Mandatory = $true, ParameterSetName = 'InProgress')]
		[Switch] $InProgress,

		[Parameter(Mandatory = $true, ParameterSetName = 'Complete')]
		[Switch] $Complete


    ) #end param

    if( $InProgress ) { $Status = "b4a78dee73634adaa417b7976a17e921" }
    if( $Complete ) { $Status = "3e7f9597a57e429389af0e76e5e964e9" }

    if($Credentials)
    {
        $UserName = $Credentials.UserName
        $Password = $Credentials.GetNetworkCredential().password


    [xml]$xml='<soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" xmlns:bsvc="urn:com.workday/bsvc">
	<soapenv:Header>
		<wsse:Security soapenv:mustUnderstand="1" xmlns:wsse="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-secext-1.0.xsd" xmlns:wsu="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-utility-1.0.xsd">
			<wsse:UsernameToken>
				<wsse:Username/>
				<wsse:Password Type="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-username-token-profile-1.0#PasswordText"/>
			</wsse:UsernameToken>
		</wsse:Security>
	</soapenv:Header>
	<soapenv:Body>
		<bsvc:Put_External_Pay_Group_Request>
			<bsvc:External_Pay_Group_Data>
				<bsvc:External_Pay_Group_Reference>
					<bsvc:ID bsvc:type="Organization_Reference_ID"/>
				</bsvc:External_Pay_Group_Reference>
				<bsvc:Period_Status_Data>
					<bsvc:Period_Reference>
						<bsvc:ID bsvc:type="Period_ID"/>
					</bsvc:Period_Reference>
					<bsvc:Period_Status_Reference>
						<bsvc:ID bsvc:type="WID"/>
					</bsvc:Period_Status_Reference>
				</bsvc:Period_Status_Data>
			</bsvc:External_Pay_Group_Data>
		</bsvc:Put_External_Pay_Group_Request>
	</soapenv:Body>
</soapenv:Envelope>'

        $ns = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
        $ns.AddNamespace("bsvc", "urn:com.workday/bsvc")

        $xml.Envelope.Header.Security.UsernameToken.Username = $UserName
        $xml.Envelope.Header.Security.UsernameToken.Password.InnerText = $Password
        $xml.Envelope.Body.Put_External_Pay_Group_Request.External_Pay_Group_Data.External_Pay_Group_Reference.ID.InnerText = $PayGroupID
        $xml.Envelope.Body.Put_External_Pay_Group_Request.External_Pay_Group_Data.Period_Status_Data.Period_Reference.ID.InnerText = $PayPeriod
        $xml.Envelope.Body.Put_External_Pay_Group_Request.External_Pay_Group_Data.Period_Status_Data.Period_Status_Reference.ID.InnerText = $Status
    }
        try {
         	$post = Invoke-WebRequest -Uri $Uri -Method Post -Body $xml -ContentType "application/xml" -UseBasicParsing
        }
        catch {
            $result = $_.Exception.Response.GetResponseStream()
            Write-Log -Path $LogFile -Level Warn -Message "Error Updating Pay Period"
            $reader = New-Object System.IO.StreamReader($result)
            [xml]$responseBody = $reader.ReadToEnd();
            if($responseBody.Envelope.Body.Fault.faultstring)
            {
                Write-Log -Path $LogFile -Level Error -Message $responseBody.Envelope.Body.Fault.faultstring
            } else {
                Write-Log -Path $LogFile -Level Error -Message $_.Exception
            }
        }


}

Function Get-WDPayGroup {
    [CmdletBinding()]
    param(
		[Parameter(Mandatory = $true)]
		[String] $PayGroupID,

		[Parameter(Mandatory = $true)]
		[System.Management.Automation.PSCredential] $Credentials,

		[Parameter(Mandatory = $true)]
		[String] $Uri,

        [Parameter(Mandatory = $true)]
        [String] $LogFile
 
    ) #end param

    if($Credentials)
    {
        $UserName = $Credentials.UserName
        $Password = $Credentials.GetNetworkCredential().password

    [xml]$xml = '<soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" xmlns:bsvc="urn:com.workday/bsvc">
	<soapenv:Header>
		<wsse:Security soapenv:mustUnderstand="1" xmlns:wsse="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-secext-1.0.xsd" xmlns:wsu="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-utility-1.0.xsd">
			<wsse:UsernameToken>
				<wsse:Username/>
				<wsse:Password Type="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-username-token-profile-1.0#PasswordText"/>
			</wsse:UsernameToken>
		</wsse:Security>
	</soapenv:Header>
	<soapenv:Body>
		<bsvc:Get_External_Pay_Groups_Request>
			<bsvc:Request_References>
				<bsvc:External_Pay_Group_Reference>
					<bsvc:ID bsvc:type="Organization_Reference_ID"/>
				</bsvc:External_Pay_Group_Reference>
			</bsvc:Request_References>
		</bsvc:Get_External_Pay_Groups_Request>
	</soapenv:Body>
</soapenv:Envelope>'

        $ns = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
        $ns.AddNamespace("bsvc", "urn:com.workday/bsvc")

        $xml.Envelope.Header.Security.UsernameToken.Username = $UserName
        $xml.Envelope.Header.Security.UsernameToken.Password.InnerText = $Password
        $xml.Envelope.Body.Get_External_Pay_Groups_Request.Request_References.External_Pay_Group_Reference.ID.InnerText = $PayGroupID
        try {
         	$post = Invoke-WebRequest -Uri $Uri -Method Post -Body $xml -ContentType "application/xml" -UseBasicParsing
#            Write-Log -Path $LogFile -Level Info -Message "Retrieved Pay Group $PayGroupID"
        }
        catch {
            Write-Log -Path $LogFile -Level Warn -Message "Error getting Pay Group"
            if ($_.Exception.Response)
			{
				$result = $_.Exception.Response.GetResponseStream()
				$reader = New-Object System.IO.StreamReader($result)
				[xml]$responseBody = $reader.ReadToEnd();
			}
			if($responseBody.Envelope.Body.Fault.faultstring)
			{
				throw $responseBody.Envelope.Body.Fault.faultstring
			} else {
				throw $_.Exception
			}
        }
        [xml]$output = $post.content
        return $output

    }
}


#Check to see if password is encrypted.  Encrypt it if not
[xml]$ConfigFile = Get-Content $ConfigFileName
try {
    $SecurePassword = ConvertTo-SecureString $ConfigFile.Settings.Password -ErrorAction Stop
}
catch {
    $secpassword = $ConfigFile.Settings.Password | ConvertTo-SecureString -AsPlainText -Force | ConvertFrom-SecureString
    $ConfigFile.Settings.Password = $secpassword.ToString()
    $ConfigFile.save($ConfigFileName)
    [xml]$ConfigFile = Get-Content $ConfigFileName
}

$wd_username = $ConfigFile.Settings.UserName
$SecurePassword = ConvertTo-SecureString $ConfigFile.Settings.Password
$credentials = New-Object System.Management.Automation.PSCredential ($wd_username, $SecurePassword)

$MailRecipients = $ConfigFile.Settings.MailTo.Split(",")
$MailFrom = $ConfigFile.Settings.MailFrom
$PayGroups = $ConfigFile.Settings.PayGroups
$wd_URL = $ConfigFile.Settings.URL

$now = get-Date
#$now = [DateTime]"09/01/2017"

ForEach ($PayGroup in $PayGroups.ChildNodes)
{
    do{   
        try{
            $PayGroupName = $PayGroup.Name
			$Close = $false
			$Open = $false
            Write-Log -Path $LogFileName -Level Info -Message "Start Processing Pay Group $PayGroupName"
            $document = Get-WDPayGroup -LogFile $LogFileName -PayGroup $PayGroup.ID -Uri $wd_URL -Credentials $Credentials
            Write-Log -Path $LogFileName -Level Info -Message "Got Pay GrouP $PayGroupName"

            $namespace = @{env="http://schemas.xmlsoap.org/soap/envelope/"; wd="urn:com.workday/bsvc"}
    #        $previous_periodXML = select-xml -Xml $document -XPath "//wd:Last_Completed_Period_Reference/wd:ID[@wd:type='Period_ID']" -Namespace $namespace
    #        $previous_period = $previous_periodXML.ToString()
            $current_periodXML = select-xml -Xml $document -XPath "//wd:Current_Period_Reference/wd:ID[@wd:type='Period_ID']" -Namespace $namespace
            Write-Log -Path $LogFileName -Level Info -Message "Got Curret Period XML: $($current_periodXML.OuterXML)"
            $next_periodXML = select-xml -Xml $document -XPath "//wd:Next_Period_Reference/wd:ID[@wd:type='Period_ID']" -Namespace $namespace
            Write-Log -Path $LogFileName -Level Info -Message "Got Next Period XML: $($next_periodXML.OuterXML)"
            if($current_periodXML){
                $current_period = $current_periodXML.ToString()
                [DateTime]$current_period_end_date = (select-xml -Xml $document -XPath "//wd:Period_Content_Data[wd:Period_Reference/wd:ID[@wd:type='Period_ID']='$current_period']/wd:End_Date" -Namespace $namespace).ToString().Substring(0,10)
                if($PayGroup.CloseDay -eq -1) {
                    $CloseDate = $current_period_end_date.addDays(1)
                } else { 
                    $CloseDateYear = $current_period_end_date.Year
                    $CloseDateMonth = $current_period_end_date.Month
                    $CloseDay = $PayGroup.CloseDay
                    $CloseDate =[DateTime]"$CloseDateYear/$CloseDateMonth/$CloseDay"
                }
                if($now -ge $CloseDate) { $Close = $true } else { $Close = $false }
            } else {
                $Close = $false
            }

            if($next_periodXML){
                $next_period = $next_periodXML.ToString()
                [DateTime]$next_period_start_date = (select-xml -Xml $document -XPath "//wd:Period_Content_Data[wd:Period_Reference/wd:ID[@wd:type='Period_ID']='$next_period']/wd:Start_Date" -Namespace $namespace).ToString().Substring(0,10)
                if($PayGroup.OpenDay -eq -1){
                    $OpenDate = $next_period_start_date
                } else {
                    $OpenDateYear = $next_period_start_date.Year
                    $OpenDateMonth = $next_period_start_date.Month
                    $OpenDay = $PayGroup.OpenDay
                    $OpenDate =[DateTime] "$OpenDateYear/$OpenDateMonth/$OpenDay"
                }
                if($now -ge $OpenDate) { $Open = $true } else { $Open = $false }
            } else {
                $Open = $false
            }

            if ($Close -eq $true) {
                Write-Log -Path $LogFileName -Level Info -Message "Closing $PayGroupName Period Ending on $current_period_end_date."
                Put-WDExternal_Pay_Group -LogFile $LogFileName -PayGroup $PayGroup.ID -PayPeriod $current_period -Credentials $Credentials -Uri $wd_URL -Complete
            }
            if ($Open -eq $true) {
                Write-Log -Path $LogFileName -Level Info -Message "Opening $PayGroupName Period Starting on $next_period_start_date." 
                Put-WDExternal_Pay_Group -LogFile $LogFileName -PayGroup $PayGroup.ID -PayPeriod $next_period -Credentials $Credentials -Uri $wd_URL -InProgress
            }
            Write-Log -Path $LogFileName -Level Info -Message "Finished Processing Pay Group $PayGroupName"

        } catch {
            Write-Log -Path $LogFileName -Level Warn -Message "Error Processing Pay Group $PayGroupName"
            Write-Log -Path $LogFileName -Level Error -Message $_.Exception
        }
    } until($Close -eq $false -and $Open -eq $false)
}