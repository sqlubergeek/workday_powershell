#Get Execution FileName and Path and use to create paths for additional processing files
$baseName = "$((Get-Item $PSCommandPath ).DirectoryName)\$((Get-Item $PSCommandPath ).Basename)"
$ConfigFileName = $baseName + ".config"
$employeeFile = $baseName + ".txt"
#load and read configuration settings from config file
[xml]$ConfigFile = Get-Content $ConfigFileName

#Check to see if password is encrypted.  Encrypt it if not
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
$ErrorFrom = $ConfigFile.Settings.MailFrom
$uri = $ConfigFile.Settings.URL
$newPassword = $ConfigFile.Settings.NewPassword
$SecurePassword = ConvertTo-SecureString $ConfigFile.Settings.Password

$credentials = New-Object System.Management.Automation.PSCredential ($wd_username, $SecurePassword)

$Message = ''

Import-Csv $employeeFile -OutVariable dt | Out-Null

$wd_username = $credentials.UserName
$wd_password = $credentials.GetNetworkCredential().password

$counter = 0

foreach ($d in $dt)
{
    [xml]$xml = '<soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" xmlns:bsvc="urn:com.workday/bsvc">
	<soapenv:Header>
		<wsse:Security soapenv:mustUnderstand="1" xmlns:wsse="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-secext-1.0.xsd" xmlns:wsu="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-utility-1.0.xsd">
			<wsse:UsernameToken>
				<wsse:Username></wsse:Username>
				<wsse:Password Type="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-username-token-profile-1.0#PasswordText"></wsse:Password>
			</wsse:UsernameToken>
		</wsse:Security>
	</soapenv:Header>
	<soapenv:Body>
		<bsvc:Workday_Account_for_Worker_Update>
			<bsvc:Worker_Reference>
				<bsvc:Employee_Reference>
					<bsvc:Integration_ID_Reference>
						<bsvc:ID bsvc:System_ID="WD-EMPLID"></bsvc:ID>
					</bsvc:Integration_ID_Reference>
				</bsvc:Employee_Reference>
				<bsvc:Contingent_Worker_Reference>
					<bsvc:Integration_ID_Reference>
						<bsvc:ID bsvc:System_ID="WD-EMPLID"></bsvc:ID>
					</bsvc:Integration_ID_Reference>
				</bsvc:Contingent_Worker_Reference>
			</bsvc:Worker_Reference>
			<bsvc:Workday_Account_for_Worker_Data>
				<bsvc:User_Name></bsvc:User_Name>
				<bsvc:Password>' + $newPassword + '</bsvc:Password>
				<bsvc:Require_New_Password_at_Next_Sign_In>1</bsvc:Require_New_Password_at_Next_Sign_In>
			</bsvc:Workday_Account_for_Worker_Data>
		</bsvc:Workday_Account_for_Worker_Update>
	</soapenv:Body>
</soapenv:Envelope>'

	$ns = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
	$ns.AddNamespace("bsvc", "urn:com.workday/bsvc")

	$xml.Envelope.Header.Security.UsernameToken.Username = $wd_username
	$xml.Envelope.Header.Security.UsernameToken.Password.InnerText = $wd_password

	if ($d.IsEmployee -eq "true") {
		$node = $xml.SelectSingleNode("//bsvc:Contingent_Worker_Reference", $ns)
		$removed = $xml.Envelope.Body.Workday_Account_for_Worker_Update.Worker_Reference.RemoveChild($node)
		$xml.Envelope.Body.Workday_Account_for_Worker_Update.Worker_Reference.Employee_Reference.Integration_ID_Reference.ID.InnerText = $d.empid
	} else {
		$node = $xml.SelectSingleNode("//bsvc:Employee_Reference", $ns)
		$removed = $xml.Envelope.Body.Workday_Account_for_Worker_Update.Worker_Reference.RemoveChild($node)
		$xml.Envelope.Body.Workday_Account_for_Worker_Update.Worker_Reference.Contingent_Worker_Reference.Integration_ID_Reference.ID.InnerText = $d.empid
	}
	$xml.Envelope.Body.Workday_Account_for_Worker_Update.Workday_Account_for_Worker_Data.User_Name = $d.alias
	try {
		$counter = $counter + 1
		$post = Invoke-WebRequest -Uri $uri -Method Post -Body $xml -ContentType "application/xml" -UseBasicParsing 
		$Message = $Message + "Changed password for " + $d.empid + "`r`n"
	}
	catch {
		$Message = $Message + "Couldn't Change password for " + $d.empid + "`r`n"
        if($_.Exception.Response) {
		    $result = $_.Exception.Response.GetResponseStream()
		    $reader = New-Object System.IO.StreamReader($result)
		    [xml]$responseBody = $reader.ReadToEnd();
            $errorMsg = "Error: " + $responseBody.Envelope.Body.Fault.faultstring + "`r`n"
		    $Message = $Message +  $errorMsg
        }
		$Message = $Message + $_.Exception + "`r`n"
		if ( $counter -le 1 ) {
			$Message = $Message + "Error occurred on first password change.`r`nAborting Process`r`n"
			break
		}
	}
}
