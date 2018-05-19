Function Set-WDEmail {
    <#
       .Synopsis
        Uploads Email to Public Work Email and Public Lync Address
       .Description
        Uses the Maintain_Contact_Information_for_Person_Event API call to load Email and Lync
        Addresses into Workday for an Employee or Contingent Worker.
       .Example
        Put-WorkdayEmailLync -Employee 4101 -Email dsmith@manh.com -Contractor FALSE
       .Parameter Employee
        Employee or Contingent Worker ID
       .Parameter Email
        Email Address
       .Parameter Contractor
        Is the worker a Contractor
     #>
    [CmdletBinding()]
    param(
          [Parameter(Mandatory = $true)]
          [alias("ID","EmpID","EmployeeID","ContractorID")]
          [String] $Employee,

          [Parameter(Mandatory = $true)]
          [alias("Mail")]
          [String] $Email,

          [Parameter(Mandatory = $false)]
          [alias("ContingentWorker","Coop","Co-Op","Intern")]
          [Switch] $Contractor,

          [Parameter(Mandatory = $false)]
          [String] $URI,

          [Parameter(Mandatory = $false)]
          [Switch] $ReturnXML,

          [Parameter(Mandatory = $true)]
          [System.Management.Automation.PSCredential] $Credentials,

          [Parameter(Mandatory = $false)]
          [DateTime] $HireDate

    ) #end param

    $UserName = $Credentials.UserName
    $Password = $Credentials.GetNetworkCredential().password
        $URI = "https://wd5-services1.myworkday.com/ccx/service/manh/Human_Resources/v23.2"

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
            <bsvc:Maintain_Contact_Information_for_Person_Event_Request xmlns:bsvc="urn:com.workday/bsvc">
                <bsvc:Business_Process_Parameters>
                    <bsvc:Auto_Complete>1</bsvc:Auto_Complete>
                    <bsvc:Run_Now>1</bsvc:Run_Now>
                    <bsvc:Comment_Data>
                        <bsvc:Comment>AD to Workday Integration</bsvc:Comment>
                    </bsvc:Comment_Data>
                </bsvc:Business_Process_Parameters>
                <bsvc:Maintain_Contact_Information_Data>
                    <bsvc:Worker_Reference>
                        <bsvc:ID bsvc:type="Employee_ID"></bsvc:ID>
                    </bsvc:Worker_Reference>
                    <bsvc:Effective_Date></bsvc:Effective_Date>
                    <bsvc:Worker_Contact_Information_Data>
                        <bsvc:Email_Address_Data>
                            <bsvc:Email_Address></bsvc:Email_Address>
                            <bsvc:Usage_Data bsvc:Public="1">
                                <bsvc:Type_Data bsvc:Primary="1">
                                    <bsvc:Type_Reference>
                                        <bsvc:ID bsvc:type="Communication_Usage_Type_ID">WORK</bsvc:ID>
                                    </bsvc:Type_Reference>
                                </bsvc:Type_Data>
                            </bsvc:Usage_Data>
                        </bsvc:Email_Address_Data>
                        <bsvc:Instant_Messenger_Data>
                            <bsvc:Instant_Messenger_Address></bsvc:Instant_Messenger_Address>
                            <bsvc:Instant_Messenger_Type_Reference>
                                <bsvc:ID bsvc:type="Instant_Messenger_Type_ID">Lync</bsvc:ID>
                            </bsvc:Instant_Messenger_Type_Reference>
                            <bsvc:Usage_Data bsvc:Public="1">
                                <bsvc:Type_Data bsvc:Primary="1">
                                    <bsvc:Type_Reference>
                                        <bsvc:ID bsvc:type="Communication_Usage_Type_ID">WORK</bsvc:ID>
                                    </bsvc:Type_Reference>
                                </bsvc:Type_Data>
                            </bsvc:Usage_Data>
                        </bsvc:Instant_Messenger_Data>
                    </bsvc:Worker_Contact_Information_Data>
                </bsvc:Maintain_Contact_Information_Data>
            </bsvc:Maintain_Contact_Information_for_Person_Event_Request>
        </soapenv:Body>
    </soapenv:Envelope>'

    $ns = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
    $ns.AddNamespace("bsvc", "urn:com.workday/bsvc")

    $now = Get-Date
    $EffectiveDate = $now

    IF ( $HireDate )
    {
        IF( $HireDate -gt $EffectiveDate )
        {
            $EffectiveDate = $HireDate
        } 
    }
    
    $xml.Envelope.Header.Security.UsernameToken.Username = $UserName
    $xml.Envelope.Header.Security.UsernameToken.Password.InnerText = $Password
    $xml.Envelope.Body.Maintain_Contact_Information_for_Person_Event_Request.Maintain_Contact_Information_Data.Worker_Reference.ID.InnerText = $Employee
    $xml.Envelope.Body.Maintain_Contact_Information_for_Person_Event_Request.Maintain_Contact_Information_Data.Worker_Contact_Information_Data.Instant_Messenger_Data.Instant_Messenger_Address = $Email
    $xml.Envelope.Body.Maintain_Contact_Information_for_Person_Event_Request.Maintain_Contact_Information_Data.Worker_Contact_Information_Data.Email_Address_Data.Email_Address = $Email

    $xml.Envelope.Body.Maintain_Contact_Information_for_Person_Event_Request.Maintain_Contact_Information_Data.Effective_Date = (Get-Date -Date $EffectiveDate -Format s).ToString()

    if ($Contractor) {
        $xml.Envelope.Body.Maintain_Contact_Information_for_Person_Event_Request.Maintain_Contact_Information_Data.Worker_Reference.ID.SetAttribute("bsvc:type","Contingent_Worker_ID")
    } else {
        $xml.Envelope.Body.Maintain_Contact_Information_for_Person_Event_Request.Maintain_Contact_Information_Data.Worker_Reference.ID.SetAttribute("bsvc:type","Employee_ID")
    }

    try {
        $post = Invoke-WebRequest -Uri $URI -Method Post -Body $xml -ContentType "application/xml"
        Write-Host "Changed email for $Employee"
    }
    catch {
        $result = $_.Exception.Response.GetResponseStream()
        $reader = New-Object System.IO.StreamReader($result)
        [xml]$responseBody = $reader.ReadToEnd();
        Write-Host "Error: $($responseBody.Envelope.Body.Fault.faultstring)"
    }

    if ($ReturnXML)
    {
        $xml.OuterXml
    }
}

Function Update-URIFormat {
    <#
       .Synopsis
        Modify Report URL to make it a specific report type
       .Parameter URI
        URI for Report
       .Parameter Format
        Format Type for the report
     #>
    [CmdletBinding()]
    param(
          [Parameter(Mandatory = $true)]
          [alias("URI","Report")]
          [String] $URL,

          [Parameter(Mandatory = $true)]
          [ValidatePattern("xml|csv|simplexml|json|gdata")]
          [String] $Format
    )

    Add-Type -AssemblyName "System.Web"

    [URI]$URI = $URL
    IF ($URI.Query -eq "" -and $Format -ne "xml")
    {
        $URI = $URI.OriginalString + "?format=$Format"
    }
    ELSEIF ($URI.Query -match "format=")
    {
        $query = [System.Web.HttpUtility]::ParseQueryString($URI.Query)
        IF ($query.GetValues("format") -ne $Format)
        {
            $URI = $URI.OriginalString -replace "format=$($query.GetValues("format"))", "format=$Format"
        }
    }
    ELSEIF ($format -ne "xml")
    {
        $URI = $URI.OriginalString + "&format=$Format"
    }
    RETURN $URI.OriginalString
}

Function Get-WDReportJSON {
    <#
       .Synopsis
        Download a Workday Report via XML
       .Parameter URI
        URI for Report
       .Parameter Credentials
        Credential object for login
       .Parameter Username
        Username to log in with
       .Parameter Password
        Password to log in with
     #>
    [CmdletBinding()]
    param(
          [Parameter(Mandatory = $true)]
          [alias("URI","Report")]
          [String] $URL,

          [Parameter(Mandatory = $false)]
          [System.Management.Automation.PSCredential] $Credentials,

          [Parameter(Mandatory = $false)]
          [String] $Username,

          [Parameter(Mandatory = $false)]
          [String] $Password

    ) #end param

    if(!$Credentials)
    {
        if ($Username -and $Password)
        {
            $secpasswd = ConvertTo-SecureString $Password -AsPlainText -Force
            $Credentials = New-Object System.Management.Automation.PSCredential ($Username, $secpasswd)        }
        else
        {
            $Credentials = Get-Credential
        }
    }

    $URI = Update-URIFormat -URL $URL -Format "json"

    if($Credentials)
    {
        try {
            Write-Host $URI
            $results = Invoke-WebRequest -Uri $URI -Credential $Credentials
            $json_content = $results.Content
            Return $json_content
        }
        catch {
            $result = $_.Exception.Response.GetResponseStream()
            $reader = New-Object System.IO.StreamReader($result)
            [xml]$responseBody = $reader.ReadToEnd();
            Write-Host "Error: $($responseBody.Envelope.Body.Fault.faultstring)"
#            $_.Exception
#            Write-Host "Couldn't retrieve report"
        }
    }
    Return $null
}

Function Get-WDReportGDATA {
    <#
       .Synopsis
        Download a Workday Report via XML
       .Parameter URI
        URI for Report
       .Parameter Credentials
        Credential object for login
       .Parameter Username
        Username to log in with
       .Parameter Password
        Password to log in with
     #>
    [CmdletBinding()]
    param(
          [Parameter(Mandatory = $true)]
          [alias("URI","Report")]
          [String] $URL,

          [Parameter(Mandatory = $false)]
          [System.Management.Automation.PSCredential] $Credentials,

          [Parameter(Mandatory = $false)]
          [String] $Username,

          [Parameter(Mandatory = $false)]
          [String] $Password

    ) #end param

    if(!$Credentials)
    {
        if ($Username -and $Password)
        {
            $secpasswd = ConvertTo-SecureString $Password -AsPlainText -Force
            $Credentials = New-Object System.Management.Automation.PSCredential ($Username, $secpasswd)        }
        else
        {
            $Credentials = Get-Credential
        }
    }

    $URI = Update-URIFormat -URL $URL -Format "gdata"

    if($Credentials)
    {
        try {
            Write-Host $URI
            $results = Invoke-WebRequest -Uri $URI -Credential $Credentials
            $gdata_content = $results.Content
            Return $gdata_content
        }
        catch {
            $_.Exception
            Write-Host "Couldn't retrieve report"
        }
    }
    Return $null
}

Function Get-WDReportXML {
    <#
       .Synopsis
        Download a Workday Report via XML
       .Parameter URI
        URI for Report
       .Parameter Credentials
        Credential object for login
       .Parameter Username
        Username to log in with
       .Parameter Password
        Password to log in with
     #>
    [CmdletBinding()]
    param(
          [Parameter(Mandatory = $true)]
          [alias("URI","Report")]
          [String] $URL,

          [Parameter(Mandatory = $false)]
          [System.Management.Automation.PSCredential] $Credentials,

          [Parameter(Mandatory = $false)]
          [String] $Username,

          [Parameter(Mandatory = $false)]
          [String] $Password

    ) #end param

    if(!$Credentials)
    {
        if ($Username -and $Password)
        {
            $secpasswd = ConvertTo-SecureString $Password -AsPlainText -Force
            $Credentials = New-Object System.Management.Automation.PSCredential ($Username, $secpasswd)        }
        else
        {
            $Credentials = Get-Credential
        }
    }

    $URI = Update-URIFormat -URL $URL -Format "xml"

    if($Credentials)
    {
        try {
            $results = Invoke-WebRequest -Uri $URI -Credential $Credentials
            [xml] $xml_content = $results.Content
            Return $xml_content
        }
        catch {
            $_.Exception
            Write-Host "Couldn't retrieve report"
        }
    }
    Return $null
}

Function Get-WDReportSimpleXML {
    <#
       .Synopsis
        Download a Workday Report via XML
       .Parameter URI
        URI for Report
       .Parameter Credentials
        Credential object for login
       .Parameter Username
        Username to log in with
       .Parameter Password
        Password to log in with
     #>
    [CmdletBinding()]
    param(
          [Parameter(Mandatory = $true)]
          [alias("URI","Report")]
          [String] $URL,

          [Parameter(Mandatory = $false)]
          [System.Management.Automation.PSCredential] $Credentials,

          [Parameter(Mandatory = $false)]
          [String] $Username,

          [Parameter(Mandatory = $false)]
          [String] $Password

    ) #end param

    if(!$Credentials)
    {
        if ($Username -and $Password)
        {
            $secpasswd = ConvertTo-SecureString $Password -AsPlainText -Force
            $Credentials = New-Object System.Management.Automation.PSCredential ($Username, $secpasswd)        }
        else
        {
            $Credentials = Get-Credential
        }
    }

    $URI = Update-URIFormat -URL $URL -Format "simplexml"

    if($Credentials)
    {
        try {
            $results = Invoke-WebRequest -Uri $URI -Credential $Credentials
            [xml] $xml_content = $results.Content
            Return $xml_content
        }
        catch {
            $_.Exception
            Write-Host "Couldn't retrieve report"
        }
    }
    Return $null
}

Function Get-WDReportCSV {
    <#
       .Synopsis
        Download a Workday Report via CSV
       .Parameter URI
        URI for Report
       .Parameter Credentials
        Credential object for login
       .Parameter Username
        Username to log in with
       .Parameter Password
        Password to log in with
     #>
    [CmdletBinding()]
    param(
          [Parameter(Mandatory = $true)]
          [alias("URI","Report")]
          [String] $URL,

          [Parameter(Mandatory = $false)]
          [System.Management.Automation.PSCredential] $Credentials,

          [Parameter(Mandatory = $false)]
          [String] $Username,

          [Parameter(Mandatory = $false)]
          [String] $Password

    ) #end param

    if(!$Credentials)
    {
        if ($Username -and $Password)
        {
            $secpasswd = ConvertTo-SecureString $Password -AsPlainText -Force
            $Credentials = New-Object System.Management.Automation.PSCredential ($Username, $secpasswd)        }
        else
        {
            $Credentials = Get-Credential
        }
    }

    $URI = Update-URIFormat -URL $URL -Format "csv"

    if($Credentials)
    {
        try {
	    	$results = ""


            [net.httpwebrequest]$httpwebrequest = [net.webrequest]::create($URI)
            $httpwebrequest.Credentials = $credentials
            [net.httpWebResponse]$httpwebresponse = $httpwebrequest.getResponse()
            $reader = new-object IO.StreamReader($httpwebresponse.getResponseStream())
            $results = $reader.ReadToEnd()
            $reader.Close()

#            $results = Invoke-WebRequest -Uri $URI -Credential $Credentials
            $csv_content = $results | ConvertFrom-Csv
            Return $csv_content
        }
        catch {
            $_.Exception
            Write-Host "Couldn't retrieve report"
        }
    }
    Return $null
}

Function Get-WDPhoto {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [alias("ID","EmpID","EmployeeID","ContractorID")]
        [String] $Employee,

        [Parameter(Mandatory = $false)]
        [alias("ContingentWorker","Coop","Co-Op","Intern")]
        [Switch] $Contractor,

        [Parameter(Mandatory = $false)]
        [Switch] $Sandbox,

        [Parameter(Mandatory = $true)]
        [System.Management.Automation.PSCredential] $Credentials

    ) #end param

    if($Credentials)
    {
        $UserName = $Credentials.UserName
        $Password = $Credentials.GetNetworkCredential().password
        if ($Sandbox)
        {
            $URI = "https://wd5-impl-services1.workday.com/ccx/service/manh/Human_Resources/v23.2"
        }
        else
        {
            $URI = "https://wd5-services1.myworkday.com/ccx/service/manh/Human_Resources/v23.2"
        }

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
                    <bsvc:Get_Workers_Request>
                        <bsvc:Request_References>
                            <bsvc:Worker_Reference>
                                <bsvc:ID bsvc:type=""></bsvc:ID>
                            </bsvc:Worker_Reference>
                        </bsvc:Request_References>
                        <bsvc:Response_Filter>
                            <bsvc:As_Of_Entry_DateTime></bsvc:As_Of_Entry_DateTime>
                        </bsvc:Response_Filter>
                        <bsvc:Response_Group>
                            <bsvc:Include_Employment_Information>1</bsvc:Include_Employment_Information>
                            <bsvc:Include_Organizations>1</bsvc:Include_Organizations>
                            <bsvc:Include_Photo>1</bsvc:Include_Photo>
                            <bsvc:Include_User_Account>1</bsvc:Include_User_Account>
                        </bsvc:Response_Group>
                    </bsvc:Get_Workers_Request>
                </soapenv:Body>
            </soapenv:Envelope>'

        $ns = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
        $ns.AddNamespace("bsvc", "urn:com.workday/bsvc")

        $now = Get-Date -Format s

        $xml.Envelope.Header.Security.UsernameToken.Username = $UserName
        $xml.Envelope.Header.Security.UsernameToken.Password.InnerText = $Password
        $xml.Envelope.Body.Get_Workers_Request.Request_References.Worker_Reference.ID.InnerText = $employee
        $xml.Envelope.Body.Get_Workers_Request.Response_Filter.As_Of_Entry_DateTime = $now.ToString()

        if ($Contractor) {
            $xml.Envelope.Body.Get_Workers_Request.Request_References.Worker_Reference.ID.SetAttribute("bsvc:type","Contingent_Worker_ID")
        } else {
            $xml.Envelope.Body.Get_Workers_Request.Request_References.Worker_Reference.ID.SetAttribute("bsvc:type","Employee_ID")
        }

        try {
            [xml]$post = Invoke-WebRequest -Uri $URI -Method Post -Body $xml -ContentType "application/xml"
        }
        catch {
            $result = $_.Exception.Response.GetResponseStream()
            $reader = New-Object System.IO.StreamReader($result)
            [xml]$responseBody = $reader.ReadToEnd();
            Write-Host "Error: $($responseBody.Envelope.Body.Fault.faultstring)"
        }

        $image = [System.Convert]::FromBase64String($post.Envelope.Body.Get_Workers_Response.Response_Data.Worker.Worker_Data.Photo_Data.Image)

        RETURN $image
    }
}

Function Set-WDPhoto {
    [CmdletBinding()]
    param(
          [Parameter(Mandatory = $true)]
          [alias("ID","EmpID","EmployeeID","ContractorID")]
          [String] $Employee,

          [Parameter(Mandatory = $true)]
          [alias("Image","Photo")]
          [String] $Image,

          [Parameter(Mandatory = $true)]
          [String] $Extension,

          [Parameter(Mandatory = $false)]
          [alias("ContingentWorker","Coop","Co-Op","Intern")]
          [Switch] $Contractor,

          [Parameter(Mandatory = $false)]
          [Switch] $Sandbox,

          [Parameter(Mandatory = $true)]
          [System.Management.Automation.PSCredential] $Credentials
    ) #end param

    if($Credentials)
    {
        $UserName = $Credentials.UserName
        $Password = $Credentials.GetNetworkCredential().password

        if ($Sandbox)
        {
            $URI = "https://wd5-impl-services1.workday.com/ccx/service/manh/Human_Resources/v23.2"
        }
        else
        {
            $URI = "https://wd5-services1.myworkday.com/ccx/service/manh/Human_Resources/v23.2"
        }

        $photo = [convert]::ToBase64String($image)

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
              <bsvc:Put_Worker_Photo_Request>
                 <bsvc:Worker_Reference>
                    <bsvc:ID bsvc:type="Employee_ID"></bsvc:ID>
                 </bsvc:Worker_Reference>
                 <bsvc:Worker_Photo_Data>
                    <bsvc:ID></bsvc:ID>
                    <bsvc:Filename></bsvc:Filename>
                    <bsvc:File></bsvc:File>
                 </bsvc:Worker_Photo_Data>
              </bsvc:Put_Worker_Photo_Request>
           </soapenv:Body>
        </soapenv:Envelope>'

        $xml.Envelope.Header.Security.UsernameToken.Username = $UserName
        $xml.Envelope.Header.Security.UsernameToken.Password.InnerText = $Password
        $xml.Envelope.Body.Put_Worker_Photo_Request.Worker_Reference.ID.InnerText = $Employee
        $xml.Envelope.Body.Put_Worker_Photo_Request.Worker_Photo_Data.ID = $Employee
        $xml.Envelope.Body.Put_Worker_Photo_Request.Worker_Photo_Data.FileName = "$Employee.$Extension"
        $xml.Envelope.Body.Put_Worker_Photo_Request.Worker_Photo_Data.File = $photo

        try {
            $post = Invoke-WebRequest -Uri $URI -Method Post -Body $xml -ContentType "application/xml"
        }
        catch {
            $result = $_.Exception.Response.GetResponseStream()
            $reader = New-Object System.IO.StreamReader($result)
            [xml]$responseBody = $reader.ReadToEnd();
            Write-Host "Error: $($responseBody.Envelope.Body.Fault.faultstring)"
        }
    }
}
