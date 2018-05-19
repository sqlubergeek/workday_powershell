Function Get-FileName($initialDirectory)
{
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
#    $OpenFileDialog.filter = "BMP|*.bmp|GIF|*.gif|JPG|*.jpg|PNG|*.png" 
    $OpenFileDialog.ShowDialog() | Out-Null
    $OpenFileDialog.filename
}

[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null


 
$empid = [Microsoft.VisualBasic.Interaction]::InputBox("Enter the Employee Number", "Employee Number", "")
if ($empid){
    $filename = Get-FileName "C:\"

    if ($filename){

        [byte[]]$image =Get-Content $filename -Encoding byte

        $uri = "https://wd5-services1.myworkday.com/ccx/service/manh/Human_Resources/v23.2"

        $credentials = Get-Credential

        $wd_username = $Credentials.UserName
        $wd_password = $Credentials.GetNetworkCredential().password


        $photo = [convert]::ToBase64String($image)
        $wd_empid = $empid.TrimStart("0")

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

        #Timestamp for request signing
        $timestamp = [DateTime]::UtcNow.ToString("yyyyMMddTHH:mm:sszz00")
        #GUID for request signing
        $nonce = [GUID]::NewGuid()

        $xml.Envelope.Header.Security.UsernameToken.Username = $wd_username
        $xml.Envelope.Header.Security.UsernameToken.Password.InnerText = $wd_password
        $xml.Envelope.Body.Put_Worker_Photo_Request.Worker_Reference.ID.InnerText = $wd_empid
        $xml.Envelope.Body.Put_Worker_Photo_Request.Worker_Photo_Data.ID = $wd_empid
        $xml.Envelope.Body.Put_Worker_Photo_Request.Worker_Photo_Data.FileName = $empid + '.jpg'
        $xml.Envelope.Body.Put_Worker_Photo_Request.Worker_Photo_Data.File = $photo

        try {
	        $post = Invoke-WebRequest -Uri $uri -Method Post -Body $xml -ContentType "application/xml"
        }
        catch {
	        $_.EXCEPTION
        Write-Host "Couldn't upload photo for " + $empid
        }
    }
}