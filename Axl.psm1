#
# Craig Petty
# Digital Generation, Inc.
# 2013
#

#https://[CCM-IP-ADDRESS]:8443/axl/

add-type @"
public struct contact {
  public string First;
  public string Last;
  public string Phone;
}
"@

function toBase64 {
  Param (
    [parameter(Mandatory=$true)][string] $msg
  )
  return [system.convert]::tobase64string([system.text.encoding]::UTF8.getbytes($msg))
}

function Execute-SOAPRequest { 
  Param (
    [Parameter(Mandatory=$true)][psobject]$AxlConn,
    [Xml] $XmlDoc
  )
  $ErrorActionPreference = "Stop";
  
  #write-host "Sending SOAP Request To Server: $URL" 
  $webReq = [System.Net.WebRequest]::Create($AxlConn.url)
  $webReq.Headers.Add("SOAPAction","SOAPAction: CUCM:DB ver=8.5")
  $webReq.Headers.Add("Authorization","Basic "+$AxlConn.creds)

  #$cred = get-credential
  #$webReq.Credentials = new-object system.net.networkcredential @($cred.username, $cred.password)
  #$webReq.PreAuthenticate = $true

  $webReq.ContentType = "text/xml;charset=`"utf-8`""
  $webReq.Accept      = "text/xml"
  $webReq.Method      = "POST"
  
  #write-host "Initiating Send."
  $requestStream = $webReq.GetRequestStream()
  $XmlDoc.Save($requestStream)
  $requestStream.Close()
  
  #write-host "Send Complete, Waiting For Response."
  $resp = $webReq.GetResponse()
  $responseStream = $resp.GetResponseStream()
  $soapReader = [System.IO.StreamReader]($responseStream)
  $ReturnXml = [Xml] $soapReader.ReadToEnd()
  $responseStream.Close()

  # check to see if a fault occurred
  $nsm = new-object system.xml.XmlNameSpaceManager -ArgumentList $xml.NameTable
  $nsm.addnamespace("soapenv", "http://schemas.xmlsoap.org/soap/envelope/")
  $faultnode = $xml.selectsinglenode("//soapenv:Fault/faultstring", $nsm)
  if ($faultnode -ne $null) {
    throw "SOAP fault: $($faultnode.innertext)"
  }

  write-verbose "Response Received."
  write-verbose $ReturnXml.InnerXML
  return $ReturnXml
}


<#
  .synopsis
  Executes an update SQL statement against the Cisco UC database and returns the number of rows updated

  .outputs
  System.String.  The number of rows updated

  .example
  Get-UcSqlUpdate $conn "update device set name = 'yes'"
#>
function Get-UcSqlUpdate {
  Param (
    # Connection object created with New-AxlConnection
    [Parameter(Mandatory=$true)][psobject]$AxlConn,

    # The text of the SQL statement to execute
    [string]$sqlText
  )
  $xml = [xml]@"
<?xml version="1.0" encoding="utf-8"?>
<soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" xmlns:ns="http://www.cisco.com/AXL/API/8.5">
  <soapenv:Header/>
  <soapenv:Body>
    <ns:executeSQLUpdate sequence="1">
      <sql/>
    </ns:executeSQLUpdate>
  </soapenv:Body>
</soapenv:Envelope>
"@
  $textNode = $xml.CreateTextNode($sqlText)
  $null = $xml.SelectSingleNode("//sql").prependchild($textNode)
  $retXml = Execute-SOAPRequest $AxlConn $xml

  $node = $retXml.selectSingleNode("//return/rowsUpdated")
  [int]$node.innertext
  remove-variable retXml
}



function Get-UcSqlQuery {
  Param (
    [Parameter(Mandatory=$true)][psobject]$AxlConn,
    [string]$sqlText
  )
  $xml = [xml]@"
<?xml version="1.0" encoding="utf-8"?>
<soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" xmlns:ns="http://www.cisco.com/AXL/API/8.5">
  <soapenv:Header/>
  <soapenv:Body>
    <ns:executeSQLQuery sequence="1">
      <sql/>
    </ns:executeSQLQuery>
  </soapenv:Body>
</soapenv:Envelope>
"@
  $textNode = $xml.CreateTextNode($sqlText)
  $null = $xml.SelectSingleNode("//sql").prependchild($textNode)
  $retXml = Execute-SOAPRequest $AxlConn $xml
  $resultArray=@()
  foreach ($row in $retXml.selectNodes("//return/row")) {
    $rowobj = new-object psobject
    foreach ($prop in $($row | get-member -membertype property)) {
      $rowobj | add-member noteproperty $prop.name $row.item($prop.name).innertext
    }
    $resultArray += $rowobj
  }

  remove-variable retXml
  return $resultArray
}

# returns an array of psobjects with (userid, enableUps, and enableUpc properties set)
function Search-UcLicenseCapabilities {
  Param (
    [Parameter(Mandatory=$true)][psobject]$AxlConn,
    [string]$searchCriteria="%"
  )
  $xml = [xml]@"
<?xml version="1.0" encoding="utf-8"?>
<soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" xmlns:ns="http://www.cisco.com/AXL/API/8.5">
  <soapenv:Header/>
  <soapenv:Body>
    <ns:listLicenseCapabilities sequence="?">
      <searchCriteria>
        <userid/>
      </searchCriteria>
      <returnedTags>
        <userid/><enableUpc/>
      </returnedTags>
    </ns:listLicenseCapabilities>
  </soapenv:Body>
</soapenv:Envelope>
"@
  $textNode = $xml.CreateTextNode($sqlText)
  $null = $xml.SelectSingleNode("//searchCriteria/userid").prependchild($textNode)
  $retXml = Execute-SOAPRequest $AxlConn $xml

  #$retXml.selectNodes("//userid[../enableUpc = 'true']") 
  $retXml.selectNodes("//licenseCapabilities") | % -begin {$resultArray=@()} {
    $node = new-object psobject
    $node | add-member noteproperty userid $_.selectsinglenode("userid").innerxml
    $node | add-member noteproperty enableUpc $_.selectsinglenode("enableUpc").innerxml
    $resultArray += $node
  }
  return $resultArray
}

function New-AxlConnection {
<#
  .synopsis
  Specify AXL server connection parameters
  
  .description
  Use this command to specify the information necessary to connect to an AXL server such as Cisco Unified 
  Communications Manager (CUCM), or Cisco Unified Presense Server (CUPS).  Assign the output of this 
  command to a variable so it can be used with other commands in this module that require it.
  
  The SSL certificate on the AXL server must be trusted by the computer runing your powershell script.
  If the certificate is self-signed, then you will need to add it to your trusted certificate store.
  If the certificate is signed by a CA, then you will need to make sure the CA certificate is added to
  your trusted certificate store.  This also means that the name you use to connect to the server (i.e.
  the 'server' parameter of this command) must match the subject name on the AXL server certificate or 
  must match one of the subject alternative names (SAN) if the certificate has SANs.
  
  .parameter Server
  Supply only the hostname (fqdn or IP) of the AXL server.  The actual URL used to connect to the server 
  is generated automatically based on this template: "https://${Server}:8443/axl/".  
  
  .parameter User
  The username required to authenticate to the AXL server
  
  .parameter Pass
  The password required to authenticate to the AXL server
  
  .example
  $cucm = New-AxlConnection cm1.example.com admin mypass
  
#>
  Param(
    [Parameter(Mandatory=$true)][string]$Server,
    [Parameter(Mandatory=$true)][string]$User,
    [Parameter(Mandatory=$true)][string]$Pass
  )
  $creds = toBase64 "$user`:$pass"
  $url = "https://${server}:8443/axl/"
  $conn = new-object psobject -property @{url=$url; user=$user; pass=$pass; creds=$creds; server=$server}
  return $conn
}

function Remove-Buddy {
<#
  .synopsis
  Deletes a conatct (aka buddy) from someone's contact list.
  
  .description
  There is no AXL method to remove contact list entries, so this function uses SQL to manipulate the database
  directly.  Use at your own risk.
  
  .parameter AxlConn
  Connection object created with New-AxlConnection.

  .parameter UserID
  userid
  
  .parameter BuddyAddr
  SIP address of the contact to delete
  
  .parameter GroupName
  Contact list group, as seen in the Jabber client, that contains the contact to be deleted.  Note, group names
  in CUP are case-sensitive.
  
  .example
  Remove-Buddy asmithee jane.doe@example.com "My Contacts"
  
  .example
  Get-BuddyList $conn asmith | ? {$_.contact_jid -match "jdoe"} | remove-buddy $conn
  
  Descripton
  ===================
  Remove entries from Amy Smith's (userid=asmith) contact list where the contact's address 
  matches regular expression "jdoe"
#>
  Param(
    [Parameter(Mandatory=$true)][psobject]
    $AxlConn,
    
    [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
    [string]
    $UserID,
    
    [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
    [alias("contact_jid")][string]
    $BuddyAddr,
    
    [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
    [alias("group_name")][string]
    $GroupName
  )
  $sql = @"
INSERT INTO RosterSyncQueue 
(userid, buddyjid, buddynickname, groupname, action) 
VALUES ('${UserID}', '${BuddyAddr}', '', '${GroupName}', 3)
"@
  $null = Get-UcSqlUpdate $AxlConn $sql
}

function Get-BuddyList {
<#
  .synopsis
  Gets the contact list (buddy list) for the specified Presence user.  This uses an SQL query since
  there does not appear to be a way to retrieve this information using another API method.
  
  .parameter user
  Presense userid who's contact list will be returned.  SQL wildcard character '%' is allowed.
  
  .parameter axlconn
  Connection object created with New-AxlConnection.
  
  .example
  Get-BuddyList -AxlConn $conn -user "John Doe"
#>
  Param(
    [Parameter(Mandatory=$true)][psobject]$AxlConn,
    
    [Parameter(Mandatory=$true)][string]$user
  )
  $sql = @"
select u.userid,r.contact_jid,r.nickname,g.group_name 
from groups g 
inner join enduser u on u.xcp_user_id = g.user_id 
inner join rosters r on g.roster_id = r.roster_id 
where u.userid like '${user}'
order by g.group_name,r.contact_jid
"@
  Get-UcSqlQuery $AxlConn $sql
}

Export-ModuleMember -Function Search-UcLicenseCapabilities, New-AxlConnection, Get-UcSqlQuery, Get-UcSqlUpdate, Get-BuddyList, Remove-Buddy
