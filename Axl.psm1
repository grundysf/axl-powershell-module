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


#  .synopsis
#  Executes an update SQL statement against the Cisco UC database and returns the number of rows updated
#  
#  .example
#  Get-UcSqlUpdate $conn "update device set name = 'yes'"
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
  return [int]$node.innertext
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
  Param(
    [Parameter(Mandatory=$true)][string]$server,
    [Parameter(Mandatory=$true)][string]$user,
    [Parameter(Mandatory=$true)][string]$pass
  )
  $creds = toBase64 "$user`:$pass"
  $url = "https://${server}:8443/axl/"
  $conn = new-object psobject -property @{url=$url; user=$user; pass=$pass; creds=$creds; server=$server}
  return $conn
}

function Get-BuddyList {
<#
  .synopsis
  Gets the contact list (buddy list) for the specified Presence user.  This uses an SQL query since
  there does not appear to be a way to retrieve this information using another API method.
  
  .parameter user
  Presense userid who's contact list will be returned.  SQL wildcard character '%' is allowed.
  
  .parameter axlconn
  Axl connection to a CUP server.  New-AxlConnection may be used to create the connection.
  
  .example
  Get-BuddyList -AxlConn $conn -user "John Doe"
#>
  Param(
    # Connection object created with New-AxlConnection
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

Export-ModuleMember -Function Search-UcLicenseCapabilities, New-AxlConnection, Get-UcSqlQuery, Get-UcSqlUpdate, Get-BuddyList
