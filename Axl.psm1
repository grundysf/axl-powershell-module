#
# Craig Petty, yttep.giarc@gmail.com
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
    [Parameter(Mandatory=$true)][string] $msg
  )
  return [system.convert]::tobase64string([system.text.encoding]::UTF8.getbytes($msg))
}

function Execute-SOAPRequest { 
  Param (
    [Parameter(Mandatory=$true)][psobject]$AxlConn,
    [Parameter(Mandatory=$true)][Xml]$XmlDoc,
    # Write XML request/response to a file for troubleshooting
    [string]$XmlTraceFile
  )
  $ErrorActionPreference = "Stop";

  if ($XmlTraceFile) {
    "`n",$XmlDoc.innerxml | out-file -encoding ascii -append $XmlTraceFile
  }
  
  #write-host "Sending SOAP Request To Server: $URL" 
  $webReq = [System.Net.WebRequest]::Create($AxlConn.url)
  $webReq.Headers.Add("SOAPAction","SOAPAction: CUCM:DB ver=8.5")
  $creds = ConvertFrom-SecureString $AxlConn.creds
  $webReq.Headers.Add("Authorization","Basic "+$creds)

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

  if ($XmlTraceFile) {
    "`n",$ReturnXml.innerxml | out-file -encoding ascii -append $XmlTraceFile
  }

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

# .outputs
#  PsObjects with (userid, enableUps, and enableUpc properties set)
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
  $textNode = $xml.CreateTextNode($searchCriteria)
  $null = $xml.SelectSingleNode("//searchCriteria/userid").prependchild($textNode)
  $retXml = Execute-SOAPRequest $AxlConn $xml

  #$retXml.selectNodes("//userid[../enableUpc = 'true']") 
  $retXml.selectNodes("//licenseCapabilities") | % -begin {$resultArray=@()} {
    $node = new-object psobject
    $node | add-member noteproperty userid $_.selectsinglenode("userid").innertext
    $node | add-member noteproperty enableUpc $_.selectsinglenode("enableUpc").innertext
    $resultArray += $node
  }
  return $resultArray
}

# .synopsis
#  psobject[] returns an array of psobjects with (userid, enableUps, and enableUpc properties set)
function Search-Users {
  Param (
    [Parameter(Mandatory=$true)][psobject]$AxlConn,
    [string]$searchCriteria="%",
    
    # Write XML request/response to a file for troubleshooting
    [string]$XmlTraceFile
  )
  $xml = [xml]@"
<?xml version="1.0" encoding="utf-8"?>
<soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" xmlns:ns="http://www.cisco.com/AXL/API/8.5">
 <soapenv:Header/>
 <soapenv:Body>
  <ns:listUser sequence="1">
   <searchCriteria>
    <userid/>
   </searchCriteria>
   <returnedTags>
    <userid/><firstName/><lastName/>
    <primaryExtension>
     <pattern/><routePartitionName/>
    </primaryExtension>
    <status/>
   </returnedTags>
  </ns:listUser>
 </soapenv:Body>
</soapenv:Envelope>
"@
  $textNode = $xml.CreateTextNode($searchCriteria)
  $null = $xml.SelectSingleNode("//searchCriteria/userid").prependchild($textNode)
  $retXml = Execute-SOAPRequest $AxlConn $xml -xmltracefile $XmlTraceFile

  $nsm = new-object system.xml.XmlNameSpaceManager -ArgumentList $retXml.NameTable
  $nsm.addnamespace("ns", "http://www.cisco.com/AXL/API/8.5")
  $retXml.selectNodes("//ns:listUserResponse/return/user", $nsm) | % {
    $node = new-object psobject
    $node | add-member noteproperty userid $_.selectsinglenode("userid").innertext
    $node | add-member noteproperty firstName $_.selectsinglenode("firstName").innertext
    $node | add-member noteproperty lastName $_.selectsinglenode("lastName").innertext
    # primaryExtension
      $ext = new-object psobject
      $ext | add-member noteproperty pattern $_.selectsinglenode("primaryExtension/pattern").innertext
      $ext | add-member noteproperty routePartitionName $_.selectsinglenode("primaryExtension/routePartitionName").innertext
      $ext | add-member scriptmethod ToString {$this.pattern} -force
      $ext.PSObject.TypeNames.Insert(0,'UcDN')
      $node | add-member noteproperty primaryExtension $ext
    $node.PSObject.TypeNames.Insert(0,'UcUser')
    $node
  }
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
  The password (in SecureString format) required to authenticate to the AXL server.  
  Use (Read-Host -AsSecureString -Prompt Password) to create a SecureString.  If you do not supply
  this parameter, then you will be prompted for the password.
  
  .example
  $cucm = New-AxlConnection cm1.example.com admin mypass
  
#>
  Param(
    [Parameter(Mandatory=$true)][string]$Server,
    [Parameter(Mandatory=$true)][string]$User,
    [System.Security.SecureString]$Pass
  )
  if (!$Pass) {$Pass = Read-Host -AsSecureString -Prompt "Password for ${User}@${Server}"}

  $clearpass = ConvertFrom-SecureString $pass
  $clearcreds = toBase64 "$user`:$clearpass"
  $creds = ConvertTo-SecureString $clearcreds

  $url = "https://${server}:8443/axl/"
  $conn = new-object psobject -property @{
    url=$url; 
    user=$user;
    pass=$pass;
    creds=$creds;
    server=$server
  }
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

function Add-Buddy {
<#
  .synopsis
  Adds a conatct (aka buddy) to someone's contact list.
  
  .description
  There is no AXL method to manipulate contact list entries, so this function uses SQL to modify the database
  directly.  Use at your own risk.
  
  .parameter AxlConn
  Connection object created with New-AxlConnection.

  .parameter UserID
  userid
  
  .parameter BuddyAddr
  SIP address of the contact to delete
  
  .parameter GroupName
  Contact list group, as seen in the Jabber client, that will contains the new contact.  Note, group names
  in CUP are case-sensitive.
  
  .parameter BuddyNickName
  Display name for this contact.  By default this is set to the BuddyAddr
  
  .example
  Add-Buddy asmithee jane.doe@example.com "My Contacts"

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
    $GroupName,
    
    [Parameter(ValueFromPipelineByPropertyName=$true)]
    [alias("nickname")][string]
    [string]
    $BuddyNickName
  )
  $sql = @"
INSERT INTO RosterSyncQueue 
(userid, buddyjid, buddynickname, groupname, action) 
VALUES ('${UserID}', '${BuddyAddr}', '${BuddyNickName}', '${GroupName}', 1)
"@
  write-verbose "sql = $sql"
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

  .parameter force
  Also show hidden contacts.  That is, contacts which are not in a group and do not appear in
  the Jabber client contact list.
  
  .example
  Get-BuddyList -AxlConn $conn -user "John Doe"
#>
  Param(
    [Parameter(Mandatory=$true)][psobject]$AxlConn,
    [Parameter(Mandatory=$true)][string]$user,
    [switch]$Force
  )
  $sql = @"
select u.userid,r.contact_jid,r.nickname,r.state,g.group_name 
from rosters r join enduser u on u.xcp_user_id = r.user_id 
left outer join groups g on g.roster_id = r.roster_id 
where lower(u.userid) like lower('${user}')
order by r.contact_jid
"@
  foreach ($row in Get-UcSqlQuery $AxlConn $sql) {
    $row.PSObject.TypeNames.Insert(0,'UcBuddy')
    if ($row.group_name.length -gt 0 -or $force) {
      $row
    }
  }
}

function ConvertFrom-SecureString {
# .synopsis
#  Converts a SecureString to a string
#
# .outputs
#  String
  Param(
    [Parameter(Mandatory=$true, ValueFromPipeline=$true)]
    [System.Security.SecureString[]]
    $secure
  )
  process {
    foreach ($ss in $secure) {
      $val = [System.Runtime.InteropServices.Marshal]::SecureStringToGlobalAllocUnicode($ss)
      [System.Runtime.InteropServices.Marshal]::PtrToStringUni($val)
      [System.Runtime.InteropServices.Marshal]::ZeroFreeGlobalAllocUnicode($val)
    }
  }
}

function ConvertTo-SecureString {
# .synopsis
#  Converts a string to a SecureString
# .outputs
#  System.Security.SecureString
  Param(
    [Parameter(Mandatory=$true, ValueFromPipeline=$true)]
    [string]
    $PlainString
  )
  Process {
    $PlainString.tochararray() | foreach-object `
      -Begin {$ss = new-object System.Security.SecureString} `
      { $ss.appendchar($_) } `
      -End { $ss }
  }
}

Export-ModuleMember -Function Search-UcLicenseCapabilities, New-AxlConnection, `
  Get-UcSqlQuery, Get-UcSqlUpdate, Get-BuddyList, Remove-Buddy, Add-buddy, `
  Search-Users
