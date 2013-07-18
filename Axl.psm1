#
# Craig Petty, yttep.giarc@gmail.com
# Digital Generation, Inc.
# 2013
#
$ErrorActionPreference = "Stop";

#https://[CCM-IP-ADDRESS]:8443/axl/

$cache = @{
  enduserpkid=@{}
}
function get-cache {
  Param(
    [parameter(mandatory=$true)]
    [ValidateSet("enduserpkid")]
    $CacheSet
    ,
    [parameter(mandatory=$true)]
    $Key
  )
  $cache.$CacheSet.$key
}
function set-cache {
  Param(
    [parameter(mandatory=$true)]
    [ValidateSet("enduserpkid")]
    $CacheSet
    ,
    [parameter(mandatory=$true)]
    $Key
    ,
    [parameter(mandatory=$true)]
    $Value
  )
  $cache.$CacheSet.$key = $value
}

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
  [cmdletbinding(DefaultParameterSetName="Inline")]
  Param (
    [Parameter(Mandatory=$true,Position=0)][psobject]$AxlConn
    ,
    [Parameter(ParameterSetName="Inline",Mandatory=$true,Position=1)]
    [string]$sqlText
    ,
    [Parameter(ParameterSetName="File",Mandatory=$true)]
    [string]$File,
    [string]$XmlTraceFile
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
  if ($PSCmdlet.ParameterSetName -eq "Inline") {
    $textNode = $xml.CreateTextNode($sqlText)
  }
  elseif ($PSCmdlet.ParameterSetName -eq "File") {
    $textNode = $xml.CreateTextNode( ((get-content $File) -join "`n") )
  }
  $null = $xml.SelectSingleNode("//sql").prependchild($textNode)
  $retXml = Execute-SOAPRequest $AxlConn $xml -xmltracefile $XmlTraceFile
  $resultArray=@()

  $nsm = new-object system.xml.XmlNameSpaceManager -ArgumentList $retXml.NameTable
  $nsm.addnamespace("ns", "http://www.cisco.com/AXL/API/8.5")

  $base = $retXml.selectSingleNode("//return")
  foreach ($row in $retXml.selectNodes("//ns:executeSQLQueryResponse/return/row", $nsm)) {
    $rowobj = new-object psobject
    foreach ($column in $row.ChildNodes) {
      $rowobj | add-member noteproperty $column.name $column.innertext
    }
    $rowobj
  }
  remove-variable retXml
}

<#
  .synopsis
  Convert Informix column type ID to a string
  
  .link
  http://publib.boulder.ibm.com/infocenter/idshelp/v10/index.jsp?topic=/com.ibm.gsg.doc/gsg140.htm
#>
function ConvertFrom-ColType {
  Param(
    [Parameter(Mandatory=$true)]
    [int][ValidateRange(0,0x8fff)]
    $coltype,
    $extended_name
  )

  $basetype = $coltype -band 0xff
  $t = ""
  switch ($basetype) {
    0  {$t="char"; break}
    1  {$t="smallint"; break}
    2  {$t="int"; break}
    3  {$t="float"; break}
    4  {$t="smallfloat"; break}
    5  {$t="decimal"; break}
    6  {$t="serial"; break}
    7  {$t="date"; break}
    8  {$t="money"; break}
    9  {$t="null"; break}
    10 {$t="datetime"; break}
    11 {$t="byte"; break}
    12 {$t="text"; break}
    13 {$t="varchar"; break}
    14 {$t="interval"; break}
    15 {$t="nchar"; break}
    16 {$t="nvarchar"; break}
    17 {$t="int8"; break}
    18 {$t="serial8"; break}
    19 {$t="set"; break}
    20 {$t="multiset"; break}
    21 {$t="list"; break}
    22 {$t="row"; break}
    23 {$t="collection"; break}
    24 {$t="rowref"; break}
    40 {
      $t="user-def vary-len"
      if ($extended_name) {$t=$extended_name}
      break
    }
    41 {
      $t="user-def fix-len"
      if ($extended_name) {$t=$extended_name}
      break
    }
    42 {$t="refser8"; break}
    default {$t="unknown (type $basetype)"}
  }
  
  if ($coltype -band 0x100) {$t += " not null"}
  if ($coltype -band 0x200) {$t += " hostvar"}
  if ($coltype -band 0x400) {$t += " netflt"}
  if ($coltype -band 0x800) {$t += " distinct"}
  if ($coltype -band 0x1000) {$t += " named"}
  if ($coltype -band 0x2000) {$t += " distinct"}
  if ($coltype -band 0x4000) {$t += " distinct"}
  if ($coltype -band 0x8000) {$t += " on client"}
  
  $t
}

# .outputs
#  PsObjects with (userid, enableUps, and enableUpc properties set)
function Get-UcLicenseCapabilities {
  Param (
    [Parameter(Mandatory=$true)][psobject]$AxlConn,
    [string]$searchCriteria="%"
  )
# Alternatively we could run this SQL query against the CUP server
# select u.userid,l.enablecups,l.enablecupc 
# from enduser u join enduserlicense l on l.fkenduser=u.pkid"
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
function Get-UcUser {
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
      $ext.PSObject.TypeNames.Insert(0,'UcExtension')
      $node | add-member noteproperty primaryExtension $ext
    $node.PSObject.TypeNames.Insert(0,'UcUser')
    $node
  }
}

# .outputs
#  psobject[] returns an array of psobjects with (userid, enableUps, and enableUpc properties set)
function Get-UcDialPlans {
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
  <ns:listDialPlan sequence="1">
   <searchCriteria>
    <name/>
   </searchCriteria>
   <returnedTags>
    <name/><description/>
    <status/>
   </returnedTags>
  </ns:listDialPlan>
 </soapenv:Body>
</soapenv:Envelope>
"@
  $textNode = $xml.CreateTextNode($searchCriteria)
  $null = $xml.SelectSingleNode("//searchCriteria/name").prependchild($textNode)
  $retXml = Execute-SOAPRequest $AxlConn $xml -xmltracefile $XmlTraceFile
  
  $nsm = new-object system.xml.XmlNameSpaceManager -ArgumentList $retXml.NameTable
  $nsm.addnamespace("ns", "http://www.cisco.com/AXL/API/8.5")
  $retXml.selectNodes("//ns:listDialPlanResponse/return/dialPlan", $nsm) | % {
    $node = new-object psobject
    $node | add-member noteproperty name $_.selectsinglenode("name").innertext
    $node | add-member noteproperty description $_.selectsinglenode("description").innertext
    $node.PSObject.TypeNames.Insert(0,'UcDialPlan')
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

function Remove-UcBuddy {
<#
  .synopsis
  Deletes a conatct (aka buddy) from someone's contact list.
  
  .description
  There is no AXL method to remove contact list entries, so this function uses SQL to manipulate the database
  directly.  Use at your own risk.
  
  .parameter AxlConn
  Connection object created with New-AxlConnection.

  .parameter Owner
  Owner (userid) who's buddy list contains the buddy to be removed.  This is not the full SIP address, but just 
  the CUP username.
  
  .parameter Buddy
  SIP address of the contact to delete
  
  .parameter Group
  Contact list group name, as seen in the Jabber client, that contains the contact to be deleted.  Note, group 
  names in CUP are case-sensitive.
  
  .example
  Remove-UcBuddy asmithee jane.doe@example.com "My Contacts"
  
  .example
  Get-UcBuddy $conn asmith | ? {$_.contact_jid -match "jdoe"} | remove-UcBuddy $conn
  
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
    $Owner,
    
    [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
    [alias("contact_jid")][string]
    $Buddy,
    
    [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
    [AllowEmptyString()][alias("group_name")][string]
    $Group
  )
  Process {
    $group = escapeSql $group
    $sql = @"
INSERT INTO RosterSyncQueue 
(userid, buddyjid, buddynickname, groupname, action) 
VALUES ('${Owner}', '${Buddy}', '', '${Group}', 3)
"@
    write-verbose "SQL: $sql"
    $null = Get-UcSqlUpdate $AxlConn $sql
  }
}

function escapeSql {
  Param([parameter(mandatory=$true)]
  [AllowEmptyString()][string]$msg)
  $msg -replace "'","''"
}

function Get-Table {
<#
  .synopsis
  Show table names from the SQL databse, or show details of a specific table.
#>
  [CmdletBinding(DefaultParameterSetName="List")]
  Param(
    # Connection object created with New-AxlConnection.
    [Parameter(Mandatory=$true,Position=0)]
    $AxlConn
    ,
    # Show column details for this named table
    [Parameter(ParameterSetName="Detail",Position=1)]
    $Name
    ,
    # Include system tables
    [Parameter(ParameterSetName="List",Position=1)]
    [switch]$IncludeSystem
  )
  if ($Name) {
    # List details for the specified table

    # first get a list of User-Defined Types
    $sql = @"
select
 colno no,
 colname name,
 coltype typecode,
 "" typename,
 xt.name xtypename,
 collength bytes
from
 syscolumns c
 join systables t on t.tabid=c.tabid
 left outer join sysxtdtypes xt on c.extended_id=xt.extended_id
where lower(t.tabname)=lower('${Name}')
order by
 colno
"@
    foreach ($row in Get-UcSqlQuery $AxlConn $sql) {
      $row.PSObject.TypeNames.Insert(0,'UcTableColumn')
      $row.typename = convertfrom-coltype $row.typecode $row.xtypename
      $row | select no,name,typename,bytes
    }
  }
  else {
    # No table specified.  List all tables
    # Hide system tables by default
    $startId = 100;
    if ($IncludeSystem) {$startId = 1}
    $sql = "select tabid id,tabname name,tabtype type from systables where tabtype in ('T','V') and tabid >= ${startId}"
    foreach ($row in Get-UcSqlQuery $AxlConn $sql) {
      $row.PSObject.TypeNames.Insert(0,'UcTable')
      $row
    }
  }
}

function Get-UcCupUsers {
<#
  .synopsis
  Get a list of CUP-enabled users
  
  .outputs
  [string] List of CUP-enabled userid's
  
  .parameter AxlConn
  Connection object created with New-AxlConnection.
  
  .parameter User
  Only return CUP users matching the specified pattern.  SQL wildcard charater % is allowed.
#>
  [CmdletBinding()]
  Param(
    [Parameter(Mandatory=$true)]
    $AxlConn,
    [ValidateScript({$_ -notmatch "[ ']"})]
    $User="%"
  )
  $sql = @"
select 
 userid
from 
 enduser e 
 join enduserlicense l on e.pkid=l.fkenduser 
where
 enablecups = 't'
 and enablecupc = 't'
 and userid like '${User}'
order by userid
"@
  foreach ($row in Get-UcSqlQuery $AxlConn $sql) {
    $row.userid
  }
}


function Get-StoredProc {
<#
  .synopsis
  Show stored procedure names from the SQL databse, or show the definition of the specified SP.
#>
  [CmdletBinding(DefaultParameterSetName="List")]
  Param(
    # Connection object created with New-AxlConnection.
    [Parameter(Mandatory=$true,Position=0)]
    $AxlConn
    ,
    # Show column details for this named table
    [Parameter(ParameterSetName="Detail",Position=1)]
    $Name
  )
  if ($Name) {
    # List definition for the specified stored proc
    $sql = @"
select
 b.data
from 
 sysprocbody b
 join sysprocedures p on b.procid = p.procid
where
 b.datakey = 'T' 
 and lower(p.procname) = lower('${Name}')
"@
    (get-ucsqlquery $cups $sql | select -ExpandProperty data) -join ""
  }
  else {
    # No SP specified.  List all stored procecures
    $sql = "select procid, procname from sysprocedures"
    foreach ($row in Get-UcSqlQuery $AxlConn $sql) {
      $row.PSObject.TypeNames.Insert(0,'UcStoredProc')
      $row
    }
  }
}


function Add-UcBuddy {
<#
  .synopsis
  Adds a conatct (aka buddy) to someone's contact list.
  
  .description
  There is no AXL method to manipulate contact list entries, so this function uses SQL to modify the database
  directly.  Use at your own risk.
  
  .parameter AxlConn
  Connection object created with New-AxlConnection.

  .parameter Owner
  Owner (userid) who's buddy list contains the buddy to be removed.  This is not the full SIP address, but just 
  the CUP username.
  
  .parameter Buddy
  SIP address of the contact to delete
  
  .parameter Group
  Group name as seen in the Jabber client that will contain the new contact.  Note, group names
  in CUP are case-sensitive.
  
  .parameter Nickname
  Display name for this contact.  By default this is set to the BuddyAddr
  
  .example
  Add-UcBuddy asmithee jane.doe@example.com "My Contacts"

#>
  Param(
    [Parameter(Mandatory=$true)][psobject]
    $AxlConn,
    
    [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
    [string]$Owner,
    
    [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
    [ValidatePattern("^(?!sip:)\S+@\S+\.\S+")]
    [string]$Buddy,
    
    [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
    [AllowEmptyString()][string]$Group,
    
    [Parameter(ValueFromPipelineByPropertyName=$true)]
    [string]$Nickname
  )
  Process {
<#
select * from typerostersyncaction
enum name                    moniker
---- ----                    -------
1    Add a new contact       CONTACT_ADD
2    Change contact nickname CONTACT_MODIFY
3    Delete a contact        DELETE_CONTACT
4    Add a new group         GROUP_ADD
5    Change group name       GROUP_MODIFY
6    Delete a group          GROUP_DELETE
7    Subscribe for presence  CONTACT_SUBSCRIBE
#>
    $nickname = escapeSql $nickname
    $group = escapeSql $group
    $sql = @"
INSERT INTO RosterSyncQueue 
(userid, buddyjid, buddynickname, groupname, action) 
VALUES ('${Owner}', '${Buddy}', '${Nickname}', '${Group}', 1)
"@
    write-verbose "sql = $sql"
    $null = Get-UcSqlUpdate $AxlConn $sql
  }
}

function Rename-UcBuddy {
<#
  .synopsis
  Replaces a contact with another, preserving group membership
  
  .description
  There are no AXL methods to add/remove/modify contact list entries, so this function uses SQL 
  to manipulate the database directly.  Use at your own risk.
  
  .parameter AxlConn
  Connection object created with New-AxlConnection.

  .parameter Owner
  CUP username of the person who's contact list you want to modify

  .parameter ForAllUsers
  Rename this contact in all contact lists for all Jabber users in the cluster.  This parameter
  cannot be used together with the Owner parameter.
  
  .parameter Buddy
  SIP address of the contact to rename
  
  .parameter NewAddr
  New SIP address for the contact
  
  .parameter Nickname
  The display name for the contact.
    
  .example
  Rename-UcBuddy $conn asmithee jane.doe@example.com jdoe@contoso.com "Jame Doe"
  
  .outputs
  [string] in -ForAllUsers mode, outputs a list of owner SIP addresses who's buddy list 
  was modified.  Otherwise, outputs nothing.
  
#>
  [CmdletBinding(DefaultParameterSetName="SingleOwner")]
  Param(
    [Parameter(Mandatory=$true, Position=0)][psobject]
    $AxlConn,
    
    [Parameter(ParameterSetName="SingleOwner", Mandatory=$true, Position=1)]
    [string]$Owner,
    
    [Parameter(ParameterSetName="AllOwners", Mandatory=$true, Position=1)]
    [switch]$ForAllUsers,

    [Parameter(Mandatory=$true, Position=2)]
    [ValidatePattern("^(?!sip:)\S+@\S+\.\S+")]
    [string]$Buddy,
    
    [Parameter(Position=3)]
    [ValidatePattern("^(?!sip:)\S+@\S+\.\S+")]
    [string]$NewAddr,

    [string]$Nickname,
    [switch]$Silent
  )
  if ($PSCmdlet.ParameterSetName -eq "AllOwners") {
    $owners = @(Get-UcWatcher -AxlConn $AxlConn $Buddy)
    write-verbose "Found $($owners.count) users watching $Buddy"
  }
  elseif ($PSCmdlet.ParameterSetName -eq "SingleOwner") {
    $owners = @($Owner)
  }
  foreach ($o in $owners) {
    if (-not $silent) {write-host $o}
    $buddylist = @(Get-UcBuddy -axlconn $axlconn $o)
    $oldBudFound = [bool]($buddylist | ? {$_.buddy -eq $buddy})
    $newBudFound = [bool]($buddylist | ? {$_.buddy -eq $newaddr})
    if (-not $oldBudFound) {
      write-warning "Buddy list for $o does not contain ${Buddy}"
      return
    }
    foreach ($bud in $buddylist | ? {$_.buddy -eq $buddy}) {
      if (-not $silent) {write-host "  Removing $buddy from group '$($bud.group)'"}
      $bud | remove-UcBuddy -axlconn $axlconn
      if (-not $newBudFound) {
        $bud.buddy = $newaddr
        if ($Nickname) {$bud.nickname = $nickname}
        if (-not $silent) {write-host "  Adding ${newaddr} to group '$($bud.group)'"}
        $bud | add-UcBuddy -axlconn $axlconn
      }
      else {
        if (-not $silent) {write-host "  $newaddr is already in list"}
      }
    }
    if ($silent) {$o}
  }
}

function Get-UcBuddy {
<#
  .synopsis
  Gets the contact list (buddy list) for the specified Presence user.  This uses an SQL query since
  there does not appear to be a way to retrieve this information using another API method.
  
  .parameter user
  Presense userid who's contact list will be returned.  SQL wildcard character '%' is allowed.
  
  .parameter buddy
  Return only buddes matching the given SIP address.  SQL wildcard character '%' is allowed.  
  If this parameter is omitted, then all buddies for the specified user will be returned.
  
  .parameter axlconn
  Connection object created with New-AxlConnection.

  .parameter force
  Also show hidden contacts.  That is, contacts which are not in a group and do not appear in
  the Jabber client contact list.
  
  .example
  Get-UcBuddy -AxlConn $conn -user "John Doe"
#>
  Param(
    [Parameter(Mandatory=$true)][psobject]$AxlConn,
    [Parameter(Mandatory=$true)][string]$user,
    [ValidateScript({$_ -notmatch "[ ']"})]$buddy="%",
    [switch]$Force
  )
  $sql = @"
select 
 u.userid Owner,
 r.contact_jid Buddy,
 r.nickname Nickname,
 r.state State,
 g.group_name Group
from 
 rosters r 
 join enduser u on u.xcp_user_id = r.user_id 
 left outer join groups g on g.roster_id = r.roster_id 
where
 lower(u.userid) like lower('${user}')
 and lower(r.contact_jid) like lower('${buddy}')
order by
 u.userid,r.contact_jid
"@
  write-verbose "SQL $sql"
  foreach ($row in Get-UcSqlQuery $AxlConn $sql) {
    $row.PSObject.TypeNames.Insert(0,'UcBuddy')
    if ($row.group.length -gt 0 -or $force) {
      $row
    }
  }
}
function Get-UcDN {
<#
  .synopsis
  Gets directory numbers (DN's) defined on UCM.  Uses an SQL query since there does not appear to be a
  way to retrieve this information using another Axl API method.
  
  .parameter DN
  DN to retrieve. SQL wildcard character '%' is allowed.  The default is '%' which shows all DNs
  
  .parameter Partition
  Get DN's in this route partition. SQL wildcard character '%' is allowed.  The default is '%' which shows all DNs

  .parameter Description
  DN to retrieve. SQL wildcard character '%' is allowed.  The default is '%' which shows all DNs

  .parameter axlconn
  Connection object created with New-AxlConnection.
#>
  Param(
    [Parameter(Mandatory=$true)][psobject]$AxlConn,
    [string]$DN="%",
    [string]$Partition="%",
    [string]$Description="%"
  )
  $sql = @"
SELECT
 np.dnorpattern DN,
 rp.name Partition,
 np.Description
FROM
 NumPlan np
 join routepartition rp on np.fkroutepartition=rp.pkid
WHERE 1=1
 and np.dnorpattern like '${DN}'
 and lower(rp.name) like lower('${Partition}')
 and lower(np.description) like lower('${Description}')
order by np.dnorpattern,rp.name
"@
  foreach ($row in Get-UcSqlQuery $AxlConn $sql) {
    $row.PSObject.TypeNames.Insert(0,'UcDN')
    $row
  }
}

function Get-UcUserAcl {
<#
  .synopsis
  Get a user's access-list (allow/block list)
  
  .parameter User
  Show access list for users matching the specified user name.  SQL wildcard character % is allowed.

  .parameter Uri
  Only show records matching the specified Uri.  SQL wildcard character % is allowed.
  The default is % which shows all records
#>
  Param(
    [Parameter(Mandatory=$true)]$AxlConn,
    [Parameter(Mandatory=$true)]$User,
    [ValidateScript({$_ -notmatch "[ ']"})]$Uri="%"
  )
  $sql = @"
select 
 eu.userid User,
 eu.pkid _userPkid,
 a.requesttouri Uri,
 p.name Policy
from
 peuriacl a
 join enduser eu on a.fkenduser = eu.pkid
 join epasprivacypolicy p on a.fkepasprivacypolicy = p.pkid
where 
  eu.userid like '${User}'
  and a.requesttouri like '${Uri}'
"@
  foreach ($row in Get-UcSqlQuery $AxlConn $sql) {
    $row.PSObject.TypeNames.Insert(0,'UcUserAcl')
    set-cache enduserpkid $row.user $row._userpkid
    $row
  }
}

function Add-UcUserAcl {
<#
  .synopsis
  Add an entry to the user's access-list (allow/block list)
  
  .parameter User
  User who's ACL is to be modified.  Either this or the pkid parameter may be specified
  If both are supplied, then pkid is used.

  .parameter Uri
  The Uri to be blocked or allowed.  The sip: prefix for individual (non-domain) addresses
  is allowed, but not required.  Do not use the prefix for domain Uri's.
  
  .parameter Policy
  Either 'allowed' or 'politeblocking'
  
  .parameter pkid
  The user's pkid from the enduser table in the CUP database.  Either this or the User parameter
  may be specified.  If both are supplied, then this parameter will be used.
#>
  Param(
    [Parameter(Mandatory=$true)]
    $AxlConn
    ,
    [Parameter(ValueFromPipelineByPropertyName=$true)]
    [string]
    [ValidateLength(1,132)]
    [ValidateScript({$_ -notmatch "[ ']"})]
    $User
    ,
    [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
    [string]
    [ValidateLength(1,128)]
    [ValidateScript({$_ -notmatch "[ ']"})]
    $Uri
    ,
    [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
    [string]
    [ValidateRange("allowed", "politeblocking")]
    $Policy
    ,
    [Parameter(ValueFromPipelineByPropertyName=$true)]
    [alias("_userpkid")]
    $pkid
  )
  if (-not $pkid -and -not $user) {
    throw "One of either pkid or User parameters must be specified"
  }
  if (-not $pkid) {
    #pkid was not given, so we need to fetch it
    if (-not ($pkid = get-cache enduserpkid $user)) {
      $sql = "select pkid from enduser where userid = '${User}'"
      if (-not ($pkid = get-ucsqlquery $AxlConn $sql | select -expand pkid)) {
        throw "Unable to find pkid for $User.  Is this a valid CUP user"
      }
      set-cache enduserpkid $user $pkid
    }
  }
  #prepend sip: prefix if necessary
  $urifixed = $uri -replace "(?!sip:)(.+@)",'sip:$1'
  $sql = @"
execute procedure ucsoapadduriacl(
  '${pkid}', 
  '${Policy}', 
  '${urifixed}'
)
"@
  foreach ($row in Get-UcSqlQuery $AxlConn $sql) {
    if ($row.success -ne "t") {
      $msg = "Parameters: User=$User, Pkid=$pkid, Policy=$Policy, Uri=$Uri. "+
        "Add ACL operation failed. Server returned Success=$($row.success), Reason=$($row.reason)"
      throw $msg
    }
  }
}

function Get-UcLineAppearance {
<#
  .synopsis
  Get CallManager line appearances

  .parameter DN
  Only get line appearances for the specified DN.  SQL wildcard character '%' is allowed.  The default is '%' which
  gets line appearances for all DNs
  
  .parameter Limit
  Limit search results to the given number of records.

  .parameter axlconn
  Connection object created with New-AxlConnection.
#>
  Param(
    [Parameter(Mandatory=$true)][psobject]$AxlConn,
    [string]$DN="%"
    ,
    [ValidateScript({[int]$_ -gt 0})]
    $Limit
  )
  $count = 0
  $defaultlimit=100
  if ($limit -eq $null) { $effectivelimit = $defaultlimit }
  else { $effectivelimit = [int]$limit }
  $sql = @"
SELECT first ${effectivelimit}
 np.dnorpattern dn,
 rp.name routePartition,
 dmp.display callerid,
 dmp.label LineLabel,
 dmp.e164mask mask,
 d.name device,
 d.description devicedescription,
 m.name model
FROM
 numplan np
  join routepartition rp on np.fkroutepartition=rp.pkid
  right outer join devicenumplanmap dmp on np.pkid=dmp.fknumplan,
 device d
  join typemodel m on d.tkmodel=m.enum
WHERE
 d.pkid=dmp.fkdevice
 and np.dnorpattern like '${DN}'
"@
  foreach ($row in Get-UcSqlQuery $AxlConn $sql) {
    $row.PSObject.TypeNames.Insert(0,'UcLineAppearance')
    $row
    $count ++
  }
  if (-not $limit -and $count -ge $defaultlimit) {
    write-warning "$($pscmdlet.myinvocation.invocationname): Results were limited to the first $count records.  Use -Limit to return more records"
  }
}

<#
  .synopsis
  Get a list of users which have the specified buddy address in their contact list

  .parameter axlconn
  Connection object created with New-AxlConnection.

  .parameter Buddy
  SIP address of the buddy

  .parameter Force
  Include buddies which are in state 2, meaning the owner has authorized the buddy to see
  their presence, but the buddy has not been added to any group.
  
  .outputs
  [string] A list of owner (userid's) who have the given buddy in their contact list
#>
function Get-UcWatcher {
  [CmdletBinding()]
  Param(
    [Parameter(Mandatory=$true, Position=0)]
    [psobject]$AxlConn
    ,
    [Parameter(Mandatory=$true, Position=1)]
    [ValidatePattern("^(?!sip:)\S+@\S+\.\S+")]
    [string]$Buddy
    ,
    [switch]$Force
  )
  $owners = @()
  $sql = @"
select 
 u.userid Owner,
 r.contact_jid Buddy,
 r.state State,
 g.group_name Group
from 
 rosters r 
 join enduser u on u.xcp_user_id = r.user_id 
 left outer join groups g on g.roster_id = r.roster_id
where
 lower(r.contact_jid) = lower('${Buddy}')
order by 
 u.userid
"@
  Get-UcSqlQuery $AxlConn $sql | % {
    #$_.PSObject.TypeNames.Insert(0,'UcBuddy')
    if ($force -or $_.group -ne "") {
      #$_
      $_.owner
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


Export-ModuleMember -Function Get-UcLicenseCapabilities, New-AxlConnection, `
  Get-UcSqlQuery, Get-UcSqlUpdate, Get-UcBuddy, Remove-UcBuddy, Add-UcBuddy, `
  Get-UcUser, Get-UcDialPlans, Get-UcDN, Get-UcLineAppearance, Rename-UcBuddy, `
  Get-UcWatcher, get-table, Get-StoredProc, Get-UcUserAcl, Get-UcCupUsers, `
  Add-UcUserAcl
