# Introduction #

Usage examples



## Delete a bunch of phones ##

```
$axl = new-axlconnection cucm1.example.com admin
$phones = "SEP9CAFCAFEF6F2","SEP001E4AF18FA1","SEP9CAFCAFEF4E8"
$phones | % { $_; try {Remove-UcPhone $axl -name $_} catch { write-warning $_.Exception.Message } }
```

## Delete a list of DNs and any phone the DN is associated with ##
```
$axl = new-axlconnection cucm1.example.com admin
$lines = "2001","2002","2003","2004"
$partition = "phones-pt"
foreach ($line in $lines) {
  write-host -for cyan $line
  $appearances = get-uclineappearance $axl $line
  $appearances | %{
    $_
    try {
      $_ | remove-ucphone $axl
    } catch {
      write-warning $_.Exception.Message
    }
  }
  try {
    remove-ucdn $axl -dn $line -part $partition
  } catch {
    write-warning "Could not delete DN $line.  $($_.Exception.Message)"
  }
}
```

## Change the route partition of some Lines (DNs) ##
```
$axl = new-axlconnection cucm1.example.com admin
$lines = get-ucdn $axl -dn % -Partition "old-partition"
foreach ($line in $lines) {
  "{0} {1}" -f $line.dn, $line.description
  try { set-ucline $axl -routep "old-partition" -newRoutePartitionName "new-partition" -pat $line.dn }
  catch { write-warning $_.Exception.Message }
}
```