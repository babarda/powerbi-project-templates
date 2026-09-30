# Developer helper: full refresh (or a DAX query) against the model of a running Power BI Desktop instance.
# Usage: .\refresh_model.ps1 -DesktopPid 1234            (refresh)
#        .\refresh_model.ps1 -DesktopPid 1234 -Dax "EVALUATE ..."
# Needs the SqlServer PowerShell module (Invoke-ASCmd). Author: Badr Abardazzou. License: MIT.
param([int]$DesktopPid, [string]$Dax = "")
Import-Module SqlServer -ErrorAction Stop
$ms = Get-CimInstance Win32_Process -Filter "Name='msmdsrv.exe'" | Where-Object ParentProcessId -eq $DesktopPid
if (-not $ms) { throw "No msmdsrv for $DesktopPid" }
$port = (Get-NetTCPConnection -State Listen | Where-Object OwningProcess -eq $ms.ProcessId | Select-Object -First 1).LocalPort
$server = "localhost:$port"
$dbxml = Invoke-ASCmd -Server $server -Query "<Discover xmlns='urn:schemas-microsoft-com:xml-analysis'><RequestType>DBSCHEMA_CATALOGS</RequestType><Restrictions/><Properties/></Discover>"
$db = ([regex]::Match($dbxml, "<CATALOG_NAME>([^<]+)</CATALOG_NAME>")).Groups[1].Value
if ($Dax) {
  $r = Invoke-ASCmd -Server $server -Database $db -Query "<Execute xmlns='urn:schemas-microsoft-com:xml-analysis'><Command><Statement>$([Security.SecurityElement]::Escape($Dax))</Statement></Command><Properties><PropertyList><Catalog>$db</Catalog><Format>Tabular</Format></PropertyList></Properties></Execute>"
  $x = [xml]$r
  $rows = $x.return.root.row
  if ($rows) { $rows | ForEach-Object { ($_.ChildNodes | ForEach-Object { "$($_.LocalName)=$($_.InnerText)" }) -join ' | ' } } else { $r }
} else {
  "server=$server db=$db"
  $tmsl = '{"refresh":{"type":"full","objects":[{"database":"' + $db + '"}]}}'
  Invoke-ASCmd -Server $server -Database $db -Query $tmsl
}
