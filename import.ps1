$config_file = "~\.ssh\config"
$winscp = "winscp\winscp.com"
$patterns = @(
    "Host (?<host>[A-Za-z-0-9\.]+)\s*"
    "HostName (?<hostname>[A-Za-z-0-9.]+)\s*"
    "Port (?<port>\d+)\s*"
    "User (?<user>\w+)\s*"
    "IdentityFile (?<key>[A-Za-z~0-9/._-]+)\s*"
)

function PathEncoding ($path) {
    return $(Split-Path -Path $path -Qualifier) + $([uri]::EscapeDataString($(Split-Path -Path $path -NoQualifier)))
}

function permutation ($array) {
    function generate($n, $array, $A) {
        if($n -eq 1) {
            $array[$A] -join ''
        }
        else{
            for( $i = 0; $i -lt ($n - 1); $i += 1) {
                generate ($n - 1) $array $A
                if($n % 2 -eq 0){
                    $i1, $i2 = $i, ($n-1)
                    $A[$i1], $A[$i2] = $A[$i2], $A[$i1]
                }
                else{
                    $i1, $i2 = 0, ($n-1)
                    $A[$i1], $A[$i2] = $A[$i2], $A[$i1]
                }
            }
            generate ($n - 1) $array $A
        }
    }
    $n = $array.Count
    if($n -gt 0) {
        (generate $n $array (0..($n-1)))
    } else {$array}
}

function Get-AllPatterns ($patterns) {
    $patternsGen = permutation $patterns[1..($patterns.Length-1)]
    $out =@()
    foreach ($item in $patternsGen) {
        $out += $patterns[0]+$item
    }
    $out
}


function Get-Data {
    $data_list = @()
    $all_patterns = Get-AllPatterns $patterns
    $data = Get-Content $config_file -Raw
    foreach ($pattern in $all_patterns) {
        $data | Select-String -Pattern $pattern -AllMatches |
        Foreach-Object {
            foreach ($item in $_.Matches) {
                $name, $hostname, $port, $user, $key = $item.Groups['host', 'hostname', 'port', 'user', 'key'].Value
                $data_list += [PSCustomObject] @{
                                    Host = $name
                                    HostName = $hostname
                                    Port = $port
                                    User = $user
                                    Key = $key
                                }
            }
        }
    }
    $data_list 
}

#Get-Data | Format-Table

function Set-XML ($name, $hostname, $port, $user, $key) {
    # Set The Formatting
    $xmlsettings = New-Object System.Xml.XmlWriterSettings
    $xmlsettings.Indent = $true
    $xmlsettings.IndentChars = "    "

    # Set the File Name Create The Document
    $XmlWriter = [System.XML.XmlWriter]::Create("${name}.netbox", $xmlsettings)

    # Write the XML Decleration and set the XSL
    $xmlWriter.WriteStartDocument()

    # Start the Root Element
    $xmlWriter.WriteStartElement("NetBox")
    $XmlWriter.WriteStartAttribute("version")
    $XmlWriter.WriteValue("2.1")
    
        $xmlWriter.WriteStartElement("Sessions") # <-- Start <Object>
            $xmlWriter.WriteStartElement("Session") # <-- Start <SubObject> 
            $XmlWriter.WriteStartAttribute("name")
            $XmlWriter.WriteValue("$name")
                $xmlWriter.WriteElementString("Version","2.4.5")
                $xmlWriter.WriteElementString("HostName","$hostname")
                $xmlWriter.WriteElementString("PortNumber","$port")
                $xmlWriter.WriteElementString("UserName","$user")
                $xmlWriter.WriteElementString("PublicKeyFile","$key")
                $xmlWriter.WriteElementString("FSProtocol","SCP")
                $xmlWriter.WriteElementString("RemoteDirectory","/home/${user}")
                $xmlWriter.WriteElementString("SFTPMaxVersion","0")
                $xmlWriter.WriteElementString("TlsCertificateFile","$key")
            $xmlWriter.WriteEndElement() # <-- End <SubObject>
        $xmlWriter.WriteEndElement() # <-- End <Object>
    $xmlWriter.WriteEndElement() # <-- End <Root> 

    # End, Finalize and close the XML Document
    $xmlWriter.WriteEndDocument()
    $xmlWriter.Flush()
    $xmlWriter.Close()
}

function PreparePath ($path) {
    $file = Get-Item $path
    $file.DirectoryName + "\" + $file.BaseName + ".ppk" 
}

function ConvertKey ($path) {
    Push-Location $PSScriptRoot
    $out_path = PreparePath($path)
    if (Test-Path $out_path) {
        Write-Host "Key file exists..."
    }
    else {
        $full_path = Get-Item $path | Select-Object -ExpandProperty FullName
        & $winscp /keygen $full_path /output=$out_path
    }
    Pop-Location
}

$data = Get-Data
foreach ($item in $data) {
    ConvertKey($item.Key)
    $key_path = PathEncoding(PreparePath($item.Key))
    Set-XML -name $item.Host -hostname $item.HostName -port $item.Port -user $item.User -key $key_path
}
