using namespace System.IO.Compression

Add-Type -AssemblyName 'System.IO.Compression'
Add-Type -AssemblyName 'System.IO.Compression.FileSystem'

$InformationPreference = 'Continue'

if (-not (test-path variable:IsWindows)) {
    [Diagnostics.CodeAnalysis.SuppressMessage('PSAvoidAssignmentToAutomaticVariable', $null, Justification = 'because polyfilling')]
    $IsWindows = 'Windows_NT' -eq $env:OS
}

$global:CurseApiBaseUri = 'https://api.curseforge.com/'

$global:CurseApiTimeoutSec = 30

$script:minecraft = '.\minecraft'

function Invoke-CurseApi {
    [CmdletBinding()]
    param (
        # The URI to request
        [Parameter(Mandatory)]
        [string]$Uri,
        # Optional HTTP method to request
        [Parameter()]
        [string]$Method = 'GET',
        # Optional
        [Parameter()]
        [Object]$Body,
        # Optional file path to write content
        [Parameter()]
        [string]$OutFile
    )

    if ([string]::IsNullOrEmpty($env:CurseApiKey)) {
        throw 'CurseApiKey environment variable not set.'
    }

    [uri]$baseUri = $null
    $isValid = [uri]::TryCreate($CurseApiBaseUri, [UriKind]::Absolute, ([ref]$baseUri))
    if (-not $isValid) {
        throw [ArgumentException]::new('$CurseApiBaseUri variable is not a valid absolute URI.', '$global:CurseApiBaseUri')
    }

    [uri]$fullUri = $null
    $isValid = [uri]::TryCreate($baseUri, $Uri, ([ref]$fullUri))
    if (-not $isValid) {
        throw [ArgumentException]::new('Uri parameter is not a valid relative URI.', 'Uri')
    }

    $req = @{
        UseBasicParsing = $true
        TimeoutSec = $CurseApiTimeoutSec
        Uri = $fullUri
        Headers = @{
            'x-api-key' = $env:CurseApiKey
        }
        ContentType = 'application/json'
        Method = $Method
        Body = $Body
        OutFile = $OutFile
    }

    Write-Verbose "Curse API requesting ""$Uri"""

    $currentProgPref = $ProgressPreference
    $ProgressPreference = 'SilentlyContinue'

    try {
        $res = Invoke-RestMethod @req
        $ProgressPreference = $currentProgPref
        Write-Verbose "Curse API completed request ""$Uri"""
        return $res
    }
    catch {
        $ProgressPreference = $currentProgPref
        Write-Verbose "Curse API failed request ""$Uri"" with error: $($_.Exception.Message)"
        return $_.Exception
    }
}

function Invoke-WebDownload {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Uri,
        [Parameter(Mandatory)]
        [string]$OutFile
    )

    # HACK: Saving to temp file first to workaround Invoke-RestMethod OutFile wildcard issue
    # See: https://stackoverflow.com/questions/55869623

    $req = @{
        Uri = $Uri
        OutFile = (New-TemporaryFile)
    }

    Write-Verbose "Downloading ""$Uri"" to ""$OutFile"""

    $currentProgPref = $ProgressPreference
    $ProgressPreference = 'SilentlyContinue'

    try {
        $res = Invoke-WebRequest @req -UseBasicParsing
        $ProgressPreference = $currentProgPref
        Move-Item -Force -LiteralPath $req.OutFile -Destination $OutFile
        Write-Verbose "Download completed ""$Uri"""
        return $res
    }
    catch {
        $ProgressPreference = $currentProgPref
        Write-Verbose "Download failed ""$Uri"" with error: $($_.Exception.Message)"
        return $_.Exception
    }
}

function Invoke-WebDownloadAll {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline)]
        [PSCustomObject]$Request,
        [Parameter()]
        [int]$ThrottleLimit = 5
    )

    begin {
        [Collections.ArrayList]$inputObjects = @()
    }
    process {
        if ($null -eq $Request) { return }
        [void]$inputObjects.Add($Request)
        Write-Verbose "Added ""$($Request.Uri)"" to download queue"
    }
    end {
        [Collections.ArrayList]$outputObjects = @()
        Write-Verbose "Downloading $($inputObjects.Count) queued URIs"
        if ((Get-Command ForEach-Object).Parameters.ContainsKey("Parallel")) {
            $inputObjects | ForEach-Object -ThrottleLimit $ThrottleLimit -Parallel {
                $outputs = $using:outputObjects
                [void]$outputs.Add(@{
                    Request = $PSItem
                    Response = (Invoke-WebDownload @PSItem)
                })
            }
        }
        else {
            $inputObjects | ForEach-Object {
                [void]$outputObjects.Add(@{
                    Request = $PSItem
                    Response = (Invoke-WebDownload @PSItem)
                })
            }
        }
        Write-Verbose "Finished downloading all $($inputObjects.Count) URIs"
        return $outputObjects
    }
}


function Invoke-ModPackDownload {
    [CmdletBinding()]
    param (
        # CurseForge Project ID of the modpack to download
        [Parameter(Mandatory)]
        [int]$ProjectId,
        # CurseForge File ID of the modpack file to download
        [Parameter()]
        [int]$FileId,
        # File location to save the download
        [Parameter()]
        [string]$OutFile
    )

    end {
        $latestFile = $null
        if (0 -eq $FileId) {
            $res = Invoke-CurseApi -Uri "/v1/mods/$ProjectId/files"
            if (-not ($res -is [Exception])) {
                $latestFile = $res.data | Sort-Object fileDate -Descending | Select-Object -First 1
            }
        }
        else {
            $res = Invoke-CurseApi -Uri "/v1/mods/$ProjectId/files/$FileId" -ErrorVariable 'res'
            if (-not ($res -is [Exception])) {
                $latestFile = $res.data
            }
        }

        if ($res -is [Exception] -or $null -eq $latestFile) {
            $err = @{
                Message = "File was not found."
            }
            if ($res -is [Exception]) {
                $err.Exception = $res
            }
            Write-Error @err
            return
        }
        Write-Verbose "Found modpack file ""$($latestFile.fileName)"""

        if ($OutFile -eq '') {
            $OutFile = $latestFile.fileName
        }
        Write-Verbose "Saving file to ""$OutFile"""

        $existingFile = Get-Item -LiteralPath $OutFile -ErrorAction SilentlyContinue
        if ($? -and $existingFile.Length -eq $latestFile.fileLength) {
            Write-Information "Modpack file ""$OutFile"" already exists"
            return $latestFile
        }

        Write-Information "Downloading modpack file ""$OutFile"""
        Invoke-WebDownload -Uri $latestFile.downloadUrl -OutFile $OutFile

        return $latestFile
    }
}

function Invoke-ModPackExtract {
    param (
        # Path to the modpack zip file to extract
        [Parameter(Mandatory)]
        [string]$FileName,
        # Path to save the files extracted from the modpack
        [string]$Destination
    )

    New-Item $Destination -Type Directory -ErrorAction SilentlyContinue | Out-Null
    Expand-Archive $FileName $Destination -Force
}

function Get-ZipContent {
    param (
        # Path to the zip file to extract
        [Parameter(Mandatory)]
        [string]$Archive,
        # Name of the file to read the content
        [Parameter()]
        [string]$FileName
    )

    [ZipArchive]$zip = [ZipFile]::OpenRead((Convert-Path $Archive))
    try
    {
        $entry = $zip.Entries | Where-Object FullName -eq $FileName | Select-Object -First 1
        if ($null -eq $entry) {
            throw "Entry not found in zip file."
        }

        $stream = New-Object IO.StreamReader $entry.Open()
        try {
            return $stream.ReadToEnd()
        }
        finally {
            $stream.Dispose()
        }
    }
    finally {
        $zip.Dispose()
    }
}

function Get-ModPackManifest {
    param (
        [Parameter(Mandatory)]
        [string]$ZipFile
    )
    return Get-ZipContent $ZipFile manifest.json | ConvertFrom-Json
}

function Get-ModPackFiles {
    param (
        [Parameter(Mandatory)]
        [PSCustomObject]$ModPackManifest
    )

    if ($ModPackManifest.files.Length -eq 0) {
        return @()
    }

    $req = @{
        Uri = '/v1/mods/files'
        Method = 'POST'
        Body = @{
            fileIds = $ModPackManifest.files | Select-Object -ExpandProperty fileID
        } | ConvertTo-Json
    }
    $res = Invoke-CurseApi @req
    if ($res -is [Exception]) {
        Write-Error -Message "Failed to get modpack files" -Exception $res
        return
    }

    # dedup because Curse API sometimes likes to repeat itself
    $files = $res.data | Group-Object id | ForEach-Object { $_.Group[0] }
    return $files
}

function Invoke-ModPackFilesDownload {
    param (
        # List of modpack files to download
        [Parameter(Mandatory)]
        [PSCustomObject[]]$Files,
        # Force download of all files even if they already exist and the length matches
        [Parameter()]
        [switch]$Force = $false
    )
    $mods = Join-Path $minecraft 'mods'
    New-Item $mods -Type Directory -ErrorAction SilentlyContinue | Out-Null
    $Files `
        | Where-Object {
            $Force `
            -or $nul -eq ($f = Get-Item -LiteralPath (Join-Path $mods $_.fileName) -ErrorAction SilentlyContinue) `
            -or $f.Length -ne $_.fileLength
        } `
        | ForEach-Object { @{
            Uri = (Get-CurseFileDownloadUrl $_)
            OutFile = (Join-Path $mods $_.fileName)
        } } `
        | Invoke-WebDownloadAll `
        | Where-Object Response -is [Exception] `
        | ForEach-Object { Write-Error -Message "Failed to download ""$($_.Request.Uri)""" -Exception $_.Response }
}

function Get-CurseFileDownloadUrl {
    param (
        [PSCustomObject]$File
    )
    if ($null -ne $File.downloadUrl) {
        return $File.downloadUrl
    }

    # Try to guess the Curse CDN uri for this file
    $id1 = [Math]::Truncate($File.id / 1000)
    $id2 = $File.id % 1000
    return "https://mediafilez.forgecdn.net/files/$id1/$id2/$([uri]::EscapeDataString($File.fileName))"
}

function Invoke-ModPackOverrides {
    param (
        # Path to the modpack zip file
        [Parameter(Mandatory)]
        [string]$ModPack,
        # Path to the source folder containing the files to be overridden
        [Parameter(Mandatory)]
        [string]$Overrides
    )

    if ($null -eq (Get-Command tar -ErrorAction SilentlyContinue)) {
        throw "Command tar is not available."
    }

    Write-Verbose "Copying overrides from ""$(Join-Path $ModPack $Overrides)"" to ""$minecraft"""
    $components = $Overrides.Trim('\', '/').Split(@('\', '/')).Count
    tar -x -f $ModPack --strip-components $components -C $minecraft $Overrides
}

function Get-ManifestMCVersion {
    param (
        [PSCustomObject]$Manifest
    )
    
    $version = @{
        MinecraftVersion = $Manifest.minecraft.version
        RawForgeVersion = $Manifest.minecraft.modLoaders | Where-Object primary | Select-Object -ExpandProperty id
    }
    $version.ForgeVersion = $version.RawForgeVersion -split '-' | Select-Object -Last 1
    
    return $version
}

function Install-ForgeServer {
    [CmdletBinding(DefaultParameterSetName = 'Manifest')]
    param (
        [Parameter(Mandatory, Position = 0, ParameterSetName = 'Manifest')]
        [PSCustomObject]$Manifest,
        [Parameter(Mandatory, ParameterSetName = 'Version')]
        [string]$MinecraftVersion,
        [Parameter(Mandatory, ParameterSetName = 'Version')]
        [string]$ForgeVersion
    )

    if ($PSCmdlet.ParameterSetName -eq 'Manifest') {
        $version = Get-ManifestMCVersion $Manifest
        $MinecraftVersion = $version.MinecraftVersion
        $ForgeVersion = $version.ForgeVersion
    }

    $fullVersion = "$MinecraftVersion-$ForgeVersion"
    $serverJar = "forge-$fullVersion.jar"

    New-Item $minecraft -Type Directory -ErrorAction SilentlyContinue | Out-Null

    Push-Location $minecraft
    try {

        if (Test-Path $serverJar -ErrorAction Ignore) {
            Write-Verbose "Forge server is already installed"
            Invoke-CreateServerStartScript $serverJar
            return;
        }

        $installerJar = "forge-$fullVersion-installer.jar"
        $installerUri = "https://maven.minecraftforge.net/net/minecraftforge/forge/$fullVersion/$installerJar"

        Write-Verbose "Downloading Forge installer $fullVersion"
        $res = Invoke-WebDownload -Uri $installerUri -OutFile $installerJar
        if ($res -is [Exception]) {
            throw New-Object -TypeName Exception -ArgumentList "Failed to download Forge installer $fullVersion`: $($res.Message)", $res
        }

        Write-Verbose "Installing Forge server $fullVersion"
        java -jar $installerJar --installServer
        if (-not $?) {
            throw New-Object -TypeName Exception -ArgumentList "Failed to install Forge server"
        }
        Remove-Item $installerJar

        Invoke-CreateServerStartScript $serverJar

        Write-Verbose "Installed Forge $fullVersion"

    }
    finally {
        Pop-Location
    }
}

function Invoke-CreateServerStartScript {
    param (
        [Parameter(Mandatory)]
        [string]$ServerJar
    )

    Write-Verbose "Creating Start-Server script"
    $maxRam = "5G"
    $javaArgs = @(
        "-Xmx$maxRam",
        #"-Xms$maxRam",
        "-version:1.8+",
        "-XX:+UseG1GC",
        "-XX:+ParallelRefProcEnabled",
        "-XX:MaxGCPauseMillis=200",
        "-XX:+UnlockExperimentalVMOptions",
        "-XX:+DisableExplicitGC",
        "-XX:+AlwaysPreTouch",
        "-XX:G1NewSizePercent=30",
        "-XX:G1MaxNewSizePercent=40",
        "-XX:G1HeapRegionSize=8M",
        "-XX:G1ReservePercent=20",
        "-XX:G1HeapWastePercent=5",
        "-XX:G1MixedGCCountTarget=4",
        "-XX:InitiatingHeapOccupancyPercent=15",
        "-XX:G1MixedGCLiveThresholdPercent=90",
        "-XX:G1RSetUpdatingPauseTimePercent=5",
        "-XX:SurvivorRatio=32",
        "-XX:+PerfDisableSharedMem",
        "-XX:MaxTenuringThreshold=1",
        "-Dusing.aikars.flags=https://mcflags.emc.gs",
        "-Daikars.new.flags=true",
        "-Dfml.readTimeout=90", # servertimeout
        "-Dfml.queryResult=confirm", # auto /fmlconfirm
        "--add-opens=java.base/sun.security.util=ALL-UNNAMED", # java16+ support
        "--add-opens=java.base/java.util.jar=ALL-UNNAMED", # java16+ support
        "-XX:+IgnoreUnrecognizedVMOptions" # java16+ support
    )
    # Set-Content -Path java-args.txt -Value $javaArgs

    if ($IsWindows) {
        $script = 'start-server.bat'
        Set-Content -Path $script -Value @(
            "@echo off"
            'cd /d "%~dp0"'
            #"java -jar $ServerJar nogui @java-args.txt"
            "java $($javaArgs -join ' ') -jar $ServerJar nogui"
            "pause"
        )
    }
    else {
        $script = 'start-server.sh'
        Set-Content -Path $script -Value @(
            "#!/bin/bash"
            #"java -jar $ServerJar nogui @java-args.txt"
            "java $($javaArgs -join ' ') -jar $ServerJar nogui"
            "pause"
        )
    }
}

function Write-ModPackInstructions {
    param (
        [PSCustomObject]$Manifest,
        [switch]$Server
    )

    if ($Server) {
        ""
        "Run ""$(Join-Path $minecraft 'start-server')"" to start the server."
        ""
        return
    }

    $version = Get-ManifestMCVersion $Manifest

    ""
    "Installation:"
    ""
    "1. Download and run the Minecraft Forge installer version $($version.MinecraftVersion) - $($version.ForgeVersion)"
    "from https://files.minecraftforge.net/net/minecraftforge/forge/index_$($version.MinecraftVersion).html"
    ""
    "2. In the Minecraft Launcher create a new instance using this Forge version"
    "and the ""$(Convert-Path $minecraft)"" folder."
    ""
}

function Install-ModPack {
    param (
        [Parameter(Mandatory)]
        [int]$ProjectId,
        [int]$FileId,
        [switch]$Server
    )

    Write-Information "Downloading modpack ""$($ProjectId)"""
    $modpack = Invoke-ModPackDownload -ProjectId $ProjectId -FileId $FileId

    Write-Information "Loading modpack manifest"
    $manifest = Get-ModPackManifest $modpack.fileName

    if ($Server) {
        Write-Information "Installing Forge Server"
        Install-ForgeServer -Manifest $manifest
    }

    Write-Information "Downloading modpack dependencies"
    Invoke-ModPackFilesDownload (Get-ModPackFiles $manifest)

    Write-Information "Copying modpack overrides"
    Invoke-ModPackOverrides $modpack.fileName $manifest.overrides

    Write-ModPackInstructions $manifest -Server:$Server
}

# Examples:
#
# Install-ModPack -ProjectId 638321 # Feed the Factory
# Install-ModPack -ProjectId 282744 # Enigmatica2Expert
# 
# @(
#     @{ Uri="https://google.com"; OutFile="google.txt" },
#     @{ Uri="https://google.com"; OutFile="google2.txt" }
# ) | Invoke-WebDownloadAll
#
# cd temp
# $latestFile = Invoke-ModPackDownload -ProjectId 638321 # Feed the Factory
# Invoke-ModPackExtract $latestFile.fileName
# Get-ModPackManifest
# Invoke-ModPackFilesDownload (Get-ModPackFiles (Get-ModPackManifest))
# $res = Invoke-CurseApi ('/v1/mods/' + $projectId + '/files')
# $minecraftCurseGameId = ((Invoke-RestMethod ($CurseApiBaseUri + '/v1/games') -Headers $headers).data | where-object slug -eq 'minecraft').id
#
# Invoke-RestMethod ($CurseApiBaseUri + '/addon/search') -Body @{
#     gameId = 432
#     categoryId = 0
#     sectionId = 4471
#     searchFilter = "Feed the Factory"
#     pageSize = 20
#     index = 0
#     sort = 1
#     sortDescending = $true
# } -Headers @{"x-api-key" = $env:CurseApiKey}
#
# (Invoke-RestMethod ($CurseApiBaseUri + '/v1/mods/search') -Body @{
#     'gameId' = 432
#     'classId' = 4471
#     'searchFilter' = $null
#     'slug' = 'feed-the-factory'
#     'pageSize' = 20
#     'index' = 0
#     'sortField' = 3
#     'sortOrder' = $true
# } -Headers @{"x-api-key" = $env:CurseApiKey} `
# ).data | format-table