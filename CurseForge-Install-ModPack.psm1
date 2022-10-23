using namespace System.IO.Compression



$global:CurseApiBaseUri = "https://api.curseforge.com"

$global:CurseApiTimeoutSec = 30

$script:minecraft = '.\minecraft'
$script:temp = '.\extracted'

# $minecraftCurseGameId = ((Invoke-RestMethod ($CurseApiBaseUri + '/v1/games') -Headers $headers).data | where-object slug -eq 'minecraft').id
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

    if ($null -eq $env:CurseApiKey -or '' -eq $env:CurseApiKey) {
        throw "CurseApiKey environment variable not set."
    }

    $req = @{
        Uri = $CurseApiBaseUri.TrimEnd('/') +'/'+ $Uri.TrimStart('/')
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
        $res = Invoke-RestMethod @req -UseBasicParsing -TimeoutSec $CurseApiTimeoutSec
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
            $res = Invoke-CurseApi -Uri "/v1/mods/$ProjectId/files/$FileId"
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

        $existingFile = Get-Item $OutFile -ErrorAction SilentlyContinue
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
        [string]$FileName
    )

    New-Item $temp -Type Directory -ErrorAction SilentlyContinue | Out-Null
    Expand-Archive $FileName $temp -Force
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
    return Get-Content .\extracted\manifest.json | ConvertFrom-Json
}

function Get-ModPackFiles {
    param (
        [Parameter(Mandatory)]
        [PSCustomObject]$ModPackManifest
    )

    $req = @{
        Uri = '/v1/mods/files'
        Method = 'POST'
        Body = @{
            fileIds = $ModPackManifest.files | Select-Object -ExpandProperty fileID
        } | ConvertTo-Json
    }
    $res = Invoke-CurseApi @req
    if ($res -is [Exception]) {
        Write-Error "Failed to get modpack files" -Exception $res
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
        [object[]]$Files,
        # Force download of all files even if they already exist and the length matches
        [Parameter()]
        [switch]$Force = $false
    )
    $mods = Join-Path $minecraft 'mods'
    New-Item $mods -Type Directory -ErrorAction SilentlyContinue | Out-Null
    $Files `
        | Where-Object {
            $Force `
            -or $nul -eq ($f = Get-Item (Join-Path $mods $_.fileName) -ErrorAction SilentlyContinue) `
            -or $f.Length -ne $_.fileLength
        } `
        | ForEach-Object { @{
            Uri = $_.downloadUrl
            OutFile = (Join-Path $mods $_.fileName)
        } } `
        | Invoke-WebDownloadAll `
        | Where-Object Response -is [Exception] `
        | ForEach-Object { Write-Error -Message "Failed to download ""$($_.Request.Uri)""" -Exception $_.Response }
}

function Invoke-ModPackOverrides {
    param (
        # Path to the source folder containing the files to be overridden
        [Parameter(Mandatory)]
        [string]$Overrides
    )
    $Overrides = (Join-Path $Overrides '*')
    Write-Verbose "Copying overrides from ""$Overrides"" to ""$minecraft"""
    Copy-Item -Path $Overrides -Destination $minecraft -Recurse -Force
}

function Invoke-ModPackCleanup {
    Remove-Item -Path .\extracted -Recurse
}

function Write-ModPackInstructions {}

function Get-ModPack {
    param (
        [Parameter(Mandatory)]
        [int]$ProjectId,
        [int]$FileId
    )

    Write-Information "Downloading modpack ""$($ProjectId)"""
    $latestFile = Invoke-ModPackDownload -ProjectId $ProjectId -FileId $FileId

    Write-Information "Extracting modpack file ""$($latestFile.displayName)"""
    Invoke-ModPackExtract $latestFile.fileName

    Write-Information "Loading modpack manifest"
    $manifest = Get-ModPackManifest

    Write-Information "Downloading modpack dependencies"
    Invoke-ModPackFilesDownload (Get-ModPackFiles $manifest)

    Write-Information "Copying modpack overrides"
    Invoke-ModPackOverrides (Join-Path '.\extracted' $manifest.overrides)

    Write-Information "Cleaning up"
    Invoke-ModPackCleanup

    Write-ModPackInstructions $manifest
}

# Examples:
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


#$res = Invoke-CurseApi ('/v1/mods/' + $projectId + '/files')
#return $res


# $ProgressPreference = 'SilentlyContinue'
# $ProgressPreference = 'Continue'

# # $projectId = 638321 # Feed the Factory

# # $res = Invoke-RestMethod ($CurseApiBaseUri + '/v1/mods/' + $projectId) -Headers $headers

# # $modFiles = (Invoke-RestMethod ($CurseApiBaseUri + '/v1/mods/' + $projectId + '/files') -Headers $headers).data
# # $latestFile = $modFiles | Sort-Object fileDate -Descending | Select-Object -First 1
# # $req = @{
# #     Uri = $latestFile.downloadUrl
# #     OutFile = Join-Path $modpacks $( $latestFile.fileName ).downloading
# #     UseBasicParsing = $true
# # }
# # Invoke-WebRequest $req

# # New-Item $modpacks\$( [IO.Path]::GetFileNameWithoutExtension($latestFile.fileName) ) -ItemType Directory -ErrorAction SilentlyContinue



# # $uri



# # $curseApiProjectSearchUrl = "/addon/search"

# # $searchFilter = "Feed the Factory"

# # $query = @{
# #     gameId = 432;
# #     categoryId = 0;
# #     sectionId = 4471;
# #     searchFilter = $searchFilter;
# #     pageSize = 20;
# #     index = 0;
# #     sort = 1;
# #     sortDescending = $true;
# # }

# # $queryString = (( $query.GetEnumerator() | foreach-object { "$($_.Name)=$($_.Value)" } ) -join '&')

# # Invoke-RestMethod ($CurseApiBaseUri + $curseApiProjectSearchUrl + '?' + $queryString) -Headers @{"x-api-key" = $env:CurseApiKey}


