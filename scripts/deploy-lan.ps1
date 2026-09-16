[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$RemoteHost,

    [string]$RemoteUser = 'jl',

    [string]$ImageTag = 'caiwu1:lan',

    [string]$ContainerName = 'caiwu1',

    [int]$ServicePort = 15618,

    [int]$DevPort = 59323
)

$ErrorActionPreference = 'Stop'

function Require-Command([string]$Name) {
    if (-not (Get-Command $Name -ErrorAction SilentlyContinue)) {
        throw "Required command not found: $Name"
    }
}

Require-Command docker
Require-Command ssh
Require-Command scp

$workspace = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$source = Get-ChildItem -LiteralPath $workspace -Directory |
    Where-Object { Test-Path -LiteralPath (Join-Path $_.FullName 'Dockerfile') } |
    Select-Object -First 1
if (-not $source) {
    throw 'The local deployment source directory was not found.'
}
$context = $source.FullName
$dockerfile = Join-Path $context 'Dockerfile'

$safeName = $ContainerName -replace '[^A-Za-z0-9_.-]', '_'
$archive = Join-Path ([System.IO.Path]::GetTempPath()) "$safeName-$([guid]::NewGuid().ToString('N')).tar"
$remoteArchive = "/tmp/$([System.IO.Path]::GetFileName($archive))"
$destination = "$RemoteUser@$RemoteHost"

try {
    Write-Host "Building $ImageTag locally..."
    & docker build --tag $ImageTag --file $dockerfile $context
    if ($LASTEXITCODE -ne 0) { throw 'Local Docker build failed.' }

    Write-Host 'Exporting image archive...'
    & docker save --output $archive $ImageTag
    if ($LASTEXITCODE -ne 0) { throw 'Local Docker image export failed.' }

    Write-Host "Sending image to $RemoteHost over the LAN..."
    & scp $archive "${destination}:$remoteArchive"
    if ($LASTEXITCODE -ne 0) { throw 'SCP transfer failed.' }

    $remoteScript = @(
        'set -eu',
        "archive='$remoteArchive'",
        "image='$ImageTag'",
        "container='$ContainerName'",
        "service_port='$ServicePort'",
        "dev_port='$DevPort'",
        'cleanup() { rm -f "$archive"; }',
        'trap cleanup EXIT',
        'docker load --input "$archive"',
        'docker rm -f "$container" 2>/dev/null || true',
        'docker run -d --name "$container" -p "$service_port:15618" -p "$dev_port:59323" --log-opt max-size=10m --log-opt max-file=3 -w /app -e PORT=15618 -e PYTHONUNBUFFERED=1 -e PYTHONDONTWRITEBYTECODE=1 "$image" python backend/run.py',
        'sleep 3',
        'docker inspect --format "{{.State.Status}}" "$container" | grep -x running',
        'docker logs --tail 30 "$container"'
    ) -join '; '

    Write-Host 'Importing and replacing the remote container...'
    & ssh $destination $remoteScript
    if ($LASTEXITCODE -ne 0) { throw 'Remote deployment or health check failed.' }

    Write-Host "Deployment complete: http://${RemoteHost}:$ServicePort"
}
finally {
    if (Test-Path -LiteralPath $archive) {
        Remove-Item -LiteralPath $archive -Force
    }
}
