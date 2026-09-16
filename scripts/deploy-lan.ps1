[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$RemoteHost,

    [string]$RemoteUser = 'jl',

    [string]$IdentityFile = (Join-Path $env:USERPROFILE '.ssh\id_ed25519'),

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
if (-not (Test-Path -LiteralPath $IdentityFile)) {
    throw "SSH identity file not found: $IdentityFile"
}

$workspace = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$source = Get-ChildItem -LiteralPath $workspace -Directory |
    Where-Object {
        (Test-Path -LiteralPath (Join-Path $_.FullName 'Dockerfile')) -and
        (Test-Path -LiteralPath (Join-Path $_.FullName 'web\frontend')) -and
        (Test-Path -LiteralPath (Join-Path $_.FullName 'web\backend')) -and
        (Test-Path -LiteralPath (Join-Path $_.FullName 'web\requirements.txt'))
    } |
    Select-Object -First 1
if (-not $source) {
    throw 'The local PDF tool deployment source is incomplete.'
}
$source = $source.FullName
$webSource = Join-Path $source 'web'
$entryScripts = @(Get-ChildItem -LiteralPath $source -File -Filter '*.py')
if ($entryScripts.Count -ne 1) {
    throw 'The local PDF tool entry script could not be identified.'
}
$entryScript = $entryScripts[0]

$safeName = $ContainerName -replace '[^A-Za-z0-9_.-]', '_'
$archive = Join-Path ([System.IO.Path]::GetTempPath()) "$safeName-$([guid]::NewGuid().ToString('N')).tar"
$context = Join-Path ([System.IO.Path]::GetTempPath()) "$safeName-context-$([guid]::NewGuid().ToString('N'))"
$remoteArchive = "/tmp/$([System.IO.Path]::GetFileName($archive))"
$destination = "$RemoteUser@$RemoteHost"

try {
    # Dockerfile expects frontend/, backend/, requirements.txt, and the top-level Python entry point at context root.
    # Assemble that flat context without changing the user's source tree.
    New-Item -ItemType Directory -Path $context | Out-Null
    Copy-Item -LiteralPath (Join-Path $source 'Dockerfile') -Destination (Join-Path $context 'Dockerfile')
    Copy-Item -LiteralPath $entryScript.FullName -Destination $context
    Copy-Item -LiteralPath (Join-Path $webSource 'requirements.txt') -Destination (Join-Path $context 'requirements.txt')
    Copy-Item -LiteralPath (Join-Path $webSource 'frontend') -Destination (Join-Path $context 'frontend') -Recurse
    Copy-Item -LiteralPath (Join-Path $webSource 'backend') -Destination (Join-Path $context 'backend') -Recurse
    @('**/__pycache__/', '**/*.pyc', '.git/', 'frontend/node_modules/', 'frontend/dist/', 'backend/.tmp/') |
        Set-Content -LiteralPath (Join-Path $context '.dockerignore') -Encoding utf8

    Write-Host "Building $ImageTag locally..."
    & docker build --tag $ImageTag --file (Join-Path $context 'Dockerfile') $context
    if ($LASTEXITCODE -ne 0) { throw 'Local Docker build failed.' }

    Write-Host 'Exporting image archive...'
    & docker save --output $archive $ImageTag
    if ($LASTEXITCODE -ne 0) { throw 'Local Docker image export failed.' }

    Write-Host "Sending image to $RemoteHost over the LAN..."
    & scp -i $IdentityFile -o IdentitiesOnly=yes -o BatchMode=yes $archive "${destination}:$remoteArchive"
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
        'previous="${container}-previous-$(date +%Y%m%d%H%M%S)"',
        'if docker inspect "$container" >/dev/null 2>&1; then docker stop "$container"; docker rename "$container" "$previous"; fi',
        'docker run -d --name "$container" -p "$service_port:15618" -p "$dev_port:59323" --log-opt max-size=10m --log-opt max-file=3 -w /app -e PORT=15618 -e PYTHONUNBUFFERED=1 -e PYTHONDONTWRITEBYTECODE=1 "$image" python backend/run.py',
        'sleep 3',
        'if ! docker inspect --format "{{.State.Status}}" "$container" | grep -x running; then docker rm -f "$container" || true; if [ -n "${previous:-}" ]; then docker rename "$previous" "$container"; docker start "$container"; fi; exit 1; fi',
        'docker logs --tail 30 "$container"'
    ) -join '; '

    Write-Host 'Importing and replacing the remote container...'
    & ssh -i $IdentityFile -o IdentitiesOnly=yes -o BatchMode=yes $destination $remoteScript
    if ($LASTEXITCODE -ne 0) { throw 'Remote deployment or health check failed.' }

    Write-Host "Deployment complete: http://${RemoteHost}:$ServicePort"
}
finally {
    if (Test-Path -LiteralPath $archive) {
        Remove-Item -LiteralPath $archive -Force
    }
    if (Test-Path -LiteralPath $context) {
        Remove-Item -LiteralPath $context -Recurse -Force
    }
}
