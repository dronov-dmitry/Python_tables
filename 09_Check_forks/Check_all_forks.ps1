# CONFIG - только URL для клонирования, если нужно
$RepoUrl = "https://github.com/district0x/ethlance.git"
$LocalDir = $PWD.Path
$PerPage = 100

# HEADERS
$Headers = @{ "User-Agent" = "PowerShellScript" }
if ($env:GITHUB_TOKEN) { $Headers["Authorization"] = "token $env:GITHUB_TOKEN" }

# Функция для извлечения owner/repo из git remote URL
function Get-GitHubRepoInfo {
    try {
        # Получаем URL origin remote
        $remoteUrl = git config --get remote.origin.url
        if (-not $remoteUrl) {
            throw "No remote origin found"
        }
        
        Write-Host "Remote URL: $remoteUrl"
        
        # Парсим URL в зависимости от формата
        if ($remoteUrl -match 'github\.com[:\/](.+?)\/(.+?)(\.git)?$') {
            $owner = $matches[1]
            $repo = $matches[2]
            return @{ Owner = $owner; Repo = $repo }
        }
        elseif ($remoteUrl -match 'git@github\.com:(.+?)\/(.+?)(\.git)?$') {
            $owner = $matches[1]
            $repo = $matches[2]
            return @{ Owner = $owner; Repo = $repo }
        }
        else {
            throw "Unsupported URL format: $remoteUrl"
        }
    }
    catch {
        Write-Host "Error parsing git remote: $_"
        return $null
    }
}

# Получаем информацию о репозитории
$repoInfo = Get-GitHubRepoInfo

if ($repoInfo) {
    $RepoOwner = $repoInfo.Owner
    $RepoName = $repoInfo.Repo
    Write-Host "Detected: Owner = $RepoOwner, Repo = $RepoName"
} else {
    # Если не удалось определить, спрашиваем пользователя
    $RepoOwner = Read-Host "Enter GitHub owner (username/organization)"
    $RepoName = Read-Host "Enter GitHub repository name"
}

# Проверяем, что мы в git репозитории
if (-not (Test-Path ".git")) {
    Write-Host "Not a git repository. Cloning..."
    git clone $RepoUrl
    $repoNameFromUrl = ($RepoUrl -split '/' | Select-Object -Last 1) -replace '\.git$', ''
    Set-Location $repoNameFromUrl
    $LocalDir = Get-Location
}

Write-Host "Current directory: $(Get-Location)"
Write-Host "Repo Owner: $RepoOwner"
Write-Host "Repo Name: $RepoName"

# DETECT BASE BRANCH
$BaseBranch = ""
foreach ($c in @("master","main")) {
    git show-ref --verify --quiet "refs/heads/$c"
    if ($LASTEXITCODE -eq 0) { $BaseBranch = $c; break }
}
if (-not $BaseBranch) {
    git fetch origin
    foreach ($c in @("master","main")) {
        git show-ref --verify --quiet "refs/remotes/origin/$c"
        if ($LASTEXITCODE -eq 0) { $BaseBranch = $c; break }
    }
}
if (-not $BaseBranch) { 
    Write-Host "Available branches:"
    git branch -a
    throw "Cannot detect base branch" 
}

Write-Host "Base branch detected: $BaseBranch"
git checkout $BaseBranch
git pull origin $BaseBranch

# REPORT FILE
$ReportPath = Join-Path (Split-Path -Parent $LocalDir) "$RepoName-ForkReport.txt"
"" | Out-File -FilePath $ReportPath -Encoding UTF8

# Добавляем дату запуска в начало отчёта
("Script run date: " + (Get-Date -Format "yyyy-MM-dd HH:mm:ss")) | Out-File -FilePath $ReportPath -Append -Encoding UTF8
"`r`n" | Out-File -FilePath $ReportPath -Append -Encoding UTF8

# LOAD FORKS
$page = 1
$Forks = @()
while ($true) {
    $url = "https://api.github.com/repos/$RepoOwner/$RepoName/forks?per_page=$PerPage&page=$page"
    Write-Host "Fetching: $url"
    try {
        $data = Invoke-RestMethod -Uri $url -Headers $Headers -Method Get
        if (-not $data -or $data.Count -eq 0) { break }
        $Forks += $data
        $page++
        Write-Host "Found $($data.Count) forks on page $($page-1)"
    }
    catch {
        Write-Host "Error fetching forks: $_"
        break
    }
}

Write-Host "Total forks found: $($Forks.Count)"

# PROCESS FORKS
$total = $Forks.Count
$i = 0
foreach ($fork in $Forks) {
    $i++
    Write-Progress -Activity "Processing forks ($i of $total)" -Status "Fork: $($fork.owner.login)" -PercentComplete ([int](($i/$total)*100))

    $remoteName = $fork.owner.login
    $forkBranch = $fork.default_branch

    @"
===============================
FORK: $remoteName
DEFAULT BRANCH: $forkBranch
===============================

"@ | Out-File -FilePath $ReportPath -Append -Encoding UTF8

    $exists = git remote | Select-String -SimpleMatch $remoteName
    if (-not $exists) { 
        Write-Host "Adding remote: $remoteName"
        git remote add $remoteName $fork.clone_url 
    }

    Write-Host "Fetching from $remoteName..."
    git fetch $remoteName --prune | Out-Null

    $ref = "refs/remotes/" + $remoteName + "/" + $forkBranch
    git show-ref --verify --quiet $ref
    if ($LASTEXITCODE -ne 0) {
        "NO BRANCH " + $forkBranch + " FOUND - SKIPPED" | Out-File -FilePath $ReportPath -Append -Encoding UTF8
        continue
    }

    ("COMMITS NOT IN " + $BaseBranch + ":") | Out-File -FilePath $ReportPath -Append -Encoding UTF8
    $log = git log "$BaseBranch..$remoteName/$forkBranch" --oneline
    if ($log) { 
        $log | Out-File -FilePath $ReportPath -Append -Encoding UTF8
        Write-Host "Found changes in $remoteName"
    }
    else { 
        "-- NO CHANGES --" | Out-File -FilePath $ReportPath -Append -Encoding UTF8
    }

    "`r`n" | Out-File -FilePath $ReportPath -Append -Encoding UTF8
}

Write-Progress -Activity "Processing forks" -Completed
Write-Host "DONE"
Write-Host "Report saved to: $ReportPath"
