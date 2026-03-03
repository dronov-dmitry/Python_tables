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

# PROCESS FORKS AND COLLECT DATA
$forkData = @()

$total = $Forks.Count
$i = 0
foreach ($fork in $Forks) {
    $i++
    Write-Progress -Activity "Analyzing forks ($i of $total)" -Status "Fork: $($fork.owner.login)" -PercentComplete ([int](($i/$total)*100))

    $remoteName = $fork.owner.login
    $forkBranch = $fork.default_branch

    $exists = git remote | Select-String -SimpleMatch $remoteName
    if (-not $exists) { 
        Write-Host "Adding remote: $remoteName"
        git remote add $remoteName $fork.clone_url 
    }

    Write-Host "Fetching from $remoteName..."
    git fetch $remoteName --prune 2>&1 | Out-Null

    $ref = "refs/remotes/" + $remoteName + "/" + $forkBranch
    git show-ref --verify --quiet $ref
    if ($LASTEXITCODE -ne 0) {
        # Форк без доступной ветки - добавляем с 0 коммитов
        $forkData += [PSCustomObject]@{
            RemoteName = $remoteName
            ForkBranch = $forkBranch
            CreatedAt = $fork.created_at
            UpdatedAt = $fork.updated_at
            HtmlUrl = $fork.html_url
            CommitCount = 0
            LogDetailed = $null
            LogFull = $null
            HasChanges = $false
            Error = "NO BRANCH $forkBranch FOUND"
        }
        continue
    }

    # Получаем коммиты с датами
    $logDetailed = git log "$BaseBranch..$remoteName/$forkBranch" --pretty=format:"%h - %ad - %s" --date=short
    $logFull = git log "$BaseBranch..$remoteName/$forkBranch" --pretty=format:"Commit: %h%nAuthor: %an%nDate: %ad%nMessage: %s%n" --date=local
    
    if ($logDetailed) { 
        $commitCount = ($logDetailed | Measure-Object -Line).Lines
    }
    else {
        $commitCount = 0
    }
    
    # Сохраняем данные для сортировки
    $forkData += [PSCustomObject]@{
        RemoteName = $remoteName
        ForkBranch = $forkBranch
        CreatedAt = $fork.created_at
        UpdatedAt = $fork.updated_at
        HtmlUrl = $fork.html_url
        CommitCount = $commitCount
        LogDetailed = $logDetailed
        LogFull = $logFull
        HasChanges = ($commitCount -gt 0)
        Error = $null
    }
}

Write-Progress -Activity "Analyzing forks" -Completed

# REPORT FILE
$ReportPath = Join-Path (Split-Path -Parent $LocalDir) "$RepoName-ForkReport.txt"
"" | Out-File -FilePath $ReportPath -Encoding UTF8

# Добавляем дату запуска в начало отчёта
("Script run date: " + (Get-Date -Format "yyyy-MM-dd HH:mm:ss")) | Out-File -FilePath $ReportPath -Append -Encoding UTF8
"`r`n" | Out-File -FilePath $ReportPath -Append -Encoding UTF8

# Сортируем форки по количеству коммитов (по убыванию)
$sortedForks = $forkData | Sort-Object -Property CommitCount -Descending

# Считаем статистику для сводки
$totalForksWithChanges = ($sortedForks | Where-Object { $_.HasChanges }).Count
$totalCommits = ($sortedForks | Measure-Object -Property CommitCount -Sum).Sum

# Добавляем сводку в начало отчёта
@"
=================================
       SUMMARY REPORT
=================================
Total forks: $($sortedForks.Count)
Forks with changes: $totalForksWithChanges
Total commits in forks: $totalCommits
Sorted by: number of changes (descending)

TOP-5 FORKS BY CHANGES:
"@ | Out-File -FilePath $ReportPath -Append -Encoding UTF8

$topForks = $sortedForks | Where-Object { $_.HasChanges } | Select-Object -First 5
foreach ($fork in $topForks) {
    "  $($fork.CommitCount) commits - $($fork.RemoteName) (updated: $($fork.UpdatedAt))" | Out-File -FilePath $ReportPath -Append -Encoding UTF8
}

"`r`n" + ("="*60) + "`r`n" | Out-File -FilePath $ReportPath -Append -Encoding UTF8

# PROCESS SORTED FORKS FOR DETAILED REPORT
$i = 0
foreach ($fork in $sortedForks) {
    $i++
    Write-Progress -Activity "Writing report ($i of $($sortedForks.Count))" -Status "Fork: $($fork.RemoteName)" -PercentComplete ([int](($i/$sortedForks.Count)*100))

    @"
===============================
[$i] FORK: $($fork.RemoteName)
CHANGES: $($fork.CommitCount) commits
DEFAULT BRANCH: $($fork.ForkBranch)
CREATED: $($fork.CreatedAt)
UPDATED: $($fork.UpdatedAt)
URL: $($fork.HtmlUrl)
===============================

"@ | Out-File -FilePath $ReportPath -Append -Encoding UTF8

    if ($fork.Error) {
        "ERROR: $($fork.Error)" | Out-File -FilePath $ReportPath -Append -Encoding UTF8
    }
    elseif ($fork.HasChanges) {
        ("COMMITS NOT IN " + $BaseBranch + ":") | Out-File -FilePath $ReportPath -Append -Encoding UTF8
        
        "`nSHORT FORMAT:" | Out-File -FilePath $ReportPath -Append -Encoding UTF8
        $fork.LogDetailed | Out-File -FilePath $ReportPath -Append -Encoding UTF8
        
        "`nDETAILED FORMAT:" | Out-File -FilePath $ReportPath -Append -Encoding UTF8
        $fork.LogFull | Out-File -FilePath $ReportPath -Append -Encoding UTF8
        
        # Статистика
        $firstCommit = $fork.LogDetailed | Select-Object -Last 1
        $lastCommit = $fork.LogDetailed | Select-Object -First 1
        
        "`nSTATISTICS:" | Out-File -FilePath $ReportPath -Append -Encoding UTF8
        "Total commits: $($fork.CommitCount)" | Out-File -FilePath $ReportPath -Append -Encoding UTF8
        if ($firstCommit) { "First commit: $firstCommit" | Out-File -FilePath $ReportPath -Append -Encoding UTF8 }
        if ($lastCommit) { "Last commit: $lastCommit" | Out-File -FilePath $ReportPath -Append -Encoding UTF8 }
    }
    else {
        "-- NO CHANGES --" | Out-File -FilePath $ReportPath -Append -Encoding UTF8
    }

    "`r`n" + ("="*60) + "`r`n" | Out-File -FilePath $ReportPath -Append -Encoding UTF8
}

Write-Progress -Activity "Writing report" -Completed

# Добавляем итоговую статистику в конец
@"

=================================
       FINAL STATISTICS
=================================
Total forks processed: $($sortedForks.Count)
Forks with changes: $totalForksWithChanges
Forks without changes: $($sortedForks.Count - $totalForksWithChanges)
Total commits: $totalCommits

Distribution by commit count:
"@ | Out-File -FilePath $ReportPath -Append -Encoding UTF8

$groups = $sortedForks | Where-Object { $_.HasChanges } | Group-Object { 
    if ($_.CommitCount -eq 1) { "1 commit" }
    elseif ($_.CommitCount -le 5) { "2-5 commits" }
    elseif ($_.CommitCount -le 20) { "6-20 commits" }
    else { "More than 20 commits" }
}

foreach ($group in $groups | Sort-Object Name) {
    "  $($group.Name): $($group.Count) forks" | Out-File -FilePath $ReportPath -Append -Encoding UTF8
}

"`r`nReport generated: " + (Get-Date -Format "yyyy-MM-dd HH:mm:ss") | Out-File -FilePath $ReportPath -Append -Encoding UTF8

Write-Host "DONE"
Write-Host "Report saved to: $ReportPath"
