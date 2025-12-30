# CONFIG
$RepoOwner = "LibreOffice"
$RepoName  = "core"
$RepoUrl   = "https://github.com/$RepoOwner/$RepoName.git"
$LocalDir  = Join-Path $PWD $RepoName
$PerPage   = 100

# HEADERS
$Headers = @{ "User-Agent" = "PowerShellScript" }
if ($env:GITHUB_TOKEN) { $Headers["Authorization"] = "token $env:GITHUB_TOKEN" }

# CLONE IF NEEDED
if (-not (Test-Path $LocalDir)) { git clone $RepoUrl }
Set-Location $LocalDir

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
if (-not $BaseBranch) { throw "Cannot detect base branch" }

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
	$data = Invoke-RestMethod -Uri $url -Headers $Headers -Method Get
	if (-not $data -or $data.Count -eq 0) { break }
	$Forks += $data
	$page++
}

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
	if (-not $exists) { git remote add $remoteName $fork.clone_url }

	git fetch $remoteName --prune | Out-Null

	$ref = "refs/remotes/" + $remoteName + "/" + $forkBranch
	git show-ref --verify --quiet $ref
	if ($LASTEXITCODE -ne 0) {
		"NO BRANCH " + $forkBranch + " FOUND - SKIPPED" | Out-File -FilePath $ReportPath -Append -Encoding UTF8
		continue
	}

	("COMMITS NOT IN " + $BaseBranch + ":") | Out-File -FilePath $ReportPath -Append -Encoding UTF8
	$log = git log "$BaseBranch..$remoteName/$forkBranch" --oneline
	if ($log) { $log | Out-File -FilePath $ReportPath -Append -Encoding UTF8 }
	else { "-- NO CHANGES --" | Out-File -FilePath $ReportPath -Append -Encoding UTF8 }

	"`r`n" | Out-File -FilePath $ReportPath -Append -Encoding UTF8
}

Write-Progress -Activity "Processing forks" -Completed
Write-Host "DONE"
Write-Host "Report saved to: $ReportPath"
