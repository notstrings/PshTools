$ErrorActionPreference = "Stop"

# セットアップ
function local:Setup() {
	winget install "Git.Git"
}

# ローカルリポジトリ存在確認
function local:IsGitInit([string] $sPath) {
	return Test-Path -LiteralPath "$sPath\.git" -PathType Container
}

# ローカルリポジトリ作成
function local:GitInit([string] $sPath) {
	Push-Location $sPath
	$null = Start-Process -NoNewWindow -Wait -FilePath "git.exe" -ArgumentList "init"
	Pop-Location
}

# リモートリポジトリ設定追加
function local:GitSetRemote([string] $sPath, [string] $URL) {
	Push-Location $sPath
	$null = Start-Process -NoNewWindow -Wait -FilePath "git.exe" -ArgumentList "remote remove origin"
	$null = Start-Process -NoNewWindow -Wait -FilePath "git.exe" -ArgumentList "remote add origin $URL"
	Pop-Location
}

# Giteaリモートリポジトリ存在確認
function local:IsGiteaInit([string] $URL, [string] $ORG, [string] $Repository, [string] $Key) {
	$Ret = $false
	$rslt = Invoke-RestMethod `
		-Method Get `
		-Uri "http://$URL/api/v1/orgs/$ORG/repos?access_token=$Key" `
		-ContentType 'application/json'
	foreach ($elm in $rslt) {
		If ($elm.name -eq $Repository) {
			$Ret = $true
		}
	}
	return $Ret
}

# Giteaリモートリポジトリ作成
function local:GiteaInit([string] $URL, [string] $ORG, [string] $Repository, [string] $Key) {
	$null = Invoke-RestMethod `
		-Method Post `
		-Uri "http://$URL/api/v1/orgs/$ORG/repos?access_token=$Key" `
		-ContentType 'application/json' `
		-Body (	[System.Text.Encoding]::UTF8.GetBytes((ConvertTo-Json @{Name = "$Repository"} )) )
}

function SetupGitea([string] $sPath, [PSCustomObject] $conf) {
	# Giteaリポジトリ作成
	$Repository = [System.IO.Path]::GetFileName($sPath)
	if (($Repository -match "[^a-zA-Z0-9-]")) {
		Write-Host "フォルダ名は英数ハイフンのみ使用可能"
		return
	}
	if ( (IsGiteaInit $conf.GITEAURL $conf.GITEAORG $Repository $conf.GITEAKEY) -eq $false){
		GiteaInit $conf.GITEAURL $conf.GITEAORG $Repository $conf.GITEAKEY
	}

	# ローカルリポジトリ作成＆関連付け
	if ( (IsGitInit $sPath) -eq $false){
		GitInit $sPath 
	}
	GitSetRemote $sPath "http://$URL/$ORG/$Repository.git"
}

# $args = @("$($ENV:USERPROFILE)\Desktop\bbb")

try {
	$null = Write-Host "---GiteaInit---"
	# 引数確認
    if ($args.Length -eq 0) {
        exit 1
    }
	# 設定取得
	$ParamPath = "$($PSScriptRoot)\Config\GiteaInit.json"
    if ((Test-Path -LiteralPath $ParamPath) -eq $false) {
		$conf = @{
			GITEAURL = "";
			GITEAORG = "";
			GITEAKEY = "";
		}
        $null = New-Item ([System.IO.Path]::GetDirectoryName($ParamPath)) -ItemType Directory -ErrorAction SilentlyContinue
		$conf | ConvertTo-Json | Out-File -FilePath $ParamPath
		exit 1
    }
    $conf = Get-Content -Path $ParamPath | ConvertFrom-Json
	if ($conf.GITEAURL -eq "" -or $conf.GITEAORG -eq "" -or $conf.GITEAKEY -eq "") {
		exit 1
	}
	# 処理実行
    foreach ($arg in $args) {
        if (Test-Path -LiteralPath $arg) {
            if ([System.IO.Directory]::Exists($arg)) {
                SetupGitea $arg $conf
            }
        }
    }
} catch {
    $null = Write-Host "---例外発生---"
    $null = Write-Host $_.Exception.Message
    $null = Write-Host $_.ScriptStackTrace
    $null = Write-Host "--------------"
}
