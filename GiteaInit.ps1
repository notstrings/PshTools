$ErrorActionPreference = "Stop"

. "$($PSScriptRoot)/ModuleMisc.ps1"

$Title    = [System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath)
$ConfPath = "$($PSScriptRoot)\Config\$($Title).json"

# セットアップ
function local:Setup() {
	winget install "Git.Git"
}

## 設定 #######################################################################

Add-Type -AssemblyName System.ComponentModel
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Windows.Forms.Design

Invoke-Expression -Command @"
class GiteaInitConf {
    [string] `$GITEAURL
    [string] `$GITEAORG
    [string] `$GITEAKEY
}
"@

# 設定初期化
function local:InitConf([string] $Path) {
    if ((Test-Path -LiteralPath $Path) -eq $false) {
        $conf = New-Object GiteaInitConf -Property @{
            GITEAURL = ""
            GITEAORG = ""
            GITEAKEY = ""
        }
        SaveConf $Path $conf
    }
}
# 設定書込
function local:SaveConf([string] $Path, [GiteaInitConf] $conf) {
    $null = New-Item ([System.IO.Path]::GetDirectoryName($Path)) -ItemType Directory -ErrorAction SilentlyContinue
    $conf | ConvertTo-Json | Out-File -FilePath $Path
}
# 設定読出
function local:LoadConf([string] $Path) {
    $json = Get-Content -Path $Path | ConvertFrom-Json
    $conf = ConvertFromPSCO ([GiteaInitConf]) $json
    return $conf
}
# 設定編集
function local:EditConf([string] $Title, [string] $Path) {
    $conf = LoadConf $Path
    $ret = ShowSettingDialog $Title $conf
    if ($ret -eq "OK") {
        SaveConf $Path $conf
    }
}

## 本体 #########################################################################

# ローカルリポジトリ存在確認
function local:IsGitInit([string] $Path) {
	return Test-Path -LiteralPath "$Path\.git" -PathType Container
}

# ローカルリポジトリ作成
function local:GitInit([string] $Path) {
	Push-Location $Path
	$null = Start-Process -NoNewWindow -Wait -FilePath "git.exe" -ArgumentList "init"
	Pop-Location
}

# リモートリポジトリ設定追加
function local:GitSetRemote([string] $Path, [string] $URL) {
	Push-Location $Path
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

function SetupGitea([string] $Path, [PSCustomObject] $conf) {
	# Giteaリポジトリ作成
	$Repository = [System.IO.Path]::GetFileName($Path)
	if (($Repository -match "[^a-zA-Z0-9-]")) {
		Write-Host "フォルダ名は英数ハイフンのみ使用可能"
		return
	}
	if ( (IsGiteaInit $conf.GITEAURL $conf.GITEAORG $Repository $conf.GITEAKEY) -eq $false){
		GiteaInit $conf.GITEAURL $conf.GITEAORG $Repository $conf.GITEAKEY
	}

	# ローカルリポジトリ作成＆関連付け
	if ( (IsGitInit $Path) -eq $false){
		GitInit $Path
	}
	GitSetRemote $Path "http://$URL/$ORG/$Repository.git"
}

## 本体 #######################################################################

# $args = @("$($ENV:USERPROFILE)\Desktop\bbb")

try {
	$null = Write-Host "---$Title---"
	# 設定取得
    InitConf $ConfPath
    $conf = LoadConf $ConfPath
	# 引数確認
    if ($args.Count -eq 0) {
        EditConf $Title $ConfPath
        exit
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
