$ErrorActionPreference = "Stop"

. "$($PSScriptRoot)/ModuleMisc.ps1"

$Title    = [System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath)
$ConfPath = "$($PSScriptRoot)\Config\$($Title).json"

# セットアップ
function local:Setup() {
	winget install "Git.Git"
}

## 設定 #######################################################################

class Conf {
    [string] $GITEAURL
    [string] $GITEAORG
    [string] $GITEAKEY
}

# 設定初期化
function local:InitConf([string] $sPath) {
    if ((Test-Path -LiteralPath $sPath) -eq $false) {
        $conf = New-Object Conf -Property @{
            GITEAURL = ""
            GITEAORG = ""
            GITEAKEY = ""
        }
        SaveConf $sPath $conf
    }
}
# 設定書込
function local:SaveConf([string] $sPath, [Conf] $conf) {
    $null = New-Item ([System.IO.Path]::GetDirectoryName($sPath)) -ItemType Directory -ErrorAction SilentlyContinue
    $conf | ConvertTo-Json | Out-File -FilePath $sPath
}
# 設定読出
function local:LoadConf([string] $sPath) {
    $json = Get-Content -Path $sPath | ConvertFrom-Json
    $conf = ConvertFromPSCO ([Conf]) $json
    return $conf
}
# 設定編集
function local:EditConf([string] $sPath) {
    $conf = LoadConf $ConfPath
    $ret = ShowSettingDialog $Title $conf
    if ($ret -eq "OK") {
        SaveConf $ConfPath $conf
    }
}

## 本体 #########################################################################

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

## 本体 #######################################################################

# $args = @("$($ENV:USERPROFILE)\Desktop\bbb")

try {
	$null = Write-Host "---$Title---"
	# 設定取得
    InitConf $ConfPath
    $conf = LoadConf $ConfPath
	# 引数確認
    if ($args.Count -eq 0) {
        EditConf $ConfPath
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
