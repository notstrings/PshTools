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
Invoke-Expression -Command @"
	class GiteaInitConf {
		[string] `$GITEAURL
		[string] `$GITEAORG
		[string] `$GITEAKEY
	}
"@

# 設定初期化
function local:InitConfFile([string] $Path) {
	if ((Test-Path -LiteralPath $Path) -eq $false) {
		$Conf = New-Object GiteaInitConf -Property @{
			GITEAURL = ""
			GITEAORG = ""
			GITEAKEY = ""
		}
		SaveConfFile $Path $Conf
    }
}
# 設定書込
function local:SaveConfFile([string] $Path, [GiteaInitConf] $Conf) {
    $null = New-Item ([System.IO.Path]::GetDirectoryName($Path)) -ItemType Directory -ErrorAction SilentlyContinue
    $Conf | ConvertTo-Json | Out-File -FilePath $Path
}
# 設定読出
function local:LoadConfFile([string] $Path) {
    $json = Get-Content -Path $Path | ConvertFrom-Json
    $Conf = ConvertFromPSCO ([GiteaInitConf]) $json
    return $Conf
}
# 設定編集
function local:EditConfFile([string] $Title, [string] $Path) {
    $Conf = LoadConfFile $Path
    $ret = ShowSettingDialog $Title $Conf
    if ($ret -eq "OK") {
        SaveConfFile $Path $Conf
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

function local:SetupGitea([string] $Path) {
	try {
		# 設定取得
		$Conf = LoadConfFile $ConfPath
		# Giteaリポジトリ作成
		$Repository = [System.IO.Path]::GetFileName($Path)
		if (($Repository -match "[^a-zA-Z0-9-]")) {
			Write-Host "フォルダ名は英数ハイフンのみ使用可能"
			return
		}
		if ( (IsGiteaInit $Conf.GITEAURL $Conf.GITEAORG $Repository $Conf.GITEAKEY) -eq $false){
			GiteaInit $Conf.GITEAURL $Conf.GITEAORG $Repository $Conf.GITEAKEY
		}
		# ローカルリポジトリ作成＆関連付け
		if ( (IsGitInit $Path) -eq $false){
			GitInit $Path
		}
		GitSetRemote $Path "http://$URL/$ORG/$Repository.git"
    } catch {
        $null = Write-Host "Error:" $_.Exception.Message
    }
}

## 本体 #######################################################################

# $args = @("$($ENV:USERPROFILE)\Desktop\bbb")

try {
	$null = Write-Host "---$Title---"
    # 設定初期化
    InitConfFile $ConfPath
	# 引数確認
    if ($args.Count -eq 0) {
        EditConfFile $Title $ConfPath
        exit
    }
	# 処理実行
    foreach ($arg in $args) {
        if (Test-Path -LiteralPath $arg) {
            if ([System.IO.Directory]::Exists($arg)) {
                SetupGitea $arg
            }
        }
    }
} catch {
    $null = Write-Host "---例外発生---"
    $null = Write-Host $_.Exception.Message
    $null = Write-Host $_.ScriptStackTrace
    $null = Write-Host "--------------"
}
