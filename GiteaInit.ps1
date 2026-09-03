$ErrorActionPreference = "Stop"

. "$($PSScriptRoot)/ModuleMisc.ps1"

$Title    = [System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath)
$ConfPath = "$($PSScriptRoot)\Config\$($Title).json"

# セットアップ
function local:Setup() {
	winget install "Git.Git"
	winget install "Gitea.Tea"
	# ココまでやっといてね！
	# tea login add --name nas0001 --url http://nas0001:3333 --token xxxxxxxxx
	# tea login default nas00001
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

function local:SetupGitea([string] $Path) {
	try {
		Push-Location $Path

		# リポジトリ名制限
		$Repository = [System.IO.Path]::GetFileName($Path)
		if (($Repository -match "[^a-zA-Z0-9-]")) {
			Write-Host "リポジトリフォルダ名は英数ハイフンのみ使用可能です"
			return
		}

		tea repos create --owner $Conf.GITEAORG --name $Repository
		if (-not (Test-Path "$Path\.git")) {
			git.exe init
		}
		git.exe remote remove origin
		git.exe remote add origin "http://$($Conf.GITEAURL)/$($Conf.GITEAORG)/$Repository.git"
	} catch {
		$null = Write-Host "Error:" $_.Exception.Message
	} finally {
		Pop-Location
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
	$Conf = LoadConfFile $ConfPath
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
