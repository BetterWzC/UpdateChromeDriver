param()
if ($args.Length -eq 0)
{
    $DownloadPath = $(Get-Location).Path
    $ExtractPath = $(Get-Location).Path
}
elseif ( $args.Length -eq 1)
{
    $DownloadPath = $args[0]
    $ExtractPath = $args[0]
    
}
elseif ($args.Length -eq 2)
{
    $DownloadPath, $ExtractPath = $args
}else
{
    "Invalid parameters"
    exit
}

function GetChromeVersion {
    param ()
    $registry = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Chrome.exe"
    $default = $registry."(default)"
    $chromeVersion = $(Get-Item $default).VersionInfo.ProductVersion.Split(".")[0]
    return $chromeVersion
}
function GetTheLatestBuild {
    param ($chromeVersion)
    $url = "http://npm.taobao.org/mirrors/chromedriver"
    $response = Invoke-WebRequest -URI $url -UseBasicParsing
    $Links = $response.Links | ForEach-Object {$_.href}
    $regex = "/mirrors/chromedriver/" + $chromeVersion + "\.?[0-9]+(\.[0-9]*)*/"
    $buildLink = $Links | Where-Object {$_ -match $regex } | Select-Object -Last 1
    $buildNumber = $buildLink.Split('/')[-2]
    return $buildNumber
}

function DownloadChromedriver {
    param ($buildNumber, $DownloadPath)
    $url = "http://npm.taobao.org/mirrors/chromedriver/" + $buildNumber + "/chromedriver_win32.zip"
    $DownloadPath = $DownloadPath + "\chromedriver_win32.zip"
    Remove-Item -Path $DownloadPath -Force -Recurse -EA SilentlyContinue
    try{
        $webClient = New-Object System.Net.WebClient
        $webClient.DownloadFile($url, $DownloadPath)
    }
    catch{
        "Failed to download!"
    }

}
function ExtractFile {
    param ($DownloadPath, $ExtractPath)
    $DownloadPath = $DownloadPath + "\chromedriver_win32.zip"
    $ExtractPath = $ExtractPath + "\chromedriver.exe"
    Stop-Process -Name "chromedriver" -EA SilentlyContinue
    Start-Sleep -s 3
    Remove-Item -Path $ExtractPath -Force -Recurse -EA SilentlyContinue
    Expand-Archive -Path $DownloadPath -Destination $ExtractPath  
}

function TestPath {
    param ($Path)
    if ($(Test-Path $Path))
    {
        return
    }
    else 
    {
        New-Item -ItemType "directory" -Path $Path | Out-Null
        return
    }
}

#Main
###########################################################
# $errpref = $ErrorActionPreference #save actual preference
# $ErrorActionPreference = "silentlycontinue"

"=" * 50
"DownloadPath: $DownloadPath"
"ExtractPath: $ExtractPath"
"=" * 50

TestPath $DownloadPath
TestPath $ExtractPath
$chromeVersion = GetChromeVersion
$buildNumber = GetTheLatestBuild $chromeVersion
DownloadChromedriver $buildNumber $DownloadPath
ExtractFile $DownloadPath $ExtractPath
# $ErrorActionPreference = $errpref #restore previous preference

