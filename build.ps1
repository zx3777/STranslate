# 从命令行参数获取版本号,如果未提供则使用默认值
param(
    [string]$Version = "2.0.0"
)

$ErrorActionPreference = "Stop"

function Log([string]$msg, [string]$color = "Yellow") {
    Write-Host $msg -ForegroundColor $color
}

# 1. 清理版本号 (移除 'v' 前缀) -> 得到完整版本 (例如: 2.0.4-Image4AI)
$CleanVersion = $Version -replace '^v', ''

# 2. 提取纯数字版本号 (用于 AssemblyVersion) -> (例如: 2.0.4)
# 匹配开头的数字部分，忽略横杠后面的内容
if ($CleanVersion -match "^(\d+(\.\d+){0,3})") {
    $NumericVersion = $matches[1]
} else {
    $NumericVersion = "1.0.0.0" # 兜底
}

Log "构建信息:" "Cyan"
Log "  完整版本 (Tag/Display): $CleanVersion" "Cyan"
Log "  数字版本 (Assembly):    $NumericVersion" "Cyan"

# 更新./src/SolutionAssemblyInfo.cs 中的版本号
$asmInfo = "./src/SolutionAssemblyInfo.cs"

if (Test-Path $asmInfo) {
    # 读取文件内容
    $content = Get-Content $asmInfo -Raw

    # 定义替换规则
    # AssemblyVersion 和 AssemblyFileVersion 必须用纯数字 ($NumericVersion)
    # AssemblyInformationalVersion 可以用带后缀的完整版本 ($CleanVersion)
    $patterns = @{
        'AssemblyVersion'              = @{ Pattern = 'AssemblyVersion\("[^"]+"\)'; Value = "AssemblyVersion(`"$NumericVersion`")" }
        'AssemblyFileVersion'          = @{ Pattern = 'AssemblyFileVersion\("[^"]+"\)'; Value = "AssemblyFileVersion(`"$NumericVersion`")" }
        'AssemblyInformationalVersion' = @{ Pattern = 'AssemblyInformationalVersion\("[^"]+"\)'; Value = "AssemblyInformationalVersion(`"$CleanVersion`")" }
    }

    foreach ($key in $patterns.Keys) {
        $item = $patterns[$key]
        $content = [regex]::Replace($content, $item.Pattern, $item.Value)
    }

    # 写回文件
    Set-Content $asmInfo $content -Encoding UTF8

    Log "SolutionAssemblyInfo.cs 已更新。" "Green"
}
else {
    Log "未找到 $asmInfo，无法执行版本写入。" "Red"
    exit 1
}

# 清理构建输出
Log "正在清理之前的构建..."
$artifactPath = ".\src\.artifacts\Release\"

if (Test-Path $artifactPath) {
    Remove-Item -Path $artifactPath -Recurse -Force -ErrorAction SilentlyContinue
}

# 更新 Fody 配置文件
Log "正在更新 FodyWeavers..."

$src = "./src/STranslate/FodyWeavers.Release.xml"
$bak = "./src/STranslate/FodyWeavers.xml.bak"
$dst = "./src/STranslate/FodyWeavers.xml"

if (Test-Path $src) {
    Copy-Item $src $bak -Force
    Move-Item -Path $bak -Destination $dst -Force
} else {
    Log "未找到 $src，跳过更新。" "Red"
}

# 构建解决方案
# 注意：这里传递给 MSBuild 的属性也需要区分
# Version/PackageVersion 依然可以使用完整版
# FileVersion/AssemblyVersion 最好使用数字版(虽然 .NET SDK 会自动处理，但为了保险起见)
Log "正在重新生成解决方案..."
dotnet build .\src\STranslate.sln `
  --configuration Release `
  --no-incremental `
  /p:Version=$CleanVersion `
  /p:AssemblyVersion=$NumericVersion `
  /p:FileVersion=$NumericVersion `
  /p:InformationalVersion=$CleanVersion


# 还原 FodyWeavers.xml
Log "正在还原 FodyWeavers.xml SolutionAssemblyInfo.cs..."
git restore $dst $asmInfo

# 清理插件目录中多余文件
Log "正在清理多余的 STranslate.Plugin 文件..."

$pluginsPath = "./src/.artifacts/Release/Plugins"
if (Test-Path $pluginsPath) {
    Get-ChildItem -Path $pluginsPath -Recurse -Include "STranslate.Plugin.dll","STranslate.Plugin.xml" |
        Remove-Item -Force -ErrorAction SilentlyContinue
}

Log "构建完成！" "Green"
