$msbuild = "E:\VisualStudio\MSBuild\15.0\Bin\MSBuild.exe"
function Build
{
    param(
        [string]$target = "both"
    )

    $config = "/p:Configuration=Release"
    $x64 = "/p:Platform=X64"
    $x86 = "/p:Platform=X86"

    switch($target)
    {
        "both" {
            &$msbuild $config $x64
            &$msbuild $config $x86
        }
        "64" { &$msbuild $config $x64 }
        "86" { &$msbuild $config $x86 }
    }
}

function Dist
{
    param (
        [Parameter(Mandatory = $true)][int16]$major,
        [Parameter(Mandatory = $true)][int16]$minor,
        [Parameter(Mandatory = $true)][int16]$patch
    )
    Remove-Item -Recurse .\obj, .\bin, .\dist -ErrorAction SilentlyContinue

    $ver = "$($major).$($minor).$($patch)"
    BumpVersion $ver
    Build

    New-Item -ItemType directory .\dist -ErrorAction SilentlyContinue

    Compress-Archive -Path .\bin\x64, .\bin\x86 -DestinationPath ".\dist\AppVolume_$($ver)_x64_x86_dll.zip"

    &".\SkinPackager.exe" .\skinDefinition.json
}

function BumpVersion
{
    param (
        [Parameter(Mandatory = $true)][string]$ver
    )

    $resFile = ".\AppVolume.rc"
    $prop = Get-Content $resFile -Encoding Unicode
    $prop = $prop -replace '"FileVersion", "[\d\.]*"', "`"FileVersion`", `"$ver.0`""
    $verComma = $ver -replace "\.", ","
    $prop = $prop -replace 'FILEVERSION [\d,]*', "FILEVERSION $verComma,0"

    $prop | Set-Content $resFile -Encoding Unicode


    $skinDef = Get-Content ".\skinDefinition.json" | ConvertFrom-Json
    $skinDef.version = $ver
    $skinDef.output = "./dist/AppVolume_$($ver).rmskin"
    $skinDef | ConvertTo-Json | Set-Content ".\skinDefinition.json" -Encoding UTF8
}