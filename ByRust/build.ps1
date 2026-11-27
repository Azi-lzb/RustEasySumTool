param([switch]$SkipChecks)
$ErrorActionPreference = 'Stop'
$rustProject = $PSScriptRoot
$projectRoot = Split-Path -Parent $rustProject
$localTools = Join-Path $projectRoot '.build/rust-tools'
if (-not (Get-Command cargo -ErrorAction SilentlyContinue)) {
    $localCargo = Join-Path $localTools 'cargo/bin'
    if (-not (Test-Path -LiteralPath (Join-Path $localCargo 'cargo.exe'))) { throw '请安装Rust stable及MSVC C++工具链，再执行本脚本。' }
    $env:RUSTUP_HOME = Join-Path $localTools 'rustup'
    $env:CARGO_HOME = Join-Path $localTools 'cargo'
    $env:PATH = $localCargo + ';' + $env:PATH
}
$manifest = Join-Path $rustProject 'Cargo.toml'
if (-not $SkipChecks) {
    & cargo fmt --manifest-path $manifest --check
    if ($LASTEXITCODE -ne 0) { throw 'Rust格式检查失败。' }
    & cargo clippy --manifest-path $manifest --all-targets --locked -- -D warnings
    if ($LASTEXITCODE -ne 0) { throw 'Rust静态检查失败。' }
    & cargo test --manifest-path $manifest --locked
    if ($LASTEXITCODE -ne 0) { throw 'Rust测试失败。' }
}
& cargo build --release --locked --manifest-path $manifest
if ($LASTEXITCODE -ne 0) { throw 'Rust发行版编译失败。' }
$metadata = & cargo metadata --format-version 1 --no-deps --manifest-path $manifest | ConvertFrom-Json
if ($LASTEXITCODE -ne 0) { throw '无法读取Cargo元数据。' }
$version = $metadata.packages[0].version
$distribution = Join-Path $projectRoot 'dist'
New-Item -ItemType Directory -Path $distribution -Force | Out-Null
$executable = Join-Path $distribution "Excel 模板汇总工具_Rust_v$version.exe"
Copy-Item -LiteralPath (Join-Path $metadata.target_directory 'release/excel-summary-rust.exe') -Destination $executable -Force
Copy-Item -LiteralPath (Join-Path $projectRoot 'config.xlsx') -Destination (Join-Path $distribution 'config.xlsx') -Force
$manual = Join-Path $projectRoot '使用说明.docx'
if (-not (Test-Path -LiteralPath $manual)) { throw '缺少使用说明.docx，请将Word使用说明放在项目根目录。' }
$manualOutput = Join-Path $distribution '使用说明.docx'
Copy-Item -LiteralPath $manual -Destination $manualOutput -Force
$zipPath = Join-Path $distribution 'Excel 模板汇总工具_Rust.zip'
Compress-Archive -LiteralPath @($executable,(Join-Path $distribution 'config.xlsx'),$manualOutput) -DestinationPath $zipPath -Force
Write-Host "已打包 $executable"
Write-Host "压缩包 $zipPath"
