# Run the same document commands used by the AI Office toolbar.
$officeRoot = Split-Path -Parent $PSScriptRoot
$officeNode = Join-Path $officeRoot '.build/node-24.18/node-v24.18.0-win-x64/node.exe'
if (-not (Test-Path -LiteralPath $officeNode)) { $officeNode = 'node' }
& $officeNode (Join-Path $officeRoot 'extensions/ai-office-pdf/out/cli.cjs') @args
exit $LASTEXITCODE
