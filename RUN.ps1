powershell -NoExit -Command @"
Set-Location -LiteralPath '$PWD'
python .\cli_tool.py
"@