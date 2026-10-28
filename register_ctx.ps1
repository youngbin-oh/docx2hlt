# Registers a right-click "HLT 파일로 변환" verb for .docx files.
# Re-run this if you move the docx2hlt folder. No admin needed (HKCU).
$base = 'HKCU:\Software\Classes\SystemFileAssociations\.docx\shell\DocxToHlt'
$cmd  = 'cmd.exe /c ""' + $PSScriptRoot + '\docx2hlt_ctx.cmd" "%1""'
$icon = 'C:\KipoNet\NKEditor\NKEditor.exe,0'
New-Item -Path $base -Force | Out-Null
Set-ItemProperty -Path $base -Name '(default)' -Value 'HLT 파일로 변환'
Set-ItemProperty -Path $base -Name 'Icon' -Value $icon
# 'Document' = one invocation receiving every selected .docx, so a
# multi-select right-click converts them sequentially instead of launching
# a parallel NKEditor instance per file.
Set-ItemProperty -Path $base -Name 'MultiSelectModel' -Value 'Document'
New-Item -Path "$base\command" -Force | Out-Null
Set-ItemProperty -Path "$base\command" -Name '(default)' -Value $cmd
Write-Output 'context menu registered'
Write-Output "command = $cmd"
