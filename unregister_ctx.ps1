# Removes the .docx right-click "HLT 파일로 변환" menu added by register_ctx.
Remove-Item -Path 'HKCU:\Software\Classes\SystemFileAssociations\.docx\shell\DocxToHlt' -Recurse -Force
Write-Output 'context menu removed'
