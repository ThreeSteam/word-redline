param([string]$JobFile, [switch]$SmokeTest)
$ErrorActionPreference = 'Stop'

function GetSourceHash([string]$Path) {
    $stream = [IO.File]::OpenRead($Path)
    $sha = [Security.Cryptography.SHA256]::Create()
    try { return [BitConverter]::ToString($sha.ComputeHash($stream)).Replace('-', '') }
    finally { $sha.Dispose(); $stream.Dispose() }
}

if ($JobFile) {
    $word = $oldDoc = $newDoc = $result = $null
    $tempOutput = $null
    $wordProcess = $null
    function SetPhase($text) {
        @{ Phase=$text; At=(Get-Date).ToString('o') } | ConvertTo-Json | Set-Content -LiteralPath ($JobFile + '.phase') -Encoding UTF8
    }
    try {
        $job = Get-Content -LiteralPath $JobFile -Raw -Encoding UTF8 | ConvertFrom-Json
        foreach ($path in @($job.Old, $job.New)) {
            if (!(Test-Path -LiteralPath $path -PathType Leaf) -or [IO.Path]::GetExtension($path) -notin @('.docx','.doc','.docm')) { throw "文件不存在或格式不支持：$path" }
        }
        if ($job.Old -eq $job.New) { throw '请提供两个不同的文件。' }
        $folder = [IO.Path]::GetDirectoryName($job.New)
        if ($job.UseSubfolder) {
            $name = if ($null -eq $job.SubfolderName) { 'Legal' } else { [string]$job.SubfolderName }
            if ([string]::IsNullOrWhiteSpace($name) -or $name -ne $name.Trim() -or $name.EndsWith('.') -or $name -in @('.','..') -or $name.Length -gt 255 -or $name.IndexOfAny([IO.Path]::GetInvalidFileNameChars()) -ge 0 -or $name -match '^(CON|PRN|AUX|NUL|COM[1-9¹²³]|LPT[1-9¹²³])(?:\.|$)') {
                throw '子文件夹名称无效，请使用单个文件夹名称，不要包含路径或特殊字符。'
            }
            $folder = Join-Path $folder $name
            [void][IO.Directory]::CreateDirectory($folder)
        }
        $stem = [IO.Path]::GetFileNameWithoutExtension($job.New)
        $tempOutput = Join-Path $folder ('redline-tmp-' + [guid]::NewGuid().ToString('N') + '.docx')
        SetPhase '启动独立 Word 实例'
        . "$PSScriptRoot\WordNative.ps1"
        $wordExe = (Get-ItemProperty 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\WINWORD.EXE' -ErrorAction SilentlyContinue).'(default)'
        if (!$wordExe) {
            foreach ($key in @('Registry::HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\WINWORD.EXE','Registry::HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\WINWORD.EXE')) {
                $wordExe = (Get-ItemProperty $key -ErrorAction SilentlyContinue).'(default)'
                if ($wordExe) { break }
            }
        }
        if (!$wordExe -or !(Test-Path -LiteralPath $wordExe)) { throw '未找到 Microsoft Word 桌面版。请先安装并正常启动 Word。' }
        $wordProcess = Start-Process -FilePath $wordExe -ArgumentList '/x','/a','/w' -WindowStyle Hidden -PassThru
        @{ Id=$wordProcess.Id; Ticks=$wordProcess.StartTime.Ticks } | ConvertTo-Json | Set-Content -LiteralPath ($JobFile + '.owner') -Encoding UTF8
        for ($attempt=0; $attempt -lt 80; $attempt++) {
            if ($wordProcess.HasExited) { throw '独立 Word 实例提前退出。' }
            $handle = [WordNativeObject]::FindMainWindow($wordProcess.Id)
            if ($handle -ne [IntPtr]::Zero) {
                try { $native = [WordNativeObject]::FromMainWindow($handle); if ($native) { $word = $native.Application; break } } catch {}
            }
            Start-Sleep -Milliseconds 500
        }
        if (!$word) { throw 'Word 启动超时，请重试。' }
        $word.Visible = $false
        $word.DisplayAlerts = 0
        $word.AutomationSecurity = 3
        $originalUpdateLinks = $word.Options.UpdateLinksAtOpen
        $word.Options.UpdateLinksAtOpen = $false
        SetPhase '打开旧版文档'
        $oldDoc = $word.Documents.Open($job.Old, $false, $true, $false)
        SetPhase '打开新版文档'
        $newDoc = $word.Documents.Open($job.New, $false, $true, $false)
        SetPhase '比较文档'
        $result = $word.CompareDocuments($oldDoc, $newDoc, 2, 1, $true, $true, $true, $true, $true, $true, $true, $true, $true, $true, 'Redline', $true)
        $result.ShowRevisions = $true
        $revisionCount = [int]$result.Revisions.Count
        $noDifference = $revisionCount -eq 0
        if ($noDifference) {
            SetPhase '保存无差别记录'
            $tempOutput = Join-Path $folder ('redline-tmp-' + [guid]::NewGuid().ToString('N') + '.txt')
            $record = @(
                'Word Redline 2.2 比对记录'
                '比对结论：新旧文件无差别（Word 本次比较检测到 0 项修订）。'
                ('比对完成时间：' + [DateTimeOffset]::Now.ToString('yyyy-MM-dd HH:mm:ss zzz'))
                ('旧版文件名：' + [IO.Path]::GetFileName($job.Old))
                ('旧版完整路径：' + $job.Old)
                ('新版文件名：' + [IO.Path]::GetFileName($job.New))
                ('新版完整路径：' + $job.New)
                ('旧版 SHA-256：' + (GetSourceHash $job.Old))
                ('新版 SHA-256：' + (GetSourceHash $job.New))
                ('Microsoft Word 版本：' + $word.Version)
                '比较范围：字符级；格式、大小写、空格、表格、页眉页脚、脚注尾注、文本框、域、批注和移动。'
                '说明：此结论表示 Word 本次比较未检测到修订，不表示两个文件的二进制内容完全一致。'
                '本 TXT 为本地比对留痕，可编辑，不是数字签名或防篡改凭证。'
            ) -join "`r`n"
            [IO.File]::WriteAllText($tempOutput, $record + "`r`n", (New-Object Text.UTF8Encoding($true)))
        } else {
        SetPhase '保存修订结果'

        # Export the complete native comparison package, bypassing Word Save As hooks.
        [xml]$package = $result.WordOpenXML
        Add-Type -AssemblyName System.IO.Compression
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        $ns = New-Object Xml.XmlNamespaceManager($package.NameTable)
        $ns.AddNamespace('pkg','http://schemas.microsoft.com/office/2006/xmlPackage')
        $contentTypes = New-Object Xml.XmlDocument
        $types = $contentTypes.CreateElement('Types','http://schemas.openxmlformats.org/package/2006/content-types')
        [void]$contentTypes.AppendChild($types)
        $zip = [IO.Compression.ZipFile]::Open($tempOutput,[IO.Compression.ZipArchiveMode]::Create)
        try {
            foreach ($part in $package.SelectNodes('/pkg:package/pkg:part',$ns)) {
                $name = $part.GetAttribute('name',$ns.LookupNamespace('pkg'))
                $type = $part.GetAttribute('contentType',$ns.LookupNamespace('pkg'))
                $override = $contentTypes.CreateElement('Override',$types.NamespaceURI)
                $override.SetAttribute('PartName',$name)
                $override.SetAttribute('ContentType',$type)
                [void]$types.AppendChild($override)
                $entry = $zip.CreateEntry($name.TrimStart('/'))
                $stream = $entry.Open()
                try {
                    $binary = $part.SelectSingleNode('pkg:binaryData',$ns)
                    if ($binary) { $bytes = [Convert]::FromBase64String($binary.InnerText) }
                    else { $bytes = [Text.Encoding]::UTF8.GetBytes($part.SelectSingleNode('pkg:xmlData',$ns).InnerXml) }
                    $stream.Write($bytes,0,$bytes.Length)
                } finally { $stream.Dispose() }
            }
            $stream = $zip.CreateEntry('[Content_Types].xml').Open()
            try { $contentTypes.Save($stream) } finally { $stream.Dispose() }
        } finally { $zip.Dispose() }
        }
        $result.Close(0)
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($result)
        $result = $null
        $index = 1
        $label = if ($noDifference) { '无差别' } else { 'redline' }
        $extension = if ($noDifference) { '.txt' } else { '.docx' }
        while ($true) {
            $suffix = if ($index -eq 1) { " - $label" } else { " - $label ($index)" }
            $output = Join-Path $folder ($stem + $suffix + $extension)
            try { [IO.File]::Move($tempOutput, $output); break }
            catch [IO.IOException] { if ([IO.File]::Exists($output)) { $index++; continue }; throw }
        }
        @{ Success = $true; Output = $output; Revisions = $revisionCount; NoDifference = $noDifference } | ConvertTo-Json | Set-Content -LiteralPath ($JobFile + '.result') -Encoding UTF8
    } catch {
        @{ Success = $false; Error = $_.Exception.Message } | ConvertTo-Json | Set-Content -LiteralPath ($JobFile + '.result') -Encoding UTF8
    } finally {
        SetPhase '关闭比对实例'
        foreach ($doc in @($result,$newDoc,$oldDoc)) {
            if ($null -ne $doc) { try { $doc.Close(0) } catch {}; try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($doc) } catch {} }
        }
        if ($null -ne $word) { try { if ($null -ne $originalUpdateLinks) { $word.Options.UpdateLinksAtOpen = $originalUpdateLinks } } catch {}; try { $word.Quit(0) } catch {}; try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($word) } catch {} }
        if ($wordProcess -and !$wordProcess.HasExited) { try { $wordProcess.Kill() } catch {} }
        if ($tempOutput -and [IO.File]::Exists($tempOutput)) { Remove-Item -LiteralPath $tempOutput -Force }
    }
    exit
}
