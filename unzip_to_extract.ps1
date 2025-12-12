# エンコーディング設定（UTF-8）
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

# エラー時に停止
$ErrorActionPreference = "Stop"

try {
    # 7z.exe のパス（7-Zipインストールパスに応じて変更）
    $ZIP_EXE = "C:\Program Files\7-Zip\7z.exe"

    # 解凍先フォルダ
    $DEST_DIR = "C:\Users\EIDAI-20240217-2\Downloads"

    # ZIPファイルパスを取得
    $zipPath = $args[0]

    if (-not $zipPath) {
        Write-Host "[ERROR] ZIPファイルが指定されていません。"
        pause
        exit 1
    }

    # パスを正規化（引用符を削除）
    $zipPath = $zipPath.Trim([char]34)

    Write-Host "[INFO] ZIPファイル: $zipPath"

    # ZIPファイルの存在確認
    if (-not (Test-Path $zipPath)) {
        Write-Host "[ERROR] ZIPファイルが見つかりません: $zipPath"
        pause
        exit 1
    }

    # 7z.exeの存在確認
    if (-not (Test-Path $ZIP_EXE)) {
        Write-Host "[ERROR] 7z.exeが見つかりません: $ZIP_EXE"
        pause
        exit 1
    }

    # ZIPファイル名（拡張子除く）を取得
    $zipName = [System.IO.Path]::GetFileNameWithoutExtension($zipPath)
    Write-Host "[INFO] ZIP名: $zipName"

    # 仮のリストファイルでルート構造を調べる
    $tempListFile = [System.IO.Path]::GetTempFileName()
    Write-Host "[INFO] ルート構造を確認中..."
    
    $process = Start-Process -FilePath $ZIP_EXE -ArgumentList "l", "`"$zipPath`"" -NoNewWindow -Wait -PassThru -RedirectStandardOutput $tempListFile
    
    if ($process.ExitCode -ne 0) {
        Write-Host "[ERROR] ZIPファイルのリスト取得に失敗しました。"
        Remove-Item $tempListFile -ErrorAction SilentlyContinue
        pause
        exit 1
    }

    # ルートフォルダがあるか調べる
    # 7-Zipのリスト出力では、フォルダは属性が "D...." で表示される
    # ルートフォルダがある場合: すべてのエントリが同じルートフォルダ名で始まる
    # ルートフォルダがない場合: エントリが直接ルートレベルに存在（パスにバックスラッシュがないエントリが複数存在）
    $listContent = Get-Content $tempListFile -ErrorAction Stop
    $rootFolders = @()  # ルートレベルのフォルダを収集
    $rootLevelEntries = @()  # ルートレベルのエントリ（ファイルとフォルダ）を収集
    
    foreach ($line in $listContent) {
        $trimmedLine = $line.Trim()
        if ($trimmedLine -eq "") { continue }
        
        # ヘッダー行や区切り線をスキップ
        if ($trimmedLine -match "^-+$" -or $trimmedLine -match "^Date" -or $trimmedLine -match "^Path =") { continue }
        
        # 7-Zipのリスト出力は固定幅フォーマット
        # 形式: "2025-12-11 13:29:41 D....            0            0  fenrir075c"
        # 日付形式（YYYY-MM-DD）で始まる行のみ処理
        if ($trimmedLine -match "^[0-9]{4}-[0-9]{2}-[0-9]{2}\s+[0-9]{2}:[0-9]{2}:[0-9]{2}\s+(D\.\.\.\.|\.\.\.\.\.)\s+\d+\s+\d+\s+(.+)$") {
            $attr = $matches[1]  # 属性（D.... がフォルダ）
            $pathPart = $matches[2].Trim()  # パス部分
            
            # パスにバックスラッシュが含まれていない場合、ルートレベルのエントリ
            if (-not $pathPart.Contains("\")) {
                $rootLevelEntries += $pathPart
                # フォルダの場合
                if ($attr -eq "D....") {
                    $rootFolders += $pathPart
                }
            }
        }
    }
    
    # ルートフォルダの判定ロジック:
    # ルートレベルのエントリが1つだけで、それがフォルダの場合 → ルートフォルダあり
    # ルートレベルのエントリが複数ある場合 → ルートフォルダなし
    if ($rootLevelEntries.Count -eq 1 -and $rootFolders.Count -eq 1) {
        $rootFolderCount = 1
    } else {
        $rootFolderCount = 0
    }

    Write-Host "[INFO] ルートフォルダ数: $rootFolderCount"

    if ($rootFolderCount -ge 1) {
        Write-Host "[INFO] ルートフォルダあり → そのまま解凍"
        $extractProcess = Start-Process -FilePath $ZIP_EXE -ArgumentList "x", "`"$zipPath`"", "-o`"$DEST_DIR`"", "-y" -NoNewWindow -Wait -PassThru
    } else {
        Write-Host "[INFO] ルートフォルダなし → ZIP名のフォルダに解凍"
        $targetDir = Join-Path $DEST_DIR $zipName
        $extractProcess = Start-Process -FilePath $ZIP_EXE -ArgumentList "x", "`"$zipPath`"", "-o`"$targetDir`"", "-y" -NoNewWindow -Wait -PassThru
    }

    if ($extractProcess.ExitCode -ne 0) {
        Write-Host "[ERROR] 解凍に失敗しました。終了コード: $($extractProcess.ExitCode)"
        Remove-Item $tempListFile -ErrorAction SilentlyContinue
        pause
        exit 1
    }

    # 一時ファイルを削除
    Remove-Item $tempListFile -ErrorAction SilentlyContinue

    Write-Host "解凍完了。"
}
catch {
    Write-Host "[ERROR] エラーが発生しました: $($_.Exception.Message)"
    Write-Host $_.ScriptStackTrace
    pause
    exit 1
}
