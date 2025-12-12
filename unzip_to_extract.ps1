# エンコーディング設定（UTF-8）
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

# エラー時に停止
$ErrorActionPreference = "Stop"

try {
    # スクリプトのディレクトリを引数から取得
    $scriptDir = $args[0]
    if ([string]::IsNullOrWhiteSpace($scriptDir)) {
        # フォールバック: 現在のディレクトリを使用
        $scriptDir = Get-Location
    }
    # 末尾のバックスラッシュを削除
    $scriptDir = $scriptDir.TrimEnd('\')
    
    # JSON設定ファイルのパス（.jsoncまたは.json）
    $configFile = Join-Path $scriptDir "unzip_to_extract_config.jsonc"
    if (-not (Test-Path $configFile)) {
        $configFile = Join-Path $scriptDir "unzip_to_extract_config.json"
    }
    
    # JSON設定ファイルから設定を読み込む
    $ZIP_EXE_DIR = $null
    $DEST_DIR = $null
    
    if (Test-Path $configFile) {
        try {
            # JSONファイルを読み込み、コメントを削除してからパース
            $jsonContent = Get-Content $configFile -Raw -Encoding UTF8
            # 行コメント（//）を削除
            $jsonContent = $jsonContent -replace '//.*', ''
            # ブロックコメント（/* */）を削除
            $jsonContent = $jsonContent -replace '/\*[\s\S]*?\*/', ''
            
            $config = $jsonContent | ConvertFrom-Json
            if ($config.zipExeDir) {
                $ZIP_EXE_DIR = $config.zipExeDir
                # 環境変数を展開
                $ZIP_EXE_DIR = [System.Environment]::ExpandEnvironmentVariables($ZIP_EXE_DIR)
            }
            if ($config.destDir) {
                $DEST_DIR = $config.destDir
                # 環境変数を展開
                $DEST_DIR = [System.Environment]::ExpandEnvironmentVariables($DEST_DIR)
            }
            Write-Host "[INFO] 設定ファイルを読み込みました: $configFile"
        }
        catch {
            Write-Host "[警告] 設定ファイルの読み込みに失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
        }
    } else {
        Write-Host "[警告] 設定ファイルが見つかりません: $configFile" -ForegroundColor Yellow
        Write-Host "[警告] デフォルトの設定を使用します。" -ForegroundColor Yellow
    }
    
    # ZIPファイルパスを引数から取得
    $zipPath = $args[1]

    if (-not $ZIP_EXE_DIR) {
        Write-Host "[ERROR] 7-Zipのディレクトリが指定されていません。設定ファイル（unzip_to_extract_config.json）でzipExeDirを指定してください。"
        pause
        exit 1
    }
    
    # 7-Zipのディレクトリの存在確認
    $ZIP_EXE_DIR = $ZIP_EXE_DIR.Trim([char]34)
    if (-not (Test-Path $ZIP_EXE_DIR)) {
        Write-Host "[ERROR] 7-Zipのディレクトリが存在しません: $ZIP_EXE_DIR"
        pause
        exit 1
    }
    if (-not (Test-Path $ZIP_EXE_DIR -PathType Container)) {
        Write-Host "[ERROR] 7-Zipのパスはディレクトリではありません: $ZIP_EXE_DIR"
        pause
        exit 1
    }
    
    # 7z.exeのパスを構築
    $ZIP_EXE = Join-Path $ZIP_EXE_DIR "7z.exe"
    
    # 7z.exeの存在確認
    if (-not (Test-Path $ZIP_EXE)) {
        Write-Host "[ERROR] 7z.exeが見つかりません: $ZIP_EXE"
        Write-Host "[ERROR] 指定されたディレクトリ内に7z.exeが存在しません。"
        pause
        exit 1
    }

    if (-not $zipPath) {
        Write-Host "[ERROR] ZIPファイルが指定されていません。"
        pause
        exit 1
    }

    # パスを正規化（引用符を削除）
    $zipPath = $zipPath.Trim([char]34)
    if ($DEST_DIR) {
        $DEST_DIR = $DEST_DIR.Trim([char]34)
    }
    
    # 解凍先フォルダが空の場合、圧縮ファイルと同じ階層に解凍
    if ([string]::IsNullOrWhiteSpace($DEST_DIR)) {
        $DEST_DIR = [System.IO.Path]::GetDirectoryName($zipPath)
        Write-Host "[INFO] 解凍先フォルダが指定されていないため、ZIPファイルと同じ階層に解凍します: $DEST_DIR"
    } else {
        # 解凍先フォルダの存在確認
        if (-not (Test-Path $DEST_DIR)) {
            Write-Host "[ERROR] 解凍先フォルダが存在しません: $DEST_DIR"
            pause
            exit 1
        }
        if (-not (Test-Path $DEST_DIR -PathType Container)) {
            Write-Host "[ERROR] 解凍先パスはフォルダではありません: $DEST_DIR"
            pause
            exit 1
        }
    }

    # 設定情報を表示
    Write-Host ""
    Write-Host "[設定情報]"
    Write-Host "  7-Zipのディレクトリ: $ZIP_EXE_DIR"
    Write-Host "  7-Zipのパス: $ZIP_EXE"
    Write-Host "  解凍先フォルダ: $DEST_DIR"
    Write-Host "  ZIPファイル: $zipPath"
    Write-Host ""

    # ZIPファイルの存在確認
    if (-not (Test-Path $zipPath)) {
        Write-Host "[ERROR] ZIPファイルが見つかりません: $zipPath"
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
        
        # ルートフォルダ名を取得
        $rootFolderName = $rootFolders[0]
        $targetRootFolder = Join-Path $DEST_DIR $rootFolderName
        
        # 同名のフォルダが既に存在する場合、日時を付与
        if (Test-Path $targetRootFolder) {
            Write-Host ""
            Write-Host "[警告] 同名のフォルダが既に存在します: $targetRootFolder" -ForegroundColor Yellow
            Write-Host "[警告] フォルダ名に日時を付与して解凍します。" -ForegroundColor Yellow
            Write-Host ""
            
            $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
            $newRootFolderName = "$rootFolderName`_$timestamp"
            $targetRootFolder = Join-Path $DEST_DIR $newRootFolderName
            Write-Host "[INFO] 新しいフォルダ名: $newRootFolderName"
            
            # 一時的に新しいフォルダ名で解凍するため、ZIPファイル内のパスを置換する必要がある
            # しかし、7-Zipでは直接パスを置換できないため、一時フォルダに解凍してから移動する方法を使用
            $tempExtractDir = Join-Path $env:TEMP "unzip_temp_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
            $extractProcess = Start-Process -FilePath $ZIP_EXE -ArgumentList "x", "`"$zipPath`"", "-o`"$tempExtractDir`"", "-y" -NoNewWindow -Wait -PassThru
            
            if ($extractProcess.ExitCode -eq 0) {
                # 解凍されたルートフォルダをリネームして移動
                $extractedRootFolder = Join-Path $tempExtractDir $rootFolderName
                if (Test-Path $extractedRootFolder) {
                    Move-Item -Path $extractedRootFolder -Destination $targetRootFolder -Force
                    # 移動後にフォルダの作成日時を現在の日時に設定
                    (Get-Item $targetRootFolder).CreationTime = Get-Date
                    (Get-Item $targetRootFolder).LastWriteTime = Get-Date
                }
                # 一時フォルダを削除
                Remove-Item $tempExtractDir -Recurse -Force -ErrorAction SilentlyContinue
            }
        } else {
            $extractProcess = Start-Process -FilePath $ZIP_EXE -ArgumentList "x", "`"$zipPath`"", "-o`"$DEST_DIR`"", "-y" -NoNewWindow -Wait -PassThru
            
            if ($extractProcess.ExitCode -eq 0) {
                # 解凍後にフォルダの作成日時を現在の日時に設定
                if (Test-Path $targetRootFolder) {
                    (Get-Item $targetRootFolder).CreationTime = Get-Date
                    (Get-Item $targetRootFolder).LastWriteTime = Get-Date
                }
            }
        }
    } else {
        Write-Host "[INFO] ルートフォルダなし → ZIP名のフォルダに解凍"
        $targetDir = Join-Path $DEST_DIR $zipName
        
        # 同名のフォルダが既に存在する場合、日時を付与
        if (Test-Path $targetDir) {
            Write-Host ""
            Write-Host "[警告] 同名のフォルダが既に存在します: $targetDir" -ForegroundColor Yellow
            Write-Host "[警告] フォルダ名に日時を付与して解凍します。" -ForegroundColor Yellow
            Write-Host ""
            
            $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
            $newZipName = "$zipName`_$timestamp"
            $targetDir = Join-Path $DEST_DIR $newZipName
            Write-Host "[INFO] 新しいフォルダ名: $newZipName"
        }
        
        $extractProcess = Start-Process -FilePath $ZIP_EXE -ArgumentList "x", "`"$zipPath`"", "-o`"$targetDir`"", "-y" -NoNewWindow -Wait -PassThru
        
        if ($extractProcess.ExitCode -eq 0) {
            # 解凍後にフォルダの作成日時を現在の日時に設定
            if (Test-Path $targetDir) {
                (Get-Item $targetDir).CreationTime = Get-Date
                (Get-Item $targetDir).LastWriteTime = Get-Date
            }
        }
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
    Write-Host ""
    Write-Host "3秒後に自動的に閉じます..."
    Start-Sleep -Seconds 3
}
catch {
    Write-Host "[ERROR] エラーが発生しました: $($_.Exception.Message)"
    Write-Host $_.ScriptStackTrace
    pause
    exit 1
}
