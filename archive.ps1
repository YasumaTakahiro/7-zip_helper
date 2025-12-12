# エンコーディング設定（UTF-8）
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

# エラー時に停止
$ErrorActionPreference = "Stop"

try {
    # スクリプトのディレクトリを取得
    # バッチファイル経由の場合は引数の最初の要素がスクリプトディレクトリ
    $scriptDir = $null
    
    # まず、環境変数から取得を試みる（archive.batで設定される）
    $envScriptDir = $env:ARCHIVE_SCRIPT_DIR
    if ($envScriptDir -and -not [string]::IsNullOrWhiteSpace($envScriptDir)) {
        $potentialScriptDir = $envScriptDir.TrimEnd('\')
        if (Test-Path $potentialScriptDir -PathType Container) {
            $configFileCheck = Join-Path $potentialScriptDir "archive_config.jsonc"
            if (Test-Path $configFileCheck) {
                $scriptDir = $potentialScriptDir
            }
        }
    }
    
    # 環境変数から取得できなかった場合、引数から取得を試みる（バッチファイル経由の場合）
    if (-not $scriptDir -and $args.Count -gt 0 -and -not [string]::IsNullOrWhiteSpace($args[0])) {
        $potentialScriptDir = $args[0].TrimEnd('\')
        # スクリプトディレクトリの可能性があるかチェック（ディレクトリが存在し、archive_config.jsoncが存在する）
        if (Test-Path $potentialScriptDir -PathType Container) {
            $configFileCheck = Join-Path $potentialScriptDir "archive_config.jsonc"
            if (Test-Path $configFileCheck) {
                $scriptDir = $potentialScriptDir
            }
        }
    }
    
    # 引数から取得できなかった場合、他の方法を試す
    if (-not $scriptDir) {
        # $PSScriptRootが利用可能な場合はそれを使用
        if ($PSScriptRoot) {
            $scriptDir = $PSScriptRoot
        } elseif ($args.Count -gt 0) {
            # archive.batから渡されたスクリプトディレクトリを再試行（存在チェックなし）
            $potentialScriptDir = $args[0].TrimEnd('\')
            if (-not [string]::IsNullOrWhiteSpace($potentialScriptDir)) {
                $scriptDir = $potentialScriptDir
            }
        }
        
        if (-not $scriptDir) {
            # 最後のフォールバック: 現在のディレクトリを使用
            $scriptDir = (Get-Location).Path
        }
    }
    
    # 末尾のバックスラッシュを削除
    $scriptDir = $scriptDir.TrimEnd('\')
    
    # JSON設定ファイルのパス（.jsoncまたは.json）
    $configFile = Join-Path $scriptDir "archive_config.jsonc"
    if (-not (Test-Path $configFile)) {
        $configFile = Join-Path $scriptDir "archive_config.json"
    }
    
    # JSON設定ファイルから設定を読み込む
    $ZIP_EXE_DIR = $null
    $COMPRESS_DEST_DIR = $null
    $PASSWORD = $null
    $COMPRESS_FORMAT = "zip"  # デフォルト: zip
    $CLOSE_DELAY_MS = 3000  # デフォルト: 3秒（3000ミリ秒）
    
    if (Test-Path $configFile) {
        try {
            # JSONファイルを読み込み、コメントを削除してからパース
            $jsonContent = Get-Content $configFile -Raw -Encoding UTF8
            
            # より安全なコメント削除処理
            # ブロックコメント（/* */）を削除
            while ($jsonContent -match '/\*[\s\S]*?\*/') {
                $jsonContent = $jsonContent -replace '/\*[\s\S]*?\*/', ''
            }
            
            # 行コメント（//）を削除（文字列内を除く）
            $lines = $jsonContent -split "`r?`n"
            $cleanedLines = @()
            foreach ($line in $lines) {
                $inString = $false
                $escaped = $false
                $result = ""
                for ($i = 0; $i -lt $line.Length; $i++) {
                    $char = $line[$i]
                    
                    if ($escaped) {
                        $result += $char
                        $escaped = $false
                        continue
                    }
                    
                    if ($char -eq '\') {
                        $escaped = $true
                        $result += $char
                        continue
                    }
                    
                    if ($char -eq '"') {
                        $inString = -not $inString
                        $result += $char
                        continue
                    }
                    
                    if (-not $inString -and $char -eq '/' -and $i + 1 -lt $line.Length -and $line[$i + 1] -eq '/') {
                        # 行コメントの開始
                        break
                    }
                    
                    $result += $char
                }
                # 行末の空白を削除
                $result = $result.TrimEnd()
                if ($result.Length -gt 0) {
                    $cleanedLines += $result
                }
            }
            $jsonContent = $cleanedLines -join "`n"
            
            # 末尾のカンマを削除（JSONの最後のプロパティの後のカンマ）
            $jsonContent = $jsonContent -replace ',\s*}', '}'
            $jsonContent = $jsonContent -replace ',\s*]', ']'
            
            $config = $jsonContent | ConvertFrom-Json
            if ($config.zipExeDir) {
                $ZIP_EXE_DIR = $config.zipExeDir
                # 環境変数を展開
                $ZIP_EXE_DIR = [System.Environment]::ExpandEnvironmentVariables($ZIP_EXE_DIR)
            }
            if ($config.compressDestDir) {
                $COMPRESS_DEST_DIR = $config.compressDestDir
                # 環境変数を展開
                $COMPRESS_DEST_DIR = [System.Environment]::ExpandEnvironmentVariables($COMPRESS_DEST_DIR)
            }
            if ($config.password) {
                $PASSWORD = $config.password
            }
            if ($config.compressFormat) {
                $COMPRESS_FORMAT = $config.compressFormat
            }
            if ($config.closeDelayMs) {
                $CLOSE_DELAY_MS = $config.closeDelayMs
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
    
    # 圧縮対象のパスとパスワード無効化フラグを引数から取得
    # 引数の順序に関係なく処理できるようにする
    $targetPath = $null
    $noPassword = $false
    
    # パスワード無効化フラグを検出
    $passwordFlags = @("--no-password", "-nopass", "/nopass")
    $filteredArgs = @()
    
    foreach ($arg in $args) {
        # 引数を文字列として正規化（トリム）
        $argStr = $arg.ToString().Trim()
        
        # パスワードフラグかどうかをチェック（大文字小文字を区別しない）
        $isPasswordFlag = $false
        foreach ($flag in $passwordFlags) {
            if ($argStr -eq $flag -or $argStr -ieq $flag) {
                $isPasswordFlag = $true
                $noPassword = $true
                break
            }
        }
        
        if (-not $isPasswordFlag) {
            $filteredArgs += $arg
        }
    }
    
    # スクリプトディレクトリ（$args[0]または$filteredArgs[0]）を除いた最初の引数が圧縮対象のパス
    # $filteredArgs[0]がスクリプトディレクトリ、$filteredArgs[1]が圧縮対象のパス
    if ($filteredArgs.Count -ge 2) {
        $targetPath = $filteredArgs[1]
    } elseif ($filteredArgs.Count -eq 1) {
        # スクリプトディレクトリのみの場合（引数なし）
        $targetPath = $null
    }

    if (-not $ZIP_EXE_DIR) {
        Write-Host "[ERROR] 7-Zipのディレクトリが指定されていません。設定ファイル（archive_config.jsonc）でzipExeDirを指定してください。"
        pause
        exit 1
    }
    
    if (-not $targetPath) {
        Write-Host "[ERROR] 圧縮対象のファイルまたはフォルダが指定されていません。"
        pause
        exit 1
    }

    # パスを正規化（引用符を削除）
    $targetPath = $targetPath.Trim([char]34)

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

    # パスを正規化（引用符を削除）
    $targetPath = $targetPath.Trim([char]34)
    
    # 圧縮対象の存在確認（-LiteralPathを使用して特殊文字を正しく処理）
    if (-not (Test-Path -LiteralPath $targetPath)) {
        Write-Host "[ERROR] 圧縮対象が見つかりません: $targetPath"
        pause
        exit 1
    }

    # 設定情報を表示
    Write-Host ""
    Write-Host "[設定情報]"
    Write-Host "  7-Zipのディレクトリ: $ZIP_EXE_DIR"
    Write-Host "  7-Zipのパス: $ZIP_EXE"
    Write-Host "  圧縮対象: $targetPath"
    if ($COMPRESS_DEST_DIR) {
        Write-Host "  圧縮先フォルダ: $COMPRESS_DEST_DIR"
    } else {
        Write-Host "  圧縮先フォルダ: 圧縮対象と同じ階層"
    }
    Write-Host "  圧縮形式: $COMPRESS_FORMAT"
    if ($PASSWORD) {
        Write-Host "  パスワード: 設定済み（表示されません）"
    } else {
        Write-Host "  パスワード: なし"
    }
    Write-Host ""

    # 圧縮対象がファイルかフォルダかを判定（-LiteralPathを使用して特殊文字を正しく処理）
    $isFile = Test-Path -LiteralPath $targetPath -PathType Leaf
    
    $compressTarget = $targetPath
    $tempFolder = $null
    
    if ($isFile) {
        # ファイルの場合、同名のフォルダを作成してその中にファイルを配置
        $fileName = [System.IO.Path]::GetFileNameWithoutExtension($targetPath)
        $fileDir = [System.IO.Path]::GetDirectoryName($targetPath)
        $tempFolder = Join-Path $fileDir $fileName
        
        Write-Host "[INFO] ファイルが指定されました。同名のフォルダを作成します: $tempFolder"
        
        # 一時フォルダを作成
        if (Test-Path -LiteralPath $tempFolder) {
            Write-Host "[警告] 同名のフォルダが既に存在します: $tempFolder" -ForegroundColor Yellow
            $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
            $tempFolder = Join-Path $fileDir "$fileName`_$timestamp"
            Write-Host "[INFO] 新しいフォルダ名: $tempFolder"
        }
        
        New-Item -ItemType Directory -Path $tempFolder -Force | Out-Null
        
        # ファイルを一時フォルダにコピー（-LiteralPathを使用）
        $destFile = Join-Path $tempFolder ([System.IO.Path]::GetFileName($targetPath))
        Copy-Item -LiteralPath $targetPath -Destination $destFile -Force
        
        $compressTarget = $tempFolder
        Write-Host "[INFO] ファイルを一時フォルダに配置しました: $destFile"
    }

    # ZIPファイル名を決定（圧縮対象と同じ名前）
    $targetName = [System.IO.Path]::GetFileNameWithoutExtension($compressTarget)
    
    # 圧縮先フォルダが指定されている場合はそれを使用、空の場合は圧縮対象と同じ階層
    if ([string]::IsNullOrWhiteSpace($COMPRESS_DEST_DIR)) {
        $targetDir = [System.IO.Path]::GetDirectoryName($compressTarget)
        Write-Host "[INFO] 圧縮先フォルダが指定されていないため、圧縮対象と同じ階層に圧縮します: $targetDir"
    } else {
        $COMPRESS_DEST_DIR = $COMPRESS_DEST_DIR.Trim([char]34)
        # 圧縮先フォルダの存在確認
        if (-not (Test-Path -LiteralPath $COMPRESS_DEST_DIR)) {
            Write-Host "[ERROR] 圧縮先フォルダが存在しません: $COMPRESS_DEST_DIR"
            pause
            exit 1
        }
        if (-not (Test-Path -LiteralPath $COMPRESS_DEST_DIR -PathType Container)) {
            Write-Host "[ERROR] 圧縮先パスはフォルダではありません: $COMPRESS_DEST_DIR"
            pause
            exit 1
        }
        $targetDir = $COMPRESS_DEST_DIR
    }
    
    # 圧縮形式に応じた拡張子を決定
    $archiveExtension = switch ($COMPRESS_FORMAT.ToLower()) {
        "7z" { "7z" }
        "zip" { "zip" }
        "gzip" { "gz" }
        "bzip2" { "bz2" }
        "tar" { "tar" }
        "wim" { "wim" }
        "xz" { "xz" }
        default { "zip" }  # デフォルトはzip
    }
    
    $archiveFileName = "$targetName.$archiveExtension"
    $archiveFilePath = Join-Path $targetDir $archiveFileName
    
    # 同名のアーカイブファイルが存在する場合、日時を付与（-LiteralPathを使用）
    if (Test-Path -LiteralPath $archiveFilePath) {
        Write-Host ""
        Write-Host "[警告] 同名のアーカイブファイルが既に存在します: $archiveFilePath" -ForegroundColor Yellow
        Write-Host "[警告] アーカイブファイル名に日時を付与します。" -ForegroundColor Yellow
        Write-Host ""
        
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $archiveFileName = "$targetName`_$timestamp.$archiveExtension"
        $archiveFilePath = Join-Path $targetDir $archiveFileName
        Write-Host "[INFO] 新しいアーカイブファイル名: $archiveFileName"
    }

    Write-Host "[INFO] 圧縮を開始します..."
    Write-Host "[INFO] 圧縮形式: $COMPRESS_FORMAT"
    Write-Host "[INFO] アーカイブファイル: $archiveFilePath"
    
    # 7-Zipで圧縮
    $arguments = @("a", "-t$COMPRESS_FORMAT", "`"$archiveFilePath`"", "`"$compressTarget`"")
    
    # パスワードが設定されている場合、パスワードオプションを追加（--no-passwordフラグが指定されていない場合のみ）
    if ($PASSWORD -and -not $noPassword) {
        $arguments += "-p$PASSWORD"
        Write-Host "[INFO] パスワード付きで圧縮します。"
    } elseif ($noPassword) {
        Write-Host "[INFO] パスワード無効化フラグが指定されました。パスワードなしで圧縮します。"
    }
    
    # 7-Zipの出力を適切なエンコーディングで読み込むため、一時ファイルにリダイレクト
    $tempOutputFile = [System.IO.Path]::GetTempFileName()
    $tempErrorFile = [System.IO.Path]::GetTempFileName()
    $compressProcess = Start-Process -FilePath $ZIP_EXE -ArgumentList $arguments -NoNewWindow -Wait -PassThru -RedirectStandardOutput $tempOutputFile -RedirectStandardError $tempErrorFile
    
    # 7-Zipの出力をシステムのデフォルトコードページ（通常はShift-JIS）で読み込んで表示
    if (Test-Path $tempOutputFile) {
        $outputContent = Get-Content $tempOutputFile -Raw -Encoding Default
        if ($outputContent) {
            Write-Host $outputContent
        }
        Remove-Item $tempOutputFile -ErrorAction SilentlyContinue
    }
    
    # エラー出力も表示
    if (Test-Path $tempErrorFile) {
        $errorContent = Get-Content $tempErrorFile -Raw -Encoding Default
        if ($errorContent) {
            Write-Host $errorContent
        }
        Remove-Item $tempErrorFile -ErrorAction SilentlyContinue
    }
    
    if ($compressProcess.ExitCode -ne 0) {
        Write-Host "[ERROR] 圧縮に失敗しました。終了コード: $($compressProcess.ExitCode)"
        
        # 一時フォルダを削除（-LiteralPathを使用）
        if ($tempFolder -and (Test-Path -LiteralPath $tempFolder)) {
            Remove-Item -LiteralPath $tempFolder -Recurse -Force -ErrorAction SilentlyContinue
        }
        
        pause
        exit 1
    }

    # 一時フォルダを削除（-LiteralPathを使用）
    if ($tempFolder -and (Test-Path -LiteralPath $tempFolder)) {
        Write-Host "[INFO] 一時フォルダを削除します: $tempFolder"
        Remove-Item -LiteralPath $tempFolder -Recurse -Force -ErrorAction SilentlyContinue
    }

    Write-Host "圧縮完了。"
    Write-Host ""
    if ($CLOSE_DELAY_MS -gt 0) {
        $seconds = [math]::Round($CLOSE_DELAY_MS / 1000, 1)
        Write-Host ("{0}秒後に自動的に閉じます..." -f $seconds)
        Start-Sleep -Milliseconds $CLOSE_DELAY_MS
    } else {
        Write-Host "自動的に閉じます..."
    }
}
catch {
    Write-Host "[ERROR] エラーが発生しました: $($_.Exception.Message)"
    Write-Host $_.ScriptStackTrace
    
    # 一時フォルダを削除（-LiteralPathを使用）
    if ($tempFolder -and (Test-Path -LiteralPath $tempFolder)) {
        Remove-Item -LiteralPath $tempFolder -Recurse -Force -ErrorAction SilentlyContinue
    }
    
    pause
    exit 1
}

