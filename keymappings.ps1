# KeyFunctionDuplicator-CLI.ps1
param (
   [switch]$Start,
   [switch]$Stop,
   [string]$AddMapping,
   [string]$RemoveMapping,
   [switch]$ListMappings,
   [switch]$Help
)

# キーコードマッピング
$keyCodeMap = @{
   "半角/全角" = 0xF3  # VK_OEM_AUTO
   "無変換" = 0x1D     # VK_CONVERT
   "変換" = 0x1C       # VK_NONCONVERT
   "ひらがな" = 0xF2   # VK_OEM_COPY (ひらがな/カタカナ)
   "英数" = 0xF0       # VK_OEM_ATTN (英数)
   "Caps" = 0x14       # VK_CAPITAL
   "左Alt" = 0xA4      # VK_LMENU
   "右Alt" = 0xA5      # VK_RMENU
   "左Ctrl" = 0xA2     # VK_LCONTROL
   "右Ctrl" = 0xA3     # VK_RCONTROL
   "左Shift" = 0xA0    # VK_LSHIFT
   "右Shift" = 0xA1    # VK_RSHIFT
   "Tab" = 0x09        # VK_TAB
   "Esc" = 0x1B        # VK_ESCAPE
}

# 設定ファイルパス
$settingsPath = "$env:APPDATA\KeyFunctionDuplicator\settings.json"
$settingsDir = "$env:APPDATA\KeyFunctionDuplicator"

# 設定フォルダが存在しなければ作成
if (-not (Test-Path $settingsDir)) {
   New-Item -Path $settingsDir -ItemType Directory -Force | Out-Null
}

# デフォルト設定
$defaultSettings = @{
   Mappings = @()
}

# 設定の読み込み
function Get-KeyMappings {
   if (Test-Path $settingsPath) {
       $settings = Get-Content -Path $settingsPath -Encoding UTF8 | ConvertFrom-Json
       return $settings
   }
   
   return $defaultSettings
}

# 設定の保存
function Save-KeyMappings($settings) {
   $settings | ConvertTo-Json | Out-File -FilePath $settingsPath -Encoding UTF8
   Write-Host "設定を保存しました"
}

# マッピングの一覧表示
function Show-Mappings {
   $settings = Get-KeyMappings
   
   if ($settings.Mappings.Count -eq 0) {
       Write-Host "現在のマッピングはありません"
       return
   }
   
   Write-Host "現在のキーマッピング:"
   for ($i = 0; $i -lt $settings.Mappings.Count; $i++) {
       $parts = $settings.Mappings[$i] -split ":"
       $source = $parts[0].Trim()
       $target = $parts[1].Trim()
       Write-Host "[$i] $source → $target"
   }
}

# マッピングの追加
function Add-KeyMapping($mappingString) {
   $settings = Get-KeyMappings
   
   # マッピング文字列をパース (例: "半角/全角:Caps")
   $parts = $mappingString -split ":"
   
   if ($parts.Count -ne 2) {
       Write-Host "無効なマッピング形式です。「元のキー:複製先キー」の形式で指定してください。"
       return
   }
   
   $sourceKey = $parts[0].Trim()
   $targetKey = $parts[1].Trim()
   
   # キーコードマップに存在するか確認
   if (-not $keyCodeMap.ContainsKey($sourceKey)) {
       Write-Host "元のキー '$sourceKey' は無効です。有効なキー: $($keyCodeMap.Keys -join ", ")"
       return
   }
   
   if (-not $keyCodeMap.ContainsKey($targetKey)) {
       Write-Host "複製先キー '$targetKey' は無効です。有効なキー: $($keyCodeMap.Keys -join ", ")"
       return
   }
   
   # 同じ複製先キーが既にあるか確認
   foreach ($mapping in $settings.Mappings) {
       $existingParts = $mapping -split ":"
       if ($existingParts[1].Trim() -eq $targetKey) {
           Write-Host "複製先キー '$targetKey' は既に使用されています。"
           return
       }
   }
   
   # マッピングを追加
   $newMapping = "$sourceKey:$targetKey"
   $settings.Mappings += $newMapping
   
   # 設定を保存
   Save-KeyMappings $settings
   Write-Host "マッピングを追加しました: $sourceKey → $targetKey"
}

# マッピングの削除
function Remove-KeyMapping($index) {
   $settings = Get-KeyMappings
   
   if ($index -lt 0 -or $index -ge $settings.Mappings.Count) {
       Write-Host "指定されたインデックスは範囲外です。`n現在のマッピング:"
       Show-Mappings
       return
   }
   
   $removedMapping = $settings.Mappings[$index]
   $parts = $removedMapping -split ":"
   $source = $parts[0].Trim()
   $target = $parts[1].Trim()
   
   $settings.Mappings = $settings.Mappings | Where-Object { $_ -ne $removedMapping }
   
   # 設定を保存
   Save-KeyMappings $settings
   Write-Host "マッピングを削除しました: $source → $target"
}

# キー監視の開始
function Start-KeyMonitoring {
   $settings = Get-KeyMappings
   
   if ($settings.Mappings.Count -eq 0) {
       Write-Host "マッピングが設定されていません。`nマッピングを追加するには: .\KeyFunctionDuplicator-CLI.ps1 -AddMapping '半角/全角:Caps'"
       return
   }
   
   # 既存のジョブの確認
   $existingJob = Get-Job -Name "KeyDuplicator" -ErrorAction SilentlyContinue
   if ($existingJob) {
       Write-Host "キー監視は既に実行中です。"
       return
   }
   
   # キー監視スクリプトブロック
   $keyMonitoringScript = {
       param($mappingsArray, $keyCodeMapSerialized)
       
       # キーコードマップをデシリアライズ
       $keyCodeMap = [System.Management.Automation.PSSerializer]::Deserialize($keyCodeMapSerialized)
       
       Add-Type @"
       using System;
       using System.Runtime.InteropServices;
       
       public class KeyboardSimulator {
           [DllImport("user32.dll")]
           public static extern short GetAsyncKeyState(int vKey);
           
           [DllImport("user32.dll")]
           public static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);
           
           public const uint KEYEVENTF_KEYDOWN = 0x0000;
           public const uint KEYEVENTF_KEYUP = 0x0002;
           
           public static void SimulateKeyPress(byte keyCode) {
               keybd_event(keyCode, 0, KEYEVENTF_KEYDOWN, UIntPtr.Zero);
               keybd_event(keyCode, 0, KEYEVENTF_KEYUP, UIntPtr.Zero);
           }
       }
"@
       
       # マッピングをパース
       $parsedMappings = @()
       foreach ($mapping in $mappingsArray) {
           $parts = $mapping -split ":"
           if ($parts.Count -eq 2) {
               $source = $parts[0].Trim()
               $target = $parts[1].Trim()
               
               if ($keyCodeMap.ContainsKey($source) -and $keyCodeMap.ContainsKey($target)) {
                   $parsedMappings += @{
                       SourceKey = $keyCodeMap[$source]
                       TargetKey = $keyCodeMap[$target]
                       SourceName = $source
                       TargetName = $target
                   }
               }
           }
       }
       
       # キーの前回の状態を記録
       $keyStates = @{}
       
       Write-Output "キー機能複製を開始しました。監視中のマッピング:"
       foreach ($mapping in $parsedMappings) {
           Write-Output "・$($mapping.SourceName) → $($mapping.TargetName)"
       }
       
       while ($true) {
           foreach ($mapping in $parsedMappings) {
               # 元のキーの状態を取得
               $sourceState = [KeyboardSimulator]::GetAsyncKeyState($mapping.SourceKey)
               $targetState = [KeyboardSimulator]::GetAsyncKeyState($mapping.TargetKey)
               
               # 複製先キーが押されたら元のキーの機能を発動
               if (($targetState -band 0x8000) -eq 0x8000) {
                   # 前回の状態を確認して、ループを避ける
                   if (-not $keyStates.ContainsKey($mapping.TargetKey) -or $keyStates[$mapping.TargetKey] -eq 0) {
                       # 複製先キーが押されたので、元のキーの機能を実行
                       [KeyboardSimulator]::SimulateKeyPress([byte]$mapping.SourceKey)
                   }
                   
                   # 状態を更新
                   $keyStates[$mapping.TargetKey] = 1
               } else {
                   # キーが離されたら状態をリセット
                   $keyStates[$mapping.TargetKey] = 0
               }
           }
           
           # 短い遅延を入れて CPU 使用率を下げる
           Start-Sleep -Milliseconds 30
       }
   }
   
   # キーコードマップをシリアライズ
   $keyCodeMapSerialized = [System.Management.Automation.PSSerializer]::Serialize($keyCodeMap)
   
   # バックグラウンドジョブとして実行
   Start-Job -Name "KeyDuplicator" -ScriptBlock $keyMonitoringScript -ArgumentList $settings.Mappings, $keyCodeMapSerialized | Out-Null
   
   Write-Host "キー機能複製を開始しました。"
   Write-Host "監視を停止するには: .\KeyFunctionDuplicator-CLI.ps1 -Stop"
}

# キー監視の停止
function Stop-KeyMonitoring {
   $job = Get-Job -Name "KeyDuplicator" -ErrorAction SilentlyContinue
   
   if ($job) {
       Stop-Job -Job $job
       Remove-Job -Job $job -Force
       Write-Host "キー機能複製を停止しました。"
   } else {
       Write-Host "実行中のキー監視ジョブはありません。"
   }
}

# ヘルプの表示
function Show-Help {
   Write-Host @"
キー機能複製ツール (コマンドライン版)

使用方法:
 .\KeyFunctionDuplicator-CLI.ps1 [オプション]

オプション:
 -Help            : このヘルプを表示します
 -ListMappings    : 現在のキーマッピングを一覧表示します
 -AddMapping      : 新しいマッピングを追加します (形式: '元のキー:複製先キー')
                    例: -AddMapping '半角/全角:Caps'
 -RemoveMapping   : 指定したインデックスのマッピングを削除します
                    例: -RemoveMapping 0
 -Start           : キー機能複製を開始します
 -Stop            : キー機能複製を停止します

有効なキー名:
 $(($keyCodeMap.Keys | Sort-Object) -join ", ")

例:
 .\KeyFunctionDuplicator-CLI.ps1 -AddMapping '半角/全角:Caps'
 .\KeyFunctionDuplicator-CLI.ps1 -Start
 .\KeyFunctionDuplicator-CLI.ps1 -ListMappings
 .\KeyFunctionDuplicator-CLI.ps1 -Stop
"@
}

# スクリプトのメイン処理
if ($Help) {
   Show-Help
} elseif ($ListMappings) {
   Show-Mappings
} elseif ($AddMapping) {
   Add-KeyMapping $AddMapping
} elseif ($RemoveMapping) {
   Remove-KeyMapping $RemoveMapping
} elseif ($Start) {
   Start-KeyMonitoring
} elseif ($Stop) {
   Stop-KeyMonitoring
} else {
   Show-Help
}