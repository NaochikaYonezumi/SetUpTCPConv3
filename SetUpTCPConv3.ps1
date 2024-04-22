Add-Type -AssemblyName System.Windows.Forms

# ファイル選択ダイアログの関数
Function Select-FileDialog {
    param([string]$Title,[string]$InitialDirectory,[string]$Filter)
    $fileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $fileDialog.InitialDirectory = $InitialDirectory
    $fileDialog.Filter = $Filter
    $fileDialog.Title = $Title
    $result = $fileDialog.ShowDialog()
    If ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        $fileDialog.FileName
    } Else {
        $null
    }
}




# フォルダ選択ダイアログの関数
Function Select-FolderDialog {
    param([string]$Description,[string]$RootFolder)
    $folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderDialog.Description = $Description
    $folderDialog.RootFolder = if ([string]::IsNullOrWhiteSpace($RootFolder)) { [System.Environment+SpecialFolder]::Desktop } else { $RootFolder }
    $result = $folderDialog.ShowDialog()
    If ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        $folderDialog.SelectedPath
    } Else {
        $null
    }
}

# ファイルとフォルダを選択
$jsonTemplatePath = Select-FileDialog -Title "JSONテンプレートファイルを選択" -Filter "JSON files (*.json)|*.json"
$csvPath = Select-FileDialog -Title "CSVファイルを選択" -Filter "CSV files (*.csv)|*.csv"
$outputFolder = Select-FolderDialog -Description "出力フォルダを選択"


# JSON テンプレートを読み込む
$template = Get-Content $jsonTemplatePath | ConvertFrom-Json

# CSVファイルからデータを読み込む
$csvData = Import-Csv $csvPath

# 各プリンターのデータに対して処理を行う
foreach ($row in $csvData) {
$printerJson = $template
    $printerJson.network.eth0.ip = $row.ip
    $printerJson.network.eth0.netmask = $row.netmask
    $printerJson.network.eth0.gateway = $row.gateway
    $printerJson.network.eth0.dns = $row.dns
    $printerJson.network.eth0.dns_ = $row.dns_
    $printerJson.network.eth0.wins = $row.wins
    $printerJson.network.eth0.wins_ = $row.wins_

    # JSONファイルを保存
    $jsonOutput = $printerJson | ConvertTo-Json -Depth 10
    $outputPath = Join-Path $outputFolder ($row.printername + ".json")
    $jsonOutput | Out-File $outputPath
}

Write-Host "すべてのJSONファイルが生成されました。"
