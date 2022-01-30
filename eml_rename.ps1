#-----------------------------------------------------------------------------
# 禁止文字の削除用Function作成
Function Remove-InvalidFileNameChars {
	param(
	  [Parameter(Mandatory=$true,
		Position=0,
		ValueFromPipeline=$true,
		ValueFromPipelineByPropertyName=$true)]
	  [String]$Name
	)

	$invalidChars = [IO.Path]::GetInvalidFileNameChars() -join ''
	$re = "[{0}]" -f [RegEx]::Escape($invalidChars)
	return ($Name -replace $re)
  }

#-----------------------------------------------------------------------------

# 置換処理開始
# フォルダ選択ダイアログ表示
Add-Type -AssemblyName System.Windows.Forms
$FolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog -Property @{ 
    RootFolder = "Desktop"
    Description = 'emlがあるフォルダを選択してください～'
}

# フォルダ選択の有無を判定
if($FolderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK){

	## 選択したファルダパスの格納
	$path = $FolderBrowser.SelectedPath

	## WSHのShellオブジェクトを生成
	$shell=New-Object -Com Shell.Application;

	## 指定されたパス配下のファイルを全て取得（子フォルダ内のファイルも取得するよ）
	$target = Get-ChildItem $path -recurse | Where-Object {-not $_.PSIsContainer}

	## ファイル分だけ繰り返し変数に格納
	$target | ForEach-Object {

		## 拡張子判定（emlを判定）
		if($_.Extension -eq ".eml"){

			## GetDetailsOfに必要なオブジェクト
			$folderobj = $shell.NameSpace($_.DirectoryName)
			$item = $folderobj.ParseName($_.name)

			## ファイルのプロパティ値を取得 [22]は件名
			$title = $folderobj.GetDetailsOf($item,22)

			## 受信日時を取得
			## emlのヘッダーを取得（5行目から100行目）
			## Dateを含む行が取得できればOKなので、環境に応じて変更
			$GetEml = (Get-Content $_.FullName)[5..100]

			## Dateを含む行を取得して、matchesにReTimeプロパティを付与して、日付時刻を切り出し
			foreach ( $Line in $GetEml){
				if( $Line -match "Date: (?<ReTime>.* )"){
					break
				}
			}
			## datetime型に変化後、「20200101_120000」の形式に変換
			$date = ([datetime]$matches.ReTime).ToString("yyyyMMdd_HHmmss")

			if($title -eq ""){
				# ファイル名を受信日時と無題にする
				$newName = $date + "_無題" + $_.Extension

				Rename-Item $_.FullName -NewName $newName

				}else {
				# 禁止文字削除
				$ReName = Remove-InvalidFileNameChars $title
				# ファイル名を受信日時と件名にする
				$newName = $date + "_" + $ReName + $_.Extension

				Rename-Item $_.FullName -NewName $newName
				}
			}
	}
}else{
}
