# Scripts

PowerShell -ExecutionPolicy RemoteSigned .\GetExcelValue.ps1



実行ポリシー	署名あり	署名なし/ローカル	署名なし/非ローカル	説明
Restricted	x	x	x	すべてのスクリプトの実行を制限 (初期設定)
AllSigned	o	x	x	署名のあるスクリプトのみ実行可能
RemoteSigned	o	o	x	ローカル上のスクリプトと非ローカル上の署名のあるスクリプトのみ実行可能
Unrestricted	o	o	△	すべてのスクリプトが実行可能だが非ローカル上のスクリプトは実行時に許可が必要
Bypass	o	o	o	すべてのスクリプトが実行可能