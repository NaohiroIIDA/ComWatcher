# ComWatcher (日本語)

ComWatcher は、USB シリアル（COM）ポートの追加・削除を監視し、トレイ通知で知らせる軽量な Windows トレイ常駐アプリ（WPF, .NET 8）です。WMI を用いた取得と `SerialPort` によるフォールバックで、安定してポート情報を表示します。

## 特長
- 起動時はウィンドウを表示せずトレイ常駐
- 1秒間隔で監視し、追加/削除の差分を検出
- USB シリアルに限定して一覧表示（デバイス名・VID/PID を表示）
- 追加時にトレイバルーン通知（削除通知は任意で有効化可能）
- WMI 取得＋`SerialPort.GetPortNames()` フォールバックで堅牢

## 使い方
1. `ComWatcher.exe` を起動（トレイに常駐）
2. トレイアイコン 左クリック: ウィンドウ表示 / 右クリック: メニュー（表示 / 終了）
3. 一覧は USB シリアル COM ポートのみを表示し、最後に挿したポートが一番上に並びます

## ビルド / 発行
- 必要環境: .NET 8 SDK, Windows 10/11
- 単一ファイル・自己完結・英語リソースのみで発行（EXE 単体配布向け）:

```powershell
# プロジェクトディレクトリで実行
 dotnet publish ComWatcher.csproj -c Release -r win-x64 -p:PublishSingleFile=true -p:SelfContained=true -p:SatelliteResourceLanguages=en
```

生成物は次のディレクトリに出力されます:
```
./bin/Release/net8.0-windows/win-x64/publish/
```

注記:
- 完全な単一 EXE を目指す場合、WPF の衛星アセンブリ（多言語リソース）を含めない設定にしています（`SatelliteResourceLanguages=en`）。これにより EXE 単体での配布がしやすくなります。
- 多言語リソースが必要な場合は、`publish` フォルダを丸ごと配布してください。

## ライセンス
未設定（必要に応じて追加してください）。
