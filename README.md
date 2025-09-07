# ReferenceFinder 使用方法

## 概要

ReferenceFinder.exe は、指定した .sln ファイルを解析し、全プロジェクトから public const フィールドを収集し、それぞれの参照元（型／メンバー／行）をコンソールおよび CSV（Shift_JIS）で出力するツールです。

## ビルド方法

1. .NET 8.0 SDK をインストールしてください。
2. コマンドプロンプトで以下を実行します。

```
dotnet restore
dotnet build -c Release
```

ビルド後、`bin\Release\net8.0\ReferenceFinder.exe` が生成されます。

## 使い方

コマンドプロンプトで以下のように実行します。

```
ReferenceFinder.exe <ソリューションファイルパス>
```

例:

```
ReferenceFinder.exe C:\work\MyApp\MyApp.sln
```

### dotnet run で実行する場合

```
dotnet run --project ReferenceFinder.csproj -- <ソリューションファイルパス>
```

## 出力

- 解析結果はコンソールに表示され、同時に実行フォルダに `ReferenceFinderResult_yyyyMMddHHmmss.csv`（Shift_JIS）として出力されます。

## CSV の内容

- 定数宣言の名前空間、クラス、定数名
- 参照元のクラス名、メンバー名、行番号、該当コード行、ファイルパス

## 注意事項

- .sln ファイルが存在しない場合や引数が不足している場合はエラーとなります。
- プロジェクトの復元（`dotnet restore`）を事前に行ってください。
- CSV の文字コードは Shift_JIS です。UTF-8 で出力したい場合はソースコードを
