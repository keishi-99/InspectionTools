# InspectionTools

各種測定機器（デジタルマルチメータ、オシロスコープ、ファンクションジェネレータ、パワーサプライ等）を制御・連携させる Windows デスクトップアプリケーションです。ホットキー操作による効率的な検査・計測ワークフローを提供します。

---

## 機能概要

- **複数測定器の統合制御** — VISA（USB/GPIB）および ADC（USB）対応機器を一元管理
- **ホットキー操作** — 製品ごとにカスタマイズ可能なキーバインドで計測値を自動入力
- **OCR 機能** — 測定器画面をキャプチャし Tesseract OCR で数値を自動読み取り
- **並列接続管理** — 複数デバイスへの非同期並列接続
- **製品別 UI** — 18製品に対応した専用操作画面
- **ヘルプパネル** — ホットキー一覧をアプリ内で表示

---

## 技術スタック

| カテゴリ | 技術 |
|---|---|
| 言語 | C# (.NET 10.0) |
| UI フレームワーク | WPF + Material Design 3 |
| 測定器通信 | VISA COM (VisaComLib 5.15) / USB ADC |
| OCR | Tesseract 5.2.0 |
| キー入力シミュレーション | H.InputSimulator 1.5.0 |

---

## 必要環境

- Windows 10 / 11（64bit）
- .NET 10.0 Desktop Runtime
- VISA ライブラリ（Keysight IO Libraries Suite 等）
  - GPIB / USB VISA 接続の測定器を使用する場合に必要
- 対応測定器（任意）
  - デジタルマルチメータ、オシロスコープ、ファンクションジェネレータ、パワーサプライ等

---

## ディレクトリ構成

```
InspectionTools/
├── InspectionTools.sln
└── InspectionTools/
    ├── App.xaml / App.xaml.cs        # アプリエントリーポイント
    ├── MainWindow.xaml / .cs         # メインウィンドウ
    ├── Common/                        # 共通ユーティリティ
    │   ├── DeviceConnectionHelper.cs  # 並列接続管理
    │   ├── DeviceController.cs        # VISA/ADC 低レベル制御
    │   ├── InstrumentService.cs       # 測定器高レベル操作
    │   ├── InstrumentHelper.cs        # リソース管理
    │   ├── InstClass.cs               # 測定器クラス定義
    │   ├── USBDeviceManager.cs        # USB デバイス管理
    │   ├── HelpManager.cs             # ヘルプシステム
    │   ├── Win32ApiManager.cs         # Win32 API ラッパー
    │   ├── ScreenCaptureWindow.xaml   # スクリーンキャプチャ
    │   └── InstListWindow.xaml        # 測定器設定 UI
    ├── Menu/
    │   └── MainMenuUserControl.xaml   # 製品選択メニュー
    └── Product/                       # 製品別 UI（18製品）
        ├── DFPDXUserControl.xaml
        ├── EL0122UserControl.xaml
        └── ...（以下 16製品）
```

---

## 実行時フォルダ構成

```
実行ディレクトリ/
├── InspectionTools.exe
├── VisaAddress.xml      # 測定器接続設定（要編集）
├── help.json            # ホットキー定義
├── tessdata/            # Tesseract OCR モデル
│   ├── eng.traineddata
│   └── jpn.traineddata
└── ausb.dll             # USB デバイス通信ライブラリ
```

---

## セットアップ

### 1. リポジトリのクローン

```bash
git clone https://github.com/keshi-99/InspectionTools.git
cd InspectionTools
```

### 2. ビルド

Visual Studio 2022 以上でソリューションを開くか、以下を実行します。

```bash
dotnet build InspectionTools.sln
```

### 3. 測定器設定ファイルの編集

`VisaAddress.xml` を環境に合わせて編集します（テンプレート `VisaAddress.template.xml` を参照）。

| SignalType | 接続方式 |
|---|---|
| 1 | ADC（USB直結） |
| 2 | VISA USB |
| 3 | VISA GPIB |
| 4 | VISA（その他） |

```xml
<dsInstList xmlns="http://tempuri.org/dsInstList.xsd">
  <tblInstList>
    <Category>デジタルマルチメータ</Category>
    <Name>DMM-01</Name>
    <VisaAddress>USB0::0x1234::0x5678::SN000001::INSTR</VisaAddress>
    <SignalType>2</SignalType>
  </tblInstList>
</dsInstList>
```

### 4. 実行

```bash
InspectionTools.exe
```

---

## ホットキー設定

`help.json` で製品ごとのホットキーを定義します（テンプレート `help.template.json` を参照）。

```json
{
  "製品名": [
    { "キー": "操作の説明" }
  ]
}
```

---

## 対応製品

| 製品 | 使用測定器 |
|---|---|
| DFPDX | DMM / OSC / FG |
| EL0122 / EL0122FI | DMM |
| EL0137 | DMM / OSC |
| EL1812 | DMM / DCS |
| EL3801 | DMM / OSC / DCS |
| EL4001 | DMM / DCS |
| EL5000 | DMM / FG |
| EL9100〜EL9240 | DMM / OSC / DCS |
| MassFlow | DMM / FG |
| PA14 / PA25 / PA2B / PAF5 | DMM / OSC / DCS / FG |

---

## ライセンス

本プロジェクトのライセンスについては、リポジトリオーナーにお問い合わせください。
