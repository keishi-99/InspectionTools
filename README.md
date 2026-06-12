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

## 測定器設定ファイルの編集

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

---
