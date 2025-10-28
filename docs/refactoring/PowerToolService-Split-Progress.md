# PowerToolService.cs 分割作業 進捗記録

## プロジェクト概要
- **目的**: PowerToolService.cs (299KB, 約9,225行) を機能別に分割し、保守性を向上
- **戦略**: 既存フォルダ構成パターンに準拠した4段階アプローチ
- **期間**: 1週間（2025-10-28 開始）

---

## 📁 目標フォルダ構成

```
Services/Core/
├── Alignment/               # 既存
├── Shape/                   # 既存
├── Text/                    # 既存
├── PowerTool/               # ★Phase 2（コア機能）
│   ├── PowerToolService.cs
│   └── PowerToolServiceHelper.cs
├── Image/                   # ★Phase 1-1 ✅
│   └── ImageCompressionService.cs
├── Selection/               # ★Phase 1-2
│   └── ShapeSelectionService.cs
├── Table/                   # ★Phase 3-1
│   ├── TableConversionService.cs
│   └── MatrixOperationService.cs
└── BuiltInShape/            # ★Phase 3-2
    └── BuiltInShapeService.cs
```

---

## 📊 Phase 1: 完全独立機能の分離

### ✅ Phase 1-1: ImageCompressionService 作成完了
- **日時**: 2025-10-28 15:51 JST
- **コミット**: `3753305798a347f55b0a8854cd3e77ff7e301038`
- **ファイル**: `Services/Core/Image/ImageCompressionService.cs`
- **サイズ**: 22.7KB (約650行)
- **含まれる機能**:
  - `CompressImages()` - 画像圧縮メイン処理
  - `ExtractVisibleImageData()` - 画像データ抽出
  - `ExecuteFinalCompressionInternal()` - 最終圧縮
  - `ApplyJpegCompression()`, `ApplyPngReducedCompression()`, `ApplyPngLosslessCompression()` - 圧縮ロジック
  - `ReplaceImageInShape()` - 画像置換
  - `CleanupTempFilesInternal()` - 一時ファイル削除
  - `FormatFileSize()` - サイズフォーマット
- **依存関係**:
  - ImageMagick (外部ライブラリ)
  - ImageCompressionDialog
  - IApplicationProvider (DI)
- **状態**: ✅ 完了

### ⏳ Phase 1-2: ShapeSelectionService 作成
- **予定日**: 2025-10-28
- **ファイル**: `Services/Core/Selection/ShapeSelectionService.cs`
- **予想サイズ**: 約200-300行
- **含まれる予定機能**:
  - `SelectSameColorShapes()` - 同色図形選択
  - `SelectSameSizeShapes()` - 同サイズ図形選択
  - `TransparencyUpToggle()` - 透過率Up
  - `TransparencyDownToggle()` - 透過率Down
- **状態**: 🔄 作業中

---

## 📋 Phase 2: 共通ヘルパーの整理

### ⏳ Phase 2-1: PowerToolServiceHelper 作成
- **予定日**: 2025-10-29
- **ファイル**: `Services/Core/PowerTool/PowerToolServiceHelper.cs`
- **予想サイズ**: 約500-600行
- **含まれる予定機能**:
  - 図形選択・検証の共通処理
  - グリッド検出ロジック
  - 図形判定ヘルパー
  - GridInfoクラス
- **状態**: ⏳ 未着手

### ⏳ Phase 2-2: PowerToolService コア機能残存
- **予定日**: 2025-10-29
- **ファイル**: `Services/Core/PowerTool/PowerToolService.cs`
- **予想サイズ**: 約1,500-2,000行
- **残す機能**:
  - テキスト合成
  - 線操作
  - 位置交換
  - Excel貼り付け
  - フォント一括統一
- **状態**: ⏳ 未着手

---

## 📋 Phase 3: 大規模機能グループの分離

### ⏳ Phase 3-1: Table フォルダ作成
- **予定日**: 2025-10-30
- **ファイル**: 
  - `Services/Core/Table/TableConversionService.cs` (約600-700行)
  - `Services/Core/Table/MatrixOperationService.cs` (約1,000-1,200行)
- **状態**: ⏳ 未着手

### ⏳ Phase 3-2: BuiltInShapeService 作成
- **予定日**: 2025-10-31
- **ファイル**: `Services/Core/BuiltInShape/BuiltInShapeService.cs`
- **予想サイズ**: 約700-800行
- **状態**: ⏳ 未着手

---

## 📅 スケジュール

| Date | Phase | 作業内容 | 状態 |
|------|-------|---------|------|
| 2025-10-28 | Phase 1-1 | ImageCompressionService 作成 | ✅ 完了 |
| 2025-10-28 | Phase 1-2 | ShapeSelectionService 作成 | 🔄 作業中 |
| 2025-10-29 | Phase 2 | PowerTool フォルダ + Helper 作成 | ⏳ 未着手 |
| 2025-10-30 | Phase 3-1 | Table フォルダ作成（2ファイル） | ⏳ 未着手 |
| 2025-10-31 | Phase 3-2 | BuiltInShapeService 作成 | ⏳ 未着手 |
| 2025-11-01 | 最終調整 | 統合テスト・動作確認 | ⏳ 未着手 |

---

## 🎯 削減見込み

| 項目 | 削減行数 | 状態 |
|------|---------|------|
| ImageCompressionService | 約650行 | ✅ 完了 |
| ShapeSelectionService | 約200-300行 | 🔄 作業中 |
| PowerToolServiceHelper | 約500-600行 | ⏳ 未着手 |
| TableConversionService | 約600-700行 | ⏳ 未着手 |
| MatrixOperationService | 約1,000-1,200行 | ⏳ 未着手 |
| BuiltInShapeService | 約700-800行 | ⏳ 未着手 |
| **合計削減見込み** | **約3,650-4,250行** | - |
| **PowerToolService残存** | **約5,000-5,500行** | - |

---

## 📝 作業ログ

### 2025-10-28 15:51 JST
- ✅ ImageCompressionService.cs 作成完了
- 新規フォルダ `Services/Core/Image/` 作成
- コミット: feat: Create ImageCompressionService (Phase 1-1)
- 次の作業: Phase 1-2 (ShapeSelectionService) に着手

---

## ⚠️ 注意事項

1. **DI対応**: すべての新規サービスは `IApplicationProvider` をDIで受け取る
2. **名前空間**: `PowerPointEfficiencyAddin.Services.Core.{FolderName}` 形式
3. **既存パターン準拠**: 1フォルダ1-2ファイル、PascalCase命名
4. **段階的テスト**: 各Phaseごとに動作確認を実施
5. **ロールバック準備**: 問題発生時は即座に戻せるようコミット単位を細かく

---

**最終更新**: 2025-10-28 15:51 JST
