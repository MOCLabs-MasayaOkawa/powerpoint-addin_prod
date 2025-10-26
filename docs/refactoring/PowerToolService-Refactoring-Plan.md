# PowerToolService.cs リファクタリング計画

## 📊 現状分析

- **ファイルサイズ**: 307.5 KB
- **推定総行数**: 約9,225行
- **メソッド数**: 50個以上
- **主な問題点**:
  - 単一責任の原則違反（1クラスに50以上の機能）
  - コードの重複（定型パターンの繰り返し）
  - 過剰なデバッグログ
  - 冗長なtry-catch
  - 不要なコメント

## 🎯 リファクタリング目標

- **削減見込み**: 約2,100行（22.8%削減）
- **整理後の推定**: 約7,125行
- **所要期間**: 2-3週間

---

## 📅 Phase 1: クイックウィン（削減見込み: 350行）

### 実施期間
- 開始日: 2025-10-26
- 完了予定日: 2025-10-28
- 所要日数: 1-2日

### タスク一覧

#### ✅ Task 1-1: デバッグログの整理（削減: 200行）
- **対象**: `logger.Debug()` 呼び出し（100箇所以上）
- **修正内容**: `#if DEBUG` ～ `#endif` で囲む
- **リスク**: 低（Debugビルド時のみ影響）
- **影響範囲**: なし（Release構成には影響なし）

#### ⬜ Task 1-2: 冗長なtry-catchの削除（削減: 100行）
- **対象**: `ComHelper.ExecuteWithComCleanup()` 内の二重try-catch
- **修正内容**: 内部try-catchを削除（ComHelperに任せる）
- **リスク**: 低
- **影響範囲**: エラーハンドリングは外側で実施されるため問題なし

#### ⬜ Task 1-3: 不要なコメント削除（削減: 50行）
- **対象**: 「★既存メソッド流用」「★修正点」などの冗長コメント
- **修正内容**: 実装済みの機能に関する不要コメントを削除
- **リスク**: 極低
- **影響範囲**: なし

---

## 📅 Phase 2: 重複統合（削減見込み: 600行）

### 実施期間
- 開始予定日: 2025-10-29
- 完了予定日: 2025-11-02
- 所要日数: 3-5日

### タスク一覧

#### ⬜ Task 2-1: セル書式コピーメソッドの統合（削減: 200行）
- **対象メソッド**:
  - `CopyCellFormat()`
  - `CopyTableCellFormatNew()`
  - `CopyObjectShapeFormat()`
- **修正内容**: 共通インターフェースで統合
- **リスク**: 中
- **影響範囲**: 表変換機能、マトリクス操作機能

#### ⬜ Task 2-2: 図形選択メソッドの統合（削減: 100行）
- **対象メソッド**:
  - `SelectShapes()`
  - `SelectCreatedShapes()`
- **修正内容**: 1つのメソッドに統合
- **リスク**: 低
- **影響範囲**: 全機能（選択処理）

#### ⬜ Task 2-3: グリッド検出ロジックの共通化（削減: 300行）
- **対象メソッド**:
  - `ConvertTableToTextBoxes()`
  - `ConvertTextBoxesToTable()`
  - `ExcelToPptx()`
  - `OptimizeMatrixRowHeights()`
  - その他マトリクス系全て
- **修正内容**: グリッド判定を共通メソッドに抽出
- **リスク**: 中
- **影響範囲**: マトリクス系機能全般

---

## 📅 Phase 3: 設計改善（削減見込み: 1,300行）

### 実施期間
- 開始予定日: 2025-11-03
- 完了予定日: 2025-11-16
- 所要日数: 1-2週間

### タスク一覧

#### ⬜ Task 3-1: 基底クラスパターン導入（削減: 800行）
- **実施内容**: 
  - `PowerToolServiceBase` 抽象クラス作成
  - 共通処理（選択取得、バリデーション、COM管理）を基底に移動
  - 全50メソッドをテンプレートメソッドパターンで簡素化
- **リスク**: 高
- **影響範囲**: 全機能

#### ⬜ Task 3-2: 戦略パターン適用（削減: 500行）
- **実施内容**:
  - `IMatrixHandler` インターフェース作成
  - `TableMatrixHandler` / `GridMatrixHandler` 実装
  - マトリクス系メソッドを統一インターフェースで処理
- **リスク**: 高
- **影響範囲**: マトリクス系機能全般

---

## 📂 修正対象セクション

### Section 1: パワーツールグループ (16-23) - 8メソッド
- ✅ `MergeText()` - Phase 1対応中
- ⬜ `MakeLineHorizontal()`
- ⬜ `MakeLineVertical()`
- ⬜ `SwapPositions()`
- ⬜ `SelectSimilarShapes()`
- ⬜ `ExcelToPptx()`
- ⬜ `AlignLineLength()`
- ⬜ `AddSequentialNumbers()`

### Section 2: 特殊機能グループ (24-27) - 4メソッド
- ⬜ `UnifyFont()`
- ⬜ `AlignLineLength()`
- ⬜ `AddSequentialNumbers()`

### Section 3: 表変換機能 - 10メソッド
- ⬜ `ConvertTableToTextBoxes()`
- ⬜ `ConvertTextBoxesToTable()`
- ⬜ `OptimizeMatrixRowHeights()`
- ⬜ `OptimizeTableComplete()`
- ⬜ `AddMatrixRowSeparators()`
- ⬜ `AlignShapesToCells()`
- ⬜ `AddHeaderRowToMatrix()`
- ⬜ `SetCellMargins()`
- ⬜ `AddMatrixRow()`
- ⬜ `AddMatrixColumn()`

### Section 4: 列幅・行高統一機能 - 4メソッド
- ⬜ `EqualizeColumnWidths()`
- ⬜ `EqualizeRowHeights()`
- ⬜ `MatrixTuner()`

### Section 5: 画像圧縮機能 - 5メソッド
- ⬜ `CompressImages()`
- ⬜ その他ヘルパーメソッド

### Section 6: 同色・同サイズ選択、透過率調整 - 4メソッド
- ⬜ `SelectSameColorShapes()`
- ⬜ `SelectSameSizeShapes()`
- ⬜ `TransparencyUpToggle()`
- ⬜ `TransparencyDownToggle()`

### Section 7: Built-in機能ヘルパー - 10メソッド以上
- ⬜ `ExecutePowerPointCommand()`
- ⬜ `ShowShapeStyleDialog()`
- ⬜ その他図形作成系メソッド

---

## 🔍 修正パターン詳細

### パターン1: デバッグログの整理

**修正前:**
```csharp
logger.Debug($"Merged text set to reference shape: {referenceShape.Name}");
logger.Debug($"Deleted shape: {shapeInfo.Name}");
```

**修正後:**
```csharp
#if DEBUG
logger.Debug($"Merged text set to reference shape: {referenceShape.Name}");
logger.Debug($"Deleted shape: {shapeInfo.Name}");
#endif
```

### パターン2: 冗長なtry-catch削除

**修正前:**
```csharp
ComHelper.ExecuteWithComCleanup(() =>
{
    try
    {
        // 処理
        logger.Debug($"...");
    }
    catch (Exception ex)
    {
        logger.Error(ex, "Failed...");
    }
}, shapes);
```

**修正後:**
```csharp
ComHelper.ExecuteWithComCleanup(() =>
{
    // 処理（ComHelperが例外処理を担当）
#if DEBUG
    logger.Debug($"...");
#endif
}, shapes);
```

### パターン3: 不要なコメント削除

**修正前:**
```csharp
// ★既存メソッド流用
var selectedShapes = GetSelectedShapeInfos();

// ★既存パターン流用
if (!ValidateSelection(...)) return;
```

**修正後:**
```csharp
var selectedShapes = GetSelectedShapeInfos();
if (!ValidateSelection(...)) return;
```

---

## 🧪 動作確認チェックリスト

### Phase 1 完了時
- [ ] Release構成でビルド成功
- [ ] Debug構成でビルド成功
- [ ] 各セクションの代表的な機能を手動テスト:
  - [ ] テキスト合成
  - [ ] 線を水平にする
  - [ ] 表をテキストボックスに変換
  - [ ] 画像圧縮
  - [ ] 同色図形選択

### Phase 2 完了時
- [ ] 全機能の回帰テスト
- [ ] マトリクス系機能の詳細テスト
- [ ] パフォーマンステスト（処理時間が悪化していないか）

### Phase 3 完了時
- [ ] 全機能の完全テスト
- [ ] コードレビュー
- [ ] ドキュメント更新

---

## 📈 進捗管理

### 全体進捗
- **Phase 1**: 🟡 進行中（1/3完了）
- **Phase 2**: ⬜ 未着手
- **Phase 3**: ⬜ 未着手

### 削減実績
- **現在の削減**: 0行
- **目標削減**: 2,100行
- **達成率**: 0%

---

## 📝 変更履歴

### 2025-10-26
- リファクタリング計画書を作成
- Phase 1 Task 1-1 開始: デバッグログの整理（Section 1から着手）

---

## 🔗 関連ドキュメント

- [PowerToolService.cs 現在のコード](../Services/PowerToolService.cs)
- [開発ポリシー v0.67](./development-policy-v0.67.md)

---

## 📞 引き継ぎ情報

### 現在の作業状態
- **実施中のPhase**: Phase 1
- **実施中のTask**: Task 1-1（デバッグログの整理）
- **実施中のSection**: Section 1（パワーツールグループ）
- **完了メソッド**: なし
- **次の作業**: Section 1 の8メソッドのデバッグログを整理

### 重要な注意点
1. 各セクションごとに修正してcommit/pushすること
2. 修正後は必ずビルド確認すること
3. 本計画書の進捗状況を随時更新すること
4. 大きな変更の前には必ずバックアップを取ること
