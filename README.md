# PowerPoint VBAアドイン - スライド選択

## 概要

このアドインは、PowerPointで以下の機能を提供します：

- **スライド一覧ダイアログ表示** - 全スライドの番号を一覧表示
- **選択したスライドへジャンプ** - ダイアログから選択したスライドへ即座に移動
- **ダブルクリック対応** - リストのダブルクリックでジャンプ
- **マウスホイール対応** - ListBox 上でスクロール可能（64bit対応）

---

## かんたんインストール（ユーザー向け）

配布物:

- [slideselector.pptm](slideselector.pptm)
- [install.bat](install.bat)

手順:

1. `install.bat` をダブルクリック
2. PowerPoint を完全に終了
3. PowerPoint を再起動
4. リボンに **「スライド選択」** タブが表示されます
5. **「スライド選択」** ボタンで起動

---

## 手動セットアップ（VBAエディタ）

1. PowerPoint を開く
2. `Alt + F11` で VBA エディタを開く
3. 標準モジュールを2つ作成し、以下を貼り付け:
   - [ModuleSlideJumper.bas](ModuleSlideJumper.bas)
   - [slideselector/ListBoxMouseScroll.bas](slideselector/ListBoxMouseScroll.bas)
4. ユーザーフォーム `UserFormSlideSelector` を作成し、
   - [UserFormSlideSelector.frm](UserFormSlideSelector.frm) のコードを貼り付け
5. コントロールを配置:
   - ListBox: `lstSlides`
   - Button: `cmdJump`（ジャンプ）
   - Button: `cmdCancel`（キャンセル）

---

## 使い方

### リボンから起動
1. **「スライド選択」** タブの **「スライド選択」** ボタンをクリック
2. ダイアログでスライド番号を選択
3. **ジャンプ** またはダブルクリックで移動

### VBA から起動（テスト用）
1. `Alt + F11` → `Ctrl + G`
2. イミディエイトウィンドウで実行:
   ```
   ShowSlideDialog
   ```

---

## ファイル一覧

| ファイル | 説明 |
|---------|------|
| ModuleSlideJumper.bas | メイン処理（ダイアログ表示、ジャンプ） |
| slideselector/ListBoxMouseScroll.bas | ListBox のホイール対応（64bit） |
| UserFormSlideSelector.frm | ユーザーフォームコード |
| UserFormSlideSelector.frx | ユーザーフォーム定義 |
| slideselector.pptm | 配布用アドイン（customUI 統合済み） |
| customUI.xml | リボン定義（統合済み） |
| install.bat | かんたんインストーラー |

---

## 更新履歴

- 2026-02-03: 64bit対応、ホイール対応、リボン統合、ドキュメント更新
# PowerPoint VBAアドイン - スライド選択

## 概要

このアドインは、PowerPointで以下の機能を提供します：

- **スライド一覧ダイアログ表示** - 全スライドのリストをダイアログで表示
- **選択したスライドへジャンプ** - ダイアログから選択したスライドへ即座に移動
- **スマートなバリデーション** - 無効なスライド番号を自動補正
- **ダブルクリック対応** - ダブルクリックでワンクリック実行可能

---

## インストール手順

### クイックセットアップ（5分）

#### ステップ 1: PowerPointを開く
```
任意のスライドファイルを開く
```

#### ステップ 2: VBAエディタを開く
```
Alt + F11
```

#### ステップ 3: 標準モジュールを作成

1. 左側「プロジェクト」ウィンドウでプレゼンテーション名を右クリック
2. `挿入` → `モジュール` を選択

#### ステップ 4: モジュールコードを貼り付け

[ModuleSlideJumper.bas](ModuleSlideJumper.bas) の内容をコピーして、作成したモジュールに貼り付けます

#### ステップ 5: ユーザーフォームを作成

1. VBAエディタでプレゼンテーション名を右クリック
2. `挿入` → `ユーザーフォーム` を選択
3. フォーム名を **`UserFormSlideSelector`** に変更

#### ステップ 6: コントロール（部品）を追加

**リストボックス:**
- ツールボックスから **ListBox** をドラッグしてフォーム上に配置
- Properties で `Name` を **`lstSlides`** に設定

**[ジャンプ]ボタン:**
- ツールボックスから **CommandButton** をドラッグしてフォーム上に配置
- Properties で `Name` を **`cmdJump`**、`Caption` を **ジャンプ** に設定

**[キャンセル]ボタン:**
- ツールボックスから **CommandButton** をドラッグしてフォーム上に配置
- Properties で `Name` を **`cmdCancel`**、`Caption` を **キャンセル** に設定

#### ステップ 7: ユーザーフォームのコードを追加

[UserFormSlideSelector.frm](UserFormSlideSelector.frm) のコード部分をコピーして、ユーザーフォーム（UserFormSlideSelector）のコードウィンドウに貼り付けます

#### ステップ 8: 保存
```
Ctrl + S で保存（*.pptm 形式を指定）
```

---

## 使用方法

### 方法1: VBAエディタからの実行（テスト用）
1. `Alt + F11` でVBAエディタを開く
2. `Ctrl + G` でイミディエイトウィンドウを表示
3. 以下を入力して **Enter** キーを押す：
   ```
   ShowSlideDialog
   ```

### 方法2: PowerPointのリボンボタンとして追加（推奨）

1. PowerPoint で: `ファイル` → `オプション` → `リボンのユーザー設定`
2. 新しいタブとグループを作成
3. マクロ `Project.ModuleSlideJumper.ShowSlideDialog` を追加
4. ボタンの表示名を「スライド選択」などに変更
5. OK をクリック

リボンに「スライド選択」ボタンが表示されます

### PowerPointリボンからのボタン追加（カスタマイズ）
1. `ファイル` → `オプション` → `リボンのユーザー設定`
2. 新しいタブまたはグループを作成
3. マクロ `ModuleSlideJumper.ShowSlideDialog` を割り当て

### 操作
1. ダイアログが表示され、全スライドのリストが表示されます
2. 移動したいスライドをリストから選択
3. `[ジャンプ]` ボタンをクリック、または選択行をダブルクリック
4. 指定したスライドへジャンプします

---

## ファイル説明

### ModuleSlideJumper.bas
- `ShowSlideDialog()` - ダイアログを表示するメインサブ
- `JumpToSlide(slideNumber)` - 指定されたスライド番号へジャンプ

### UserFormSlideSelector.frm
- ユーザーフォームの表示と制御
- ListBoxコントロール - スライド一覧表示
- Commandボタン - [ジャンプ]と[キャンセル]

### UserFormSlideSelector.frx
- バイナリ形式のユーザーフォーム定義ファイル

---

## 機能詳細

### スライド情報の取得
- スライド番号を表示
- スライドのタイトルテキストを取得（存在する場合）
- タイトルがない場合は「（タイトルなし）」と表示

### ジャンプ機能
- **スライドショー実行中の場合** - `GotoSlide` メソッドで対象スライドへ遷移
- **通常編集画面の場合** - スライドショーを開始して指定スライドから開始

### エラー処理
- プレゼンテーションが開いていない場合
- スライドが存在しない場合
- スライド遷移エラー時

---

## トラブルシューティング

### マクロが実行できない
- PowerPointのマクロセキュリティ設定を確認
- ファイルを信頼できる場所に配置

### ダイアログが表示されない
- VBAエディタでコードを確認
- PowerPointを再起動してみる

### スライドジャンプが機能しない
- スライドショーが実行されているか確認
- スライド番号が正しいか確認
- PowerPoint再起動後に再試行

---

## 注意事項
- このアドインはPowerPoint 2010以降で動作確認されています
- VBAは古い技術のため、互換性の問題が発生する可能性があります
- 重要なプレゼンテーションのテスト前にバックアップを取ることをお勧めします

---

## ライセンス
MIT2.0ライセンスの下で提供されています。詳細は [LICENSE](LICENSE) ファイルを参照してください。

---

## 更新履歴
- 2026-02-02: 初版作成

## 参考にしたページ
- https://www.relief.jp/docs/powerpoint-vba-goto-slide.html
- https://note.com/pippi_777/n/n7e1a1c8e7308#BpMWR
- https://jizilog.com/vba-listbox-userform
- https://kitagawa.group/hobby/excel_vba/listbox%E3%81%AE%E3%83%9E%E3%82%A6%E3%82%B9%E3%82%B9%E3%82%AF%E3%83%AD%E3%83%BC%E3%83%AB%E5%8C%96/