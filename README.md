# CrossHexx
CrossHexxは、クロスヘア（レティクル）を画面上に表示するツールです。
* ブログ [ミジンコハック](https://mjh.blog.jp/)
* [Discord](https://discord.gg/s4vsJZrmNj)

## 機能
* クロスヘア / ドット / ドット＆サークル / 画像の4種類を表示
* 大きさ（Medium / Large）、色（RGB）、位置・オフセットの調整
* ウィンドウ中央への自動配置、右クリック中の非表示
* 設定の保存・読込（5プリセット＋起動時デフォルト指定）
* マルチモニタ対応（カーソル位置のスクリーン中央に表示）

## Usage
1. exeを起動
2. 「起動」ボタンで表示、「停止」ボタンで非表示
3. Setting1で種類・大きさ・色、Setting2で位置、Setting3で画像を設定
4. Save/Loadタブでプリセットを保存・読込

## Build
* GitHub Actionsが `master` へのpush毎にReleaseビルドし、Artifactsに `CrossHexx-Release` を保存します
* 正式版は `v*` タグのpushで自動作成され、Releasesページからzipで取得できます
* ローカルビルドには Visual Studio 2022（.NET Framework 4.8 targeting pack）が必要です
  * `CrossHexx.sln` を開いてReleaseビルド

## 動作環境
* Windows 10 / 11（.NET Framework 4.8以降）

## Licence
The Unlicense
