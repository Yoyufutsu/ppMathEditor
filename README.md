# ppMath 数式エディタ（PowerPoint アドイン）

PowerPoint の作業ウィンドウで LaTeX から数式を作り、スライドに SVG として挿入するアドインです。
挿入した数式には LaTeX ソースが図形タグとして埋め込まれるので、**スライド上で数式を選択するとエディタに読み込まれ、そのまま再編集**できます。

エディタ本体は GitHub Pages（https://yoyufutsu.github.io/ppMathEditor/taskpane.html）から読み込むので、ローカルサーバーは不要です（ネット接続は必要）。

## 機能
- LaTeX で入力、MathJax でリアルタイムプレビュー（AMS / physics / mhchem / braket / cancel 対応）
- サイズ(pt)・色・ディスプレイ/インライン体裁
- **Ctrl+Enter** で挿入/更新
- **自動更新**（既定オン）：編集中の数式は、入力が 0.6 秒止まるとスライドに自動反映（構文エラー中は反映しない）
- 更新時は元の位置と、手動で拡大縮小した倍率を維持
- プレゼン内の数式一覧（クリックでその数式へジャンプして編集）
- `\newcommand` などのマクロ定義（すべての数式に適用）

## インストール（1 回だけ）
共有フォルダー カタログ方式で登録します（再起動しても消えません）。**管理者として開いた PowerShell** で：
```powershell
cd "C:\Users\minat\Documents\ツール作成\ppMathEditor"
powershell -ExecutionPolicy Bypass -File .\install-catalog.ps1
```
スクリプトがすること：`%USERPROFILE%\ppMathCatalog` に manifest.xml をコピー → そのフォルダーを自分だけ読み取り可で共有 → Office の「信頼できるアドイン カタログ」に登録 → 開発用（デバッグ）登録を削除。

その後 PowerPoint を再起動し、**ホーム → アドイン → その他のアドイン（個人用アドイン）→ 共有フォルダー → ppMath → 追加**。
以後は **ホーム** タブ右端に「ppMath → 数式エディタ」ボタンが常に出ます。

削除は同じく管理者 PowerShell で `powershell -ExecutionPolicy Bypass -File .\uninstall-catalog.ps1`。

## エディタを改良したとき
`docs/` を編集して GitHub に push すれば、数分で Pages に反映されます（PowerPoint 側の再登録は不要）。
`manifest.xml` を変えたときだけ `install-catalog.ps1` をやり直してください。

### ローカルで試す（開発用）
```powershell
npm.cmd run certs        # 初回のみ：localhost 用の開発証明書
npm.cmd start            # https://localhost:3000 + manifest.dev.xml（「ppMath dev」グループ）
npm.cmd run stop
```
UI だけブラウザで確認：`npm.cmd run preview` → http://localhost:3000/taskpane.html

## 仕組み
- 数式は MathJax で SVG 化して `setSelectedDataAsync(..., XmlSvg)` で挿入
- 図形タグ `PPMATH_LATEX` / `PPMATH_SIZE` / `PPMATH_COLOR` / `PPMATH_DISPLAY` / `PPMATH_H` に情報を保存
- タグは .pptx に保存されるので、開き直しても再編集できます（アドインのない PC では普通の画像として表示）

## ファイル
- `manifest.xml` … 本番用マニフェスト（GitHub Pages を参照）
- `manifest.dev.xml` … 開発用マニフェスト（localhost を参照）
- `docs/` … エディタ本体（GitHub Pages の公開フォルダ）
- `server.js` … 開発用の静的サーバー
