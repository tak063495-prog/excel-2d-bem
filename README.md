# Elastic BEM 2D / 3D — Excel VBA

弾性境界要素法（BEM）の解析本体をExcel VBAへ移植したブックです。
入力、オートメッシュ、dense/FMM演算、GMRES、内部変位・応力、結果出力をWindows版Excel内で実行します。
通常の解析にPythonや外部の数値計算DLLは不要です。

## ダウンロードと実行

- [VBA組込済みExcelブック](outputs/bem_vba_stage3/Elastic_BEM_VBA_Full.xlsm)
- [配布一式ZIP](outputs/Elastic_BEM_VBA_Stage3.zip)
- [詳細な操作説明](outputs/bem_vba_stage3/FMM・操作説明.md)
- [検証結果](outputs/bem_vba_stage3/検証結果.json)

1. ブックをWindows版Excelで開き、マクロを有効にします。
2. 「操作」B5で2D/3D、B6で2Dの平面状態、B7で定数0/二次2を選びます。
3. 「材料」「領域」「境界条件」「接合」「評価点」を入力して「メッシュ生成」を押します。
4. B13で`auto` / `dense` / `fmm`を選び、「VBA解析実行」を押します。
5. 「境界結果」「内点結果」を確認し、「JSON・CSV出力」で保存します。

FMM・GMRESの設定はB49:B61です。初期モデルは各辺16分割の2D二次要素で、FMMによる解析結果を保存しています。

## 実装

- 2D平面ひずみ・平面応力、3D線形等方弾性。
- 定数／不連続二次要素、Gauss／DE積分、自己項と境界近傍の処理。
- 混合境界条件、異材の完全接合、内部変位・応力6成分。
- 四分木／八分木、Chebyshev展開、上向き伝達・M2L・下向き伝達、近傍ブロック疎行列、容量制限付きLRUキャッシュ。
- ブロック前処理付き再始動GMRES。前処理の擬似逆は片側Jacobi SVD。denseではLUも選択できます。
- JSON、メッシュ／結果CSV、VTK、NumPy互換NPZ、演算・積分の詳細レポート。

FMM経路は全G/H行列やモデル全体の密行列を生成しません。`auto`はVBAで測定した演算時間とメモリ見積りにより領域ごとに選択します。

入力表は2D RECT/CIRCLE/POLYGON、3D BOX/SPHEREを扱います。境界グループ内の条件は一定値です。
任意CAD形状、領域ごとに異なる次数、非一致メッシュの接合・接触は入力画面の対応範囲に含みません。
「収束OK」は線形残差と内点積分の判定です。FMM展開誤差・離散化誤差はpとメッシュを変えた比較で確認してください。

## ソースと検証の再実行

`vba_src/`は編集用UTF-8ソース、`outputs/bem_vba_stage3/vba/`はVBEにインポートするWindows形式です。
ThisWorkbookのコードは標準モジュールではなく、ブックモジュールへ貼り付けます。
`source/elastic_bem_unified/elasticbem/`は添付元Pythonの数値比較用です。

2026-10-08、Windows版Excel 16.0で、Pythonとの作用ベクトル・境界解・内部場の比較と、Excelの入出力・設定変更を検証しました。
2D 1536未知数、3D二次1224未知数を含みます。今回追加した23項目ではdense/FMM併用、材料定数の範囲、設定変更後の無効化、評価点なし、複数領域の結果順序、失敗後の再実行を確認しています。
詳細な件数と数値は検証結果JSONを参照してください。

以下のPythonは開発・比較検証に使用します。利用者のVBA解析には不要です。
Python 3.12、NumPy、SciPy、pywin32、threadpoolctlと、Windows版デスクトップExcelを使用します。
検証時はExcelの「VBAプロジェクト オブジェクト モデルへのアクセスを信頼する」を有効にする必要があります。
通常のブック利用では、この開発用設定は不要です。

```powershell
python -m venv .venv
.venv\Scripts\python.exe -m pip install -r requirements-dev.txt
.venv\Scripts\python.exe verify_fmm.py
.venv\Scripts\python.exe verify_fmm_features.py
.venv\Scripts\python.exe verify_solver_edges.py --stage3
.venv\Scripts\python.exe verify_regression.py
```

検証スクリプトは専用の非表示Excelを起動し、ブックを読取専用で開きます。検証用VBAを一時注入し、保存せず閉じます。
利用中のExcelブックを閉じる処理はありません。検証の出力は`qa_stage3/`へ生成されます。
配布ブック内に検証用モジュールやPython起動処理はありません。
