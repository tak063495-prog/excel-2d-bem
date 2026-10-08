# 統合版：2次元／3次元 線形弾性BEM

2D・3Dで **同じ `elasticbem` パッケージ、同じ `run_bem.py`、同じJSON仕様**を使います。材料・Kelvin核・密行列組立・多段Chebyshev FMM・完全接着の連立処理・精度判定・出力を共通化しました。半無限地盤の演算子、地表ゲージ、入力例、起動BATは統合版から除去しています。閉じた有限領域のみです。

| 項目 | 2次元 | 3次元 |
|---|---|---|
| `dimension` | 2 | 3 |
| 応力状態 | `plane_strain` / `plane_stress` | 通常の3次元弾性 |
| 一定要素 `element_order: 0` | 直線2幾何節点、1場節点 | 三角形3幾何節点、1場節点 |
| 二次要素 `element_order: 2` | 曲線3幾何節点、3場節点 | 曲面三角形6幾何節点、6場節点 |
| 部材間の接着 | 変位2成分連続、表面力2成分釣合い | 変位3成分連続、表面力3成分釣合い |
| バックエンド | dense / FMM / auto | 同左、共通実装 |
| 内点積分 | Gauss / DE / auto | 同左、曲面写像と極座標変換 |

単位は m、kN、kPa。2Dの合力は奥行き単位長さ当たりkN/m、3DはkN。等方・線形・静的弾性、体積力なし。複数部材は完全接着です。摩擦接触・開離・非適合メッシュの界面補間は実装していません。

## 起動

Windows：`install_requirements.bat` → `run_2d.bat` または `run_3d.bat`。`run_2d.bat plane_stress` で平面応力の例を選べます。`run_model.bat examples\任意の入力.json` でも同じ起動口を使います。

```bash
python -m pip install -r requirements.txt
python run_bem.py examples/two_materials_2d_order2_plane_strain.json --out output_2d
python run_bem.py examples/two_materials_3d_order2_three_dimensional.json --out output_3d
python verify.py
python verify_fmm3d.py
```

`verify_fmm3d.py` は大きめの3D二次要素メッシュで密行列と遠方FMMを比較します。メモリは数百MB以上、実行時間は小規模の検証より長くなります。Linux・倍精度・BLAS 1スレッドで検証済み、Windows上の実行は未検証です。

## 入力と二次要素

`examples/` に単一部材／異材接着、2D両平面状態／3D、一定／二次要素の12例があります。`python generate_examples.py` で再生成できます。設定値はJSONで編集してください。

```json
{
  "dimension": 2,
  "state": "plane_strain",
  "element_order": 2,
  "solver": {"backend": "auto", "expected_iterations": 120},
  "quadrature": {"mode": "auto", "displacement_tol": 1e-8, "stress_tol": 0.001},
  "regions": [{
    "name": "body",
    "material": {"E": 30000, "nu": 0.3},
    "vertices": [[0,0],[1,0],[1,1],[0,1],[0.5,0],[1,0.5],[0.5,1],[0,0.5]],
    "elements": [[0,4,1],[1,5,2],[2,6,3],[3,7,0]],
    "boundary_conditions": [
      {"elements": [3], "type": ["u","t"], "values": [0,0]},
      {"elements": [0], "type": ["t","u"], "values": [0,0]},
      {"elements": [1], "type": ["t","t"], "values": [100,0]}
    ],
    "interior_points": [[0.5,0.5]]
  }],
  "interfaces": []
}
```

3Dでは `dimension: 3`、座標・境界条件を3成分、要素を6幾何節点へ変更します。`state` を省略するか `three_dimensional` とします。3Dで `plane_stress` 等を指定すると入力エラーです。

### 幾何節点と場の節点

**不連続二次境界要素**を採用しています。曲線・曲面は幾何節点で連続に接続しますが、変位・表面力の節点は要素ごとに独立です。幾何節点の値を、そのまま場の値として配列に並べないでください。角・稜線に異なる面の表面力を一値へ押し込めません。

* 2D line3 の幾何接続は `[始点, 中点, 終点]`、自然座標 -1,0,1。場の座標は -2/3,0,2/3。
* 3D tri6 の幾何接続は `[頂点0, 頂点1, 頂点2, 辺01中点, 辺12中点, 辺20中点]`。頂点の自然座標 `(s,t)` は (0,0),(1,0),(0,1)、中点は (.5,0),(.5,.5),(0,.5)。
* tri6 の6場節点は、それぞれの幾何自然座標を重心 (1/3,1/3) に向かって縮小した `重心 + 0.65*(幾何自然座標-重心)`。すべて要素内部です。この6点で完全な二次多項式を補間します。
* 3Dの中点を弦から移動すると曲面になります。積分点ごとに面のJacobianと法線を求め、FMMの源モーメントにも可変法線を含めます。
* 2Dは反時計回りの一つの閉曲線。3Dは一つの連結した水密閉曲面で外向き法線。同じ辺の幾何頂点・中点は隣接要素で同じ節点番号を共有してください。
* 3DのFMM支持箱は二次三角形のBernstein制御点による包絡を使い、曲面が幾何節点の外側へ出る場合も支持領域を含めます。
* 極端に歪んだ・折り返した曲面、自己交差、部材同士の重なりは使わないでください。入力検査はJacobian、標本点の法線、閉面の辺、向き・体積を検査しますが、自己交差を完全には検出しません。2Dの穴、3Dの別閉面による内部空洞はこの入力仕様では扱いません。

`element_order` はモデルの既定値で、領域内でも上書きできます。一つの領域内で要素の次数は統一します。接着界面の両側は同じ次数・要素分割・曲面形状を指定します。

### 外部境界条件と界面

番号は0始まり、領域名は英数字・`_`・`-`。

`type` は各成分について `u` (変位指定、m) / `t` (外向き表面力指定、kPa)。`values` は成分数2/3の一組、または場節点数1/3/6×成分数の配列です。配列の行順は上記の場節点の順序です。指定しない外部要素は表面力ゼロです。

```json
"interfaces": [{
  "region_a": "left", "elements_a": [2,3],
  "region_b": "right", "elements_b": [0,1]
}]
```

同じリスト位置で要素を対応させます。各要素内の幾何頂点の逆順・三角形の頂点置換に伴う場の対応は座標から自動認識します。法線は反対である必要があります。界面へ外部条件を重ねて指定するとエラーです。

\[u_A=u_B,\qquad t_A+t_B=0\]

を未知数の共有・符号反転で組み込みます。接続成分全体の剛体モードを除去する変位条件が必要です。2Dは3、3Dは6モードで、各部材を個別に固定する必要はありません。

## 自動選択と高速化

`solver` の指定例：

```json
{"backend":"auto", "leaf":24, "theta":0.7, "cache_mb":128,
 "expected_iterations":120, "memory_mb":2048}
```

`backend` は `auto` / `dense` / `fmm`。FMM次数 `p` の既定は2D:6、3D:4。指定範囲は2D:2〜8、3D:2〜6。密行列では `p` は解の精度を変更しません。

共通の四分木／八分木・Chebyshev展開で、P2M、M2M、M2L、L2L、L2Pを実行します。近接相互作用は直接積分です。近接源の葉をターゲットの葉ごとにまとめ、同じ幾何要素の積分を重複させないようにしています。

実機の組立ブロック・行列作用・実際の階層と相互作用数を測り、予想GMRES反復数を含めた時間を推定します。選択前に全密行列を組みません。予算はモデル全体で共用し、各領域の保守的な推定分を順次差し引きます。展開・キャッシュ・M2L生成用の一時領域も見積もり、FMMが予算外なら大きな推定用展開を作りません。これはOSのメモリ使用量の厳密上界ではありません。

`report.json` に両方式の推定時間・メモリ、選択・校正時間、実際の組立時間・反復数を保存します。メモリ制約で方式が不適なら、その時間推定をnullとする場合があります。`expected_iterations` は入力値による予測で、実際の連成問題の反復数・条件数や計算時間を保証しません。元の一定要素版の履歴学習機能は新統合版では使用していません。

3D二次要素では一要素6場節点となり、同じ要素数の一定要素より行列が大きくなります。小規模・遠方相互作用の少ないモデルでは密行列が有利です。FMMを固定すると常に高速になるわけではありません。

## DEと精度判定

内点の近接性を曲線・曲面への最近距離／要素長で判定し、近接ならDE、その他はGaussから開始します。Gaussの判定が成立しなければDEへ切り替えます。2Dは最近点分割とasinh変換、3Dは自然三角形の最近点を中心とする極座標・角度asinh・半径変換の後に曲面へ写像します。

`displacement_tol` は変位各成分の絶対許容値(m)、`stress_tol` は応力各成分(kPa)。2Dではzzも含みます。各要素へ許容値を配分し、二回連続の積分細分化差と丸め誤差の推定値を確認します。`quadrature` は `mode: auto/gauss/de`、`max_points` (点・要素ごと、既定200000)、`max_de_level` (6)、`max_gauss_level` (6)、`near_ratio` (0.2) も指定できます。

積分値を得たが判定未達の場合は値を保持し、`QUADRATURE_TOLERANCE_UNMET` を通知・記録します。内点以外、境界に極端に近い点、積分値を一度も得られない場合はNaN (JSONではnull)。GMRES未達も取得できた反復解を `LINEAR_SOLVE_TOLERANCE_UNMET` として保持します。出力完了後の終了コードは2、成功は0です。入力不正・剛体未拘束・予算不足は事前エラーです。

**これは積分の事後誤差推定であり、総誤差の保証ではありません。** 境界離散化・FMM近似・反復誤差・条件数は別に評価してください。境界分割の増加、FMM次数・thetaの変更、密行列比較を使います。近接内点で予算内に二回の判定を終えられない場合、計算値が厳密解に近くても未達フラグを出します。

境界積分の制御は別の `integration`：

```json
{"boundary_order":12, "boundary_de_level":1, "self_order":24}
```

上記は3Dの既定です。2Dの既定 `boundary_order` は48、`boundary_de_level` は3。自己相互作用は変位差で正則化し、2Dは分割DE、3Dは三分割Duffy積分 (半径Gauss、角度Gauss) を使います。強い曲率・近接・薄い部材では境界積分設定も上げて収束を確認してください。この境界積分は内点の物理許容値から自動保証されません。

## 出力・追加後処理

* `solution.npz`：寸法・入力モデル・各部材の幾何・場の節点・変位・表面力・内点値。
* `report.json`：線形解法、選択推定、部材合力、内点判定。CSVと組で解の状態を判断してください。
* `部材名_boundary.csv`：場の節点での境界変位・表面力。
* `部材名_boundary.vtk`：ParaView用。不連続場を保つため幾何節点を要素ごとに複製し、場の値を幾何節点へ補間／外挿します。2D二次線・3D二次三角形のVTKセル。
* `部材名_interior.csv`：変位・応力6成分・積分判定・エラーコード・値の有無。

保存済み境界解から内点を追加評価できます。

```bash
python evaluate_fields.py output_3d/solution.npz points_3d.csv --region left --out extra_fields
```

点CSVのヘッダは `x,y` または `x,y,z`。例を同梱しています。`--u-tol` / `--stress-tol` で許容値を変更できます。別部材の点を指定すると対象領域外として扱います。

## 元の成果物からの移行

有限領域の旧2D二次JSONは `dimension: 2` を追加し、そのままの幾何・界面・境界条件を使用できます。補助変換も可能です。

```bash
python convert_legacy.py old_2d_model.json --out model.json
python convert_legacy.py old_mesh.npz old_bc.npz --E 30000 --nu 0.3 --out model.json
```

旧NPZの `segments` / `triangles` と `is_displacement` / `values` は一定要素として取り込みます。材料・2D応力状態は変換時に指定してください。一定要素を二次要素へ変える場合は幾何中点と場節点の境界条件を新たに定義し、単に `element_order` を変更しないでください。

無限領域の入力を有限領域へ自動置換しません。旧 `kind: halfspace` や `reference_x` は拒否します。統合版へ半無限の基本解は含めていません。

## コード構成

`elasticbem/material.py` (材料・ND核・微分)、`meshes.py` (有限幾何と補間)、`rules.py` (積分規則)、`tree.py` (ND階層)、`operators.py` (共通dense/FMMと時間推定)、`model.py` (共通界面・GMRES)、`precision.py` (共通内点精度判定)、`io.py` (共通出力)。2Dと3Dを別フォルダから起動する仕組みではありません。

検証の結果と制約は `VALIDATION_ja.md`、機械可読値は `validation/` にあります。
