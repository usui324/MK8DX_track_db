# MK8DX_track_db
**マリオカート8DXのカスタムマッチにおけるコース毎の戦績を記録できるマクロ付きエクセルファイル。**  
[ラウンジ](https://www.mk8dx-lounge.com/)での対戦における戦績記録を想定。
- 対応言語: 日本語  
- 対応コース: 有料追加コンテンツ「コース追加パス」第4弾  まで

## 使い方
`./trackDB.xlsm`をダウンロードし、エクセルで開く。  
`./src`ディレクトリ内はgithubでの差分の可視化のために用いている為、使用の際に必要はない。

### データの入力
1. DataInputシートにレースで走ったコースとそのレースの順位を入力する。
2. 登録ボタンを押下する。

<img width="300" alt="registData" src="https://user-images.githubusercontent.com/54677286/189478406-8779796c-ba90-47bc-9b29-35e45b20a64b.png">



### ランキング
蓄積された入力データからいくつかのランキングを生成して、Displayシートに表示している。
- (不)人気コースランキング:  走ったレース数に基づいてランキングを生成。
- 得意(苦手)コースランキング: 平均順位、得点にもとづいてそれぞれランキングを生成。
- 上位期待値ランキング: 1, 2位を獲得したときの上位ボーナス点（後述）の期待値を算出し、その値に基づきランキングを生成。

  <img width="400" alt="ranking" src="https://user-images.githubusercontent.com/54677286/189478425-dc5ee28b-d3fb-4c5e-a3d9-c3da39a81089.png">


--- 


### コースMEMO（おまけ）
どのコースを押すかメモするための場所。  
カラムは3つ用意してあり、それぞれ`スタート順位, コース名, コース英略称`をメモするよう想定している。

  
<img width="300" alt="trackMemo" src="https://user-images.githubusercontent.com/54677286/189478449-717a22a7-55b0-408e-86ca-53db42e27052.png">


--- 

## その他
### 上位ボーナス点について
マリオカート8DXのフレンド戦（12人対戦）において3位以下の順位を取ると、`13 - (順位)`の点数が得られる。  
対して1位を取ると上記の式に`+3`されたポイント（`13 - 1 + 3 = 15ポイント`）、2位を取ると式に`+1`されたポイント（`13 - 2 + 1 = 12ポイント`）を獲得できる。  
よってこの`+3, +1` を上位を獲得したプレーヤーへのボーナス点ととらえて、この値をコースごとに算出してランキング化した。  
毎レース1位をとればこの値は`3.0`、毎レース2位を取れば`1.0`となる。3位以下しかとれない場合は`0.0`となる。
