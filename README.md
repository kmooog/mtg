# mtg draft calculator
内輪のMTGサークルで使用している、google spreadsheetを利用してマジックザギャザリング(MTG)のドラフトにおける勝率や各色の使用率、レーティングなどを計算するスクリプトです。
"list"という名前のシートを作成し、test.tsvに示すような形でデータを記述してください

レーティングの計算は、確率の逆数の端数切り上げで計算しています。
例えば、三回戦の場合は以下のようになります。
			
勝、負、確率の逆数、獲得ポイント

3、0、8、8

2、1、2.66、3

1、2、-2.66、-3

0、3、-8、-8
