## 概要
Googleフォームでは回答をスプレッドシートに記録することができます。その回答内容をNatural Language Classifier(NLC)を利用して任意のクラスタに分類するツールです。

<img src="https://github.com/softbank-developer/gsuite_with_watson/blob/master/form/readme_images/logo.png" width="60px">

&emsp;↓

![img](https://github.com/softbank-developer/gsuite_with_watson/blob/master/form/readme_images/finished.png)

- `タイムスタンプ〜質問3`：フォームからの結果
- `分類器[1-3]:手入力`: 手動による分類(NLCの学習対象列)
- `分類器[1-3]:Watson`: NLCによる分類
1つのNLCで3つまでの分類器を利用します。


## 利用条件
- [Google](https://accounts.google.com/)アカウントを持っていること
  - Google フォーム、Google スプレッドシート用
- [IBM Bluemix](https://accounts.google.com/)アカウントを持っていること
  - Natural Language Classifier用(一つ以上のインスタンスを用意しておくこと)


## 準備
Google ドライブにてスプレッドシートとGaoogle スプレッドシートとGoogle Apps Scriptを用意しています。ドライブからコピーして利用する場合は、本準備手順の7番のみの対応で準備は終わります。

1. スプレッドシートの作成  
任意の名称でスプレッドシートを作成します。

2. "設定"シートの作成
作成したスプレッドシート内に、各種設定を管理するシートを作成します。
sheetsディレクトリ内に、サンプルのシートを置いています。

	(シート設定)
	- `データシート名`: データの格納シート
	- `開始列`: データを格納する開始列
	- `開始行`: データを格納する開始行
	- `学習・分類対象`: NLCによる学習・分類の対象
	- `分類器[1-3]:手入力`: 手動で分類させる列
	- `分類器[1-3]:watson`: 分類結果を記載する列
	- `分類日時[1-3]`: 分類された日時を記載する列
	- `ログシート名`:  NLCの学習と分類ログの保存シート
	- `メール通知`:  メール通知機能の切り替え
	- `メール通知設定シート名`: メール通知条件の指定

	(分類器)
	- `Classifier ID`: NLCの分類器のID(学習後自動で挿入されます)

	![img](https://github.com/softbank-developer/gsuite_with_watson/blob/master/form/readme_images/config.png)

3. "データ管理"シートの作成
作成したスプレッドシート内に、データ管理用のシートを作成します。
sheetsディレクトリ内に、サンプルのシートを置いています。各項目の行列番号は、設定シートでの指定に従っ
てください。

3. "データ管理"シートの作成   
フォームからの回答内容を管理します。
sheetsディレクトリ内に、サンプルのシートを置いています。
タイムスタンプから質問3まではフォームから投げられるデータが挿入されます。

4. "ログ"保存シートの作成
作成したスプレッドシート内に、ログ保存用のシートを作成します。
sheetsディレクトリ内に、サンプルのシートを置いています。

5. "通知"シートの作成
作成したスプレッドシート内に、メールで通知を行う設定用のシートを作成します。
sheetsディレクトリ内に、サンプルのシートを置いています。
	- `分類器`: 情報を取得源
	- `フィードURL`：フィード用のURL

6. GASスクリプトの読み込み
スクリプトエディタを起動(ツール -> スクリプト エディタ)し、すべての.gsファイルをインポートします。
	- コード.gs
	- NLCLIB.gs
	- FORMLIB.gs

7. NLCの属性値の設定
  NLCの属性値をScript propertiesとして設定します(ファイル -> プロパティ)。
	- `CREDS_URL`: [NLC用のURL]
	- `CREDS_USERNAME`: [NLC用のユーザ名]
	- `CREDS_PASSWORD`: [NLC用のパスワード]


## 使い方
1. フォームの回答先の選択  
質問タブからフォームの質問設定が終わったのち、回答タブににて回答の保存先として作成したスプレッドシートを指定します。

	|質問設定|回答設定| 
	|---|---| 
	|![](https://github.com/softbank-developer/gsuite_with_watson/blob/master/form/readme_images/form1.png)|![](https://github.com/softbank-developer/gsuite_with_watson/blob/master/form/readme_images/form2.png)|

2. NLCの分類器の学習
	- `分類器[1-3]:手入力`へ学習させたいインテント名を付与します(空白を含んだ行は無視します)。
	- `Watson` -> `学習`を実行します。  
	学習したNLCの分類IDは設定用のシートに自動で記録されます。
	**Training**中は利用できません。**Available**になるまで待つ必要があります。
	![img](https://github.com/softbank-developer/gsuite_with_watson/blob/master/form/readme_images/menu.png)

3. データの分類
学習させたNLCの分類器を利用して分類させます。
	- `Watson` -> `分類`  
	分類結果はデータを記録するシートの分類器[1-3]:Watsonに記録されます。


## ライセンス
[MIT](https://accounts.google.com/https://github.com/softbank-developer/gsuite_with_watson/blob/master/LICENSE)
