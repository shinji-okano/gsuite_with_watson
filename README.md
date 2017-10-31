G Suite with Watson
============================================
GoogleのG Suite（Gmail、フォーム、スプレッドシート、GAS等）とIBM Watson™を組み合わせた業務支援ツールのテンプレート集です。メールやフォーム、RSS等からスプレッドシート上に蓄積したデータを、Watsonへの学習や、Watsonを活用した処理に利用できます。

テンプレートは主に利用Watson APIと、データ収集元により分かれます。現状、Watson APIはNatural Language Classifier（NLC）のみに対応しています。各テンプレートは次のコンポーネントが含まれます。

- データの収集と動作設定用Google スプレッドシート  
- IBM Watson™と連携してデータを処理するためのGASスクリプト


## 各テンプレート概要
- ### form
	- Google フォームからのデータを収集・処理
	- Natural Language Classifierを利用した分類
- ### gmail
	- Gmailからのデータを収集・処理
	- Natural Language Classifierを利用した分類
- ### rss
	- RSSからのデータを収集・処理
	- Natural Language Classifierを利用した分類

## 環境構築
各テンプレートは、Googleドライブ上で、スプレッドシートの形で公開されています。各スプレッドシートには、対応するGASスクリプトも内包されています。導入のための最も簡単な方法は、このGoogle Drive上で公開されているスプレッドシートをコピーする方法です。各スプレッドシートのリンクは下を参照してください。

なんらかの理由で上記の方法が取れない場合は、手動で環境を構築するための手順も用意しています。具体的な手順について
は下のリンク先よりご確認ください。
- ### form
	- [Googleドライブから環境をコピー](https://drive.google.com/drive/folders/0B_L8p3LDeJqLb2tNVXVXMFpKcjg)
	- [手動で環境構築](https://github.com/softbank-developer/gsuite_with_watson/tree/master/form)
- ### gmail
	- [Googleドライブから環境をコピー](https://drive.google.com/drive/folders/0B_L8p3LDeJqLb2tNVXVXMFpKcjg)
	- [手動で環境構築](https://github.com/softbank-developer/gsuite_with_watson/tree/master/gmail)
- ### rss
	- [Googleドライブから環境をコピー](https://github.com/softbank-developer/gsuite_with_watson/tree/master/rss)
	- [手動で環境構築](https://github.com/softbank-developer/gsuite_with_watson/tree/master/rss)


## ライセンス
[MIT](https://accounts.google.com/https://github.com/softbank-developer/gsuite_with_watson/blob/master/LICENSE)
