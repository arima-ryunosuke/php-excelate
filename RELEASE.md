# RELEASE

バージョニングはセマンティックバージョニングでは**ありません**。

| バージョン   | 説明
|:--           |:--
| メジャー     | 大規模な仕様変更の際にアップします（クラス構造・メソッド体系などの根本的な変更）。<br>メジャーバージョンアップ対応は多大なコストを伴います。
| マイナー     | 小規模な仕様変更の際にアップします（中機能追加・メソッドの追加など）。<br>マイナーバージョンアップ対応は1日程度の修正で終わるようにします。
| パッチ       | バグフィックス・小機能追加の際にアップします（基本的には互換性を維持するバグフィックス）。<br>パッチバージョンアップは特殊なことをしてない限り何も行う必要はありません。

なお、下記の一覧のプレフィックスは下記のような意味合いです。

- change: 仕様変更
- feature: 新機能
- fixbug: バグ修正
- refactor: 内部動作の変更
- `*` 付きは互換性破壊

## x.y.z

- パースがメチャクチャなので構文を php に寄せて token_get_all あたりで楽したい

## 1.1.2

- [fixbug] 2回レンダリングされている不具合
- [fixbug] template 範囲が A1 を含まないと template タグが残存する不具合

## 1.1.1

- [feature] col を追加＋簡易フォーマットの指定機能
- [refactor] (row|col)shift が異常に遅かったので是正
- [fixbug] (row|col)shift でデータが空の時にテンプレートが残る不具合
- [fixbug] insertNew(Row|Column) の負数で消すと変なデータが残ることがある
- [fixbug] (row|col)shift でテンプレート範囲が増減しない不具合
- [fixbug] A1 に template タグを入れると他の構文が吹き飛ぶ不具合

## 1.1.0

- [*change] 互換のために残していた render メソッドを削除
- [*change] Variable を削除
- [composer] phpspreadsheet:1,2,3
- [composer] phpunit9
- [change] php8.2 のエラーを修正

## 1.0.7

- [feature] row 構文を追加

## 1.0.6

- [feature] rowcol 構文を追加

## 1.0.5

- [feature] アクティブシートの指定と完了コールバック機能
- [feature] 入力規則エフェクタを追加

## 1.0.4

- [change] メソッド名の整合性のため既存 render を renderSheet に改名
- [feature] book ごとレンダリングする renderBook を追加
- [Utils] copyCells が遅いのを改善

## 1.0.3

- [refactor] 気になる箇所を修正
- [fixbug][Renderer] roweach がネストしているとき行数によっては動作しない不具合を修正
- [feature][Utils] デバッグがとても困難なので dumpCellValues をセル範囲を表形式で出力する処理に変更

## 1.0.2

- bump version
  - php: 7.4

## 1.0.1

- [feature][Renderer] template 範囲を指定しなかったときのデフォルトをデータの範囲に修正

## 1.0.0

- 公開
