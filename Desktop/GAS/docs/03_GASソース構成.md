# 03_GASソース構成

## source/ 配下について
このパッケージでは、`ソース.zip` を展開したままの構成を保持しています。
**フォルダ名・ファイル名は変更していません。**

## プロジェクト1
### `1er9ofamXntX_iCYxiSZ9gTgPrzcIikYpG5dhKwjQo40xR65Mf1QSmma6`
主な役割: GSC取得、改善候補抽出、差分比較、OpenAI実行

含まれる主ファイル:
- `#U30b3#U30fc#U30c9.js` → `コード.js`
- `SEO#U30d7#U30ed#U30f3#U30d7#U30c8#U4f5c#U6210.js` → `SEOプロンプト作成.js`
- `GSC#U30ad#U30fc#U30ef#U30fc#U30c9#U30d9#U30fc#U30b9.js` → `GSCキーワードベース.js`
- `#U30d7#U30ed#U30f3#U30d7#U30c8API#U5b9f#U884c.js` → `プロンプトAPI実行.js`
- `#U5dee#U5206#U53d6#U5f97.js` → `差分取得.js`
- `#U5dee#U5206#U6bd4#U8f03.js` → `差分比較.js`
- `#U66f4#U65b0#U9593#U306eGSC#U5dee#U5206#U8abf#U67fb.js` → `更新間のGSC差分調査.js`
- `#U77ed#U671f#U8a55#U4fa1.js` → `短期評価.js`
- `#U30d6#U30e9#U30f3#U30c9#U4e00#U89a7#U30b5#U30de#U30ea#U4f5c#U6210(#U65e5#U6642).js` → `ブランド一覧サマリ作成(日時).js`
- `#U66f4#U65b0#U5c65#U6b74#U30d0#U30c3#U30af#U30a2#U30c3#U30d7.js` → `更新履歴バックアップ.js`

## プロジェクト2
### `1uSCA6wB10RozQoHPEZmI_eV3dxMBqoEiXVuXslov6PQdP57Yn2apfE3-`
主な役割: HTMLキャッシュ、内部リンク、階層マスタ、禁則ワード

含まれる主ファイル:
- `HTML#U30ad#U30e3#U30c3#U30b7#U30e5.js` → `HTMLキャッシュ.js`
- `#U5185#U90e8#U30ea#U30f3#U30af#U30de#U30b9#U30bf.js` → `内部リンクマスタ.js`
- `#U7981#U5247#U30ef#U30fc#U30c9.js` → `禁則ワード.js`
- `#U89aa#U5b50#U95a2#U4fc2#U6b63#U898f#U5316.js` → `親子関係正規化.js`
- `#U968e#U5c64#U30de#U30b9#U30bf#U66f4#U65b0.js` → `階層マスタ更新.js`

## appsscript.json
- 両プロジェクトとも `runtimeVersion: V8`
- GSC側プロジェクトは `SearchConsole v1` の Advanced Service を有効化
- `プロンプトAPI実行.js` は Script Properties の `OPENAI_API_KEY` を使用
