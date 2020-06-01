# [Excel VBA ]appFormDef :

## overveiw :

- プロトタイプを作成する際に、利用する画面レイアウトから、HTMLファイルを生成するVBAツール  

- Excelにて、画面レイアウトを作成する。  
  + 画面レイアウト作成では、あるルールに沿って、レイアウトするコンポーネントを配置すること。

- 画面レイアウトにて定義した情報をもとに、定義情報をJSONファイルに出力する。  
- 画面レイアウトにて定義した情報をもとに、プロトタイプ用のHTMLフォームのHTMLを生成する。  

## Installation :

- GitHubより、Cloneする。  

## Usage :
- アプリ本体  ：appFieldDef.xlsm  
- アプリconfig：config.json  
- 定義情報取得form：forms/_\_TRANS_FIELDS__.xlsm  
- 定義情報取得config：forms/_\_TRANS_FIELDS__.config.json 

- 初期コンフィグ設定  
```
{
  "BASE_FOLDER": "", 
  "INPUT_FOLDER": "input",      // inputシートフォルダ
  "OUTPUT_FOLDER": "output",    // 変換結果のアウトプット
  "TEMP_FOLDER": "input/temp",
  "BACKUP_FOLDER": "input/backup",
  "FORM_FOLDER": "forms",       // formアプリの配置
  "FIELD_DEF": {
    "SHEET_TYPE": "FIELD",
    "INPUT_LIKE": "SSE*.xlsx",  // 対象inputシート名
    "FORM_FILE": "__TRANS_FIELDS__.xlsm",
    "FORM_SHEET": "form",
    "MACRO_GET_METHOD": "GetFields"
  },
  "CONTROL_PREFIX": "__",
  "SOURCE_FROM": "_source",
  "APP_NAME" : "appFieldDef"
}
```

## note :
- 落ち着いたら、もう少し記述を追加します。  

