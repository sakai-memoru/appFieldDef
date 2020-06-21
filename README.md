# [Excel VBA] appFieldDef : 画面レイアウトからHTMLを生成するツール

## overveiw :

- プロトタイプでForm画面を作成する際に利用する設計補助ツール(VBA Excel)。
- 画面レイアウト設計書から、HTMLファイルを生成する。設計情報は、JSONファイルにも出力できる。

### 機能 :
- Excelにて、画面レイアウトを作成し、指定フォルダ（input folder)に配置する。  
    + 画面レイアウト定義では、ルールに沿って、レイアウトするコンポーネントを配置すること。
- メニューより、以下の機能を利用する。
    + 画面レイアウトにて定義した情報をもとに、定義情報をJSONファイルに出力する。  
    + 画面レイアウトにて定義した情報をもとに、プロトタイプ用のHTMLフォームのHTMLを生成する。  (CSSはBootstrap4)  

## 実行環境 :

- Microsoft Excel for Microsoft 365 MSO (16.0.13001.20142) 32-bit

## Installation :

- GitHubより、Cloneする。  
    +  https://github.com/sakai-memoru/appFieldDef  

- 参照設定が必要。  

![image](https://gyazo.com/7d30f2387e7818067fd7596a82e507e9.png) 



## Usage :
- アプリは以下。
    + アプリ本体  ：appFieldDef.xlsm  
        - Batch : FieldDefMain.bas 
            + OutputJson  
            + OutputTemplate
    + アプリconfig：config.json  
    + 定義情報取得form：forms/_\_TRANS_FIELDS__.xlsm  
         - GetFields  : TransFields.bas
    + 定義情報取得config：forms/_\_TRANS_FIELDS__.config.json 

- appFieldDef.xlsmを開く。  

![image](https://gyazo.com/76dc4c13106ac9fc03dbb064d6c37f43.png)  


### 初期コンフィグ設定 :
   
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
    "INPUT_LIKE": "SSE_BD_SD*.xlsx",  // 対象inputシート名
    "FORM_FILE": "__TRANS_FIELDS__.xlsm",
    "FORM_SHEET": "form",
    "MACRO_GET_METHOD": "GetFields"
  },
  "CONTROL_PREFIX": "__",
  "SOURCE_FROM": "_source",
  "APP_NAME" : "appFieldDef"
}
```  

### Environment :

![env](https://gyazo.com/ca8236f45f17fe16e8eac31618e166ff.png)

## Execution sample :

- Excel Layout : SSE_SD_Order.xlsx  
![image excel](https://gyazo.com/aabe9329bfb59a21ac59675e2f520beb.png)

- HTML Layout : SSE_SD_Order_Order_Form_200601182618.html  
![image html](https://gyazo.com/207f59e741bf505395ba26f6fb4aabdc.png)  

- Excel Layout : SSE_SD_User.xlsx  
![image excel2](https://gyazo.com/9c3788f029f4f7fe9e098e2744e4bbaa.png)

- HTML Layout : SSE_SD_User_User_Form_200601191504.html  
![image html](https://gyazo.com/dde2bd52d2c1ba632201f95440acb7e4.png)  

## application I/F :

```vb
'''' **********************************************
'''' @file FieldDefMain.bas
'''' @parent appFieldDef.xlsm
''''

Public Function Batch(ByVal datatype As String, _
                      Optional ByVal outTemplOn As Variant = False, _
                      Optional ByVal moveOn As Variant = False) As Variant
'''' **********************************************
'''' @function batch
'''' @param datatype {String} 処理データタイプ "FIELD_DEF"
'''' @param outTemplOn {Variant<boolean>}
''''            Template出力flag
'''' @param moveOn  {Variant<boolean>}
''''            Inputファイル移動flag
''
```

## note :
- 落ち着いたら、もう少し記述を追加します。  

## reference :

- 以下の外部ライブラリを使用しています。  
  + VBA-JSON : JsonConverter.bas  
    - https://github.com/VBA-tools/VBA-JSON  
  + MiniTemplator  
    - https://www.source-code.biz/MiniTemplator/  

// --- end of README.md