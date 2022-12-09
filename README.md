# ExcelSheetGenerator
 
「ExcelSheetGenerator」は、データをExcelテンプレートファイルの指定セルに出力して保存するESAPI スクリプトです。

# Features

* 以下のデータをExcelテンプレートファイルの指定セルに出力して保存することができます。
  * Patient ID
  * Patient name
  * Course ID
  * Plan ID
  * Dose per fraction
  * No of fraction
  * Total dose
  * PlanApproval status
  * Isocenter position
  * Structure volume
* 患者・プラン情報を基にフォルダ名とファイル名を指定して作成することができます。
* 設定はXML形式で記述し、複数条件をダイアログで選択実行が可能です。
* Excelがインストールされていない端末でも実行可能です。

* 用途
  * 患者毎に作成するチェックシートに患者情報を自動入力

# Demo

![Screen capture of planCompare UI](https://github.com/tkmd94/ExcelSheetGenerator/blob/master/demo.gif)

# Requirement

* Eclipse 15.6 以上 (古いバージョンではチェックされていません。)

# Installation
1. Excelテンプレートファイルを作成します。
2.  [ExcelSheetGeneratorPreference.xml]を配置し、[MainWindow.xaml.cs]の３５行目のパスを変更します。

    ```PreferencePath = @"D:\Dshare\ExcelSheetGenerator\ExcelSheetGeneratorPreference.xml";```
    
3. [ExcelSheetGeneratorPreference.xml]の設定を修正します。
    * templatename：ダイアログで表示する名称
    * templatePath：テンプレートファイルパス
    * exportPath：保存するディレクトリパス
    * filename：保存するファイル名
    * type：データタイプ（info/DVH）
    * data：データ名
    * structure：データを出力する輪郭名（type:DVH,data:Volumeのみ）
    * address：データを出力するセルの位置

4. プロジェクトをビルドして、DLL ファイル [ExcelSheetGenerator.esapi.dll] を生成します。
5. 生成された DLL ファイルを、Eclipse ツールバーの [ツール] -> [スクリプト] で指定したフォルダーにコピーします
6. [2]で設定したテンプレートファイルパスにExcelテンプレートファイルを配置します。

# Usage

**※本ソースコードは自己責任で使用してください。**

1. Eclipse ツールバーの [ツール] -> [スクリプト] からスクリプトを実行します。
2. ダイアログウィンドウから条件を選択し、[Generate]ボタンを押下します。
3. スクリプトの実行が完了すると、保存先のディレクトリが表示されるので、作成されたファイルを確認します。
 
# Author
 
* Takashi Kodama
 
# License
 
"ExcelSheetGenerator" は [MIT ライセンス](https://en.wikipedia.org/wiki/MIT_License) の下にあります。
