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
  * 
* 用途
  * 書式が決まっているExcelシートに患者情報を自動で入力することができます。
  * 患者毎の位置照合結果シートに患者情報を自動入力、
  * 

# Demo

![Screen capture of planCompare UI](https://github.com/tkmd94/ExcelSheetGenerator/blob/master/demo.gif)

# Requirement

* Eclipse 15.6 以上 (古いバージョンではチェックされていません。)

# Installation
1. [MainWindow.xaml.cs]の３５行目　
2. ```PreferencePath = @"D:\Dshare\ExcelSheetGenerator\ExcelSheetGeneratorPreference.xml"];```

3. このプロジェクトをビルドして、DLL ファイル [ExcelSheetGenerator.esapi.dll] を生成します。
4. 生成された DLL ファイルを、Eclipse ツールバーの [ツール] -> [スクリプト] で指定したフォルダーにコピーします。
5. このスクリプトをお気に入りとしてマークし、キーボード ショートカットを設定します。
6. ARIAeDocProfile_ENU_ESAPI_ScreenCapture.xml を ARIA eDoc の Profiles フォルダーにコピーし、設定を変更します。

# Usage

**※本ソースコードは自己責任で使用してください。**

1. キャプチャしたい画面を表示します。
2. 登録したキーボード ショートカットからスクリプトを実行します。
3. スクリプトの実行が完了すると、[OK] ダイアログが表示されます。
 
# Author
 
* Takashi Kodama
 
# License
 
"ScreenCapture_eDoc" は [MIT ライセンス](https://en.wikipedia.org/wiki/MIT_License) の下にあります。
