# expenses
[![Open in Visual Studio Code](https://img.shields.io/static/v1?logo=visualstudiocode&label=&message=Open%20in%20Visual%20Studio%20Code&labelColor=2c2c32&color=007acc&logoColor=007acc)](https://open.vscode.dev/hosoya17/expenses)
## 開発の概要
windows、mac共に動作する家計簿アプリです。
## システムの概要
・カレンダーから日付を選択し、追加ボタンをクリックします。<br>
・支出の場合、日付、金額、支払方法、カテゴリを入力または選択します。<br>
・収入の場合、日付、金額、カテゴリを入力または選択します。<br>
・追加ボタンをクリックしたら、家計簿.xlsxにデータが追加されます。
### 開発環境
開発環境：Visual Studio Code<br>
開発言語：python3<br>
ライブラリ:tkinter, tkinter.ttk, tkcalendar, openpyxl<br>
[![My Skills](https://skillicons.dev/icons?i=vscode,py)](https://skillicons.dev)
#### 環境構築
このプログラムはMicrosoftのExcelがインストールされていない場合、xlsxファイルが開けない為、使用することができません。<br>
<br>
事前にopenpyxlをインストールする必要があります。インストール方法は以下の通りです。<br>

```Shell
pip install openpyxl
```
<br>
また、openpyxlはバージョンによって文法が異なります。<br>
念のため以下の方法でアップグレードしてください。

```Shell
pip install openpyxl --upgrade
```
<br>
事前にtkcalendarをインストールする必要があります。インストール方法は以下の通りです。<br>

```Shell
pip install tkcalendar
```
<br>
expenses.pyの148行目と166行目の''の中はExcelフォルダの家計簿.xlsxのパスを指定してください。<br>
以下に記述例を示します。<br>
<br>

148行目<br>

```python
wb = load_workbook('C:\\Python\\expenses\\Excel\\家計簿.xlsx')
```
<br>
166行目<br>

```python
wb.save('C:\\Python\\expenses\\Excel\\家計簿.xlsx')
```
