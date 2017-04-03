---
title: XPath.SetValue メソッド (Excel)
keywords: vbaxl10.chm760076
f1_keywords:
- vbaxl10.chm760076
ms.prod: EXCEL
api_name:
- Excel.XPath.SetValue
ms.assetid: 9d7e9eea-0962-cff8-6909-b31d349eb78a
ms.locale: ja-JP
---


# XPath.SetValue メソッド (Excel)

指定された  ** [XPath](xpath-object-excel.md)** オブジェクトを、 ** [ListColumn](listcolumn-object-excel.md)** オブジェクトまたは ** [Range](range-object-excel.md)** コレクションに対してマップします。 **XPath** オブジェクトが既に、 **ListColumn** オブジェクトまたは **Range** コレクションに対してマップされている場合は、 **SetValue** メソッドは **XPath** オブジェクトのプロパティを設定します。
 


## 構文

 *式*  . **SetValue**( ***Map***, ***XPath***, ***SelectionNamespace***, ***Repeating*** )
 

 
 *式*  **XPath** オブジェクトを表す変数です。
 

 

### パラメーター



|**名前**|**必須/オプション**|**データ型**|**説明**|
|:-----|:-----|:-----|:-----|
| _Map_|必須|** [XmlMap](xmlmap-object-excel.md)**|マップ範囲に関連付けるマップ情報を指定します。|
| _XPath_|必須|**文字列型 (String)**|このマップ範囲に表示する XML データの有効な XPath 式を指定します。XPath 文字列にはフィルターを含めることができます。フィルターを含めた場合、XPath で指定されたデータのサブセットだけがこのマップ範囲に表示されます。|
| _SelectionNamespace_|省略可能|**バリアント型 (Variant)**|**XPath** 引数内で使用されている、任意の名前空間接頭辞を指定します。 **XPath** オブジェクトが接頭辞をまったく保持していない場合や、XPath オブジェクトが Microsoft Office Excel 接頭辞を使用している場合は、この引数を省略することもできます。|
| _Repeating_|省略可能|**バリアント型 (Variant)**|**XPath** オブジェクトが、XML リストにある 1 つの列に対してバインドされているのか、それとも 1 つのセルに対してマップされているのかを指定します。 **True** の場合は、 **XPath** オブジェクトを、XML リストにある 1 つの列にバインドします。 **False** の場合は、非繰り返しセルを強制的に作成します。範囲に複数のセルが含まれるとき、 **False** が指定された場合、ランタイム エラーが発生します。|

## 注釈

Excel における XPath サポートの詳細については、 ** [IsExportable](xmlmap.isexportable-property-excel.md)** プロパティを参照してください。 XPath 式が無効であるか、指定された XPath が既にマップ済みであった場合は、ランタイム エラーが発生します。
 

 
Excel が名前空間の解決に失敗した場合、ランタイム エラーが発生します。
 

 
次のいずれかの条件に該当する場合はエラーが発生します。
 

 

- 範囲がグリッドの複数列にまたがっている。
    
 
- 範囲の一部分しかセルにマップされていない (セルのマップされていない範囲が存在する)。
    
 
- 同じ範囲に対して異なるマッピングが指定されているか、異なる XPath が指定されている。
    
 

 

 
範囲が単一セルであった場合、既定により、単一対応付けセル (非繰り返しマップのセル) が作成されます。非繰り返しセルに見出しは設定されません。
 

 
ただし、単一セル範囲が ListObject 内に存在する場合、マッピング情報は列全体に適用されます。
 

 
範囲が複数のセルにまたがっている場合、繰り返し XML リストが自動的に作成されます。選択された範囲はすべてデータ値として扱われます。つまり、XML リストの作成時、範囲は 1 行下に移動され、一番上のセルに見出しが設定されます。移動された範囲の一番下の行は挿入行になります。
 

 

 <BR/><BR/>**メモ**<BR/>   Excel の見出し検出アルゴリズムは、オブジェクト モデルでは使用されません。グリッド内に見出しは存在しないものと見なされます。 <BR/>オブジェクト モデルにおけるマップ範囲の作成時は、セルの結合やサイズ調整を自動で行う機能は無効化されます。
 


 

 

## 例

次の使用例は、ワークブックに添付されているスキーマ マップ "Contacts" に基づいて 1 つの XML リストを作成し、 **SetValue** メソッドを使用して、各列を 1 つの **XPath** オブジェクトにバインドします。
 

 

```
Sub CreateXMLList() 
    Dim mapContact As XmlMap 
    Dim strXPath As String 
    Dim lstContacts As ListObject 
    Dim objNewCol As ListColumn 
 
    ' Specify the schema map to use. 
    Set mapContact = ActiveWorkbook.XmlMaps("Contacts") 
     
    ' Create a new list. 
    Set lstContacts = ActiveSheet.ListObjects.Add 
         
    ' Specify the first element to map. 
    strXPath = "/Root/Person/FirstName" 
    ' Map the element. 
    lstContacts.ListColumns(1).XPath.SetValue mapContact, strXPath 
 
    ' Specify the second element to map. 
    strXPath = "/Root/Person/LastName" 
    ' Add a column to the list. 
    Set objNewCol = lstContacts.ListColumns.Add 
    ' Map the element. 
    objNewCol.XPath.SetValue mapContact, strXPath 
 
    strXPath = "/Root/Person/Address/Zip" 
    Set objNewCol = lstContacts.ListColumns.Add 
    objNewCol.XPath.SetValue mapContact, strXPath 
End Sub 

```


## 関連項目


#### 概念


 
 [XPath オブジェクト](xpath-object-excel.md)
#### その他の技術情報


 
 