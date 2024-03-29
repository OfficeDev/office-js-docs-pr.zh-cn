---
title: 创建字典任务窗格加载项
description: 了解如何创建字典任务窗格加载项。
ms.date: 07/15/2022
ms.localizationpriority: medium
ms.openlocfilehash: f02b128166ba66eca5db54ceb98ee25e4f3bea56
ms.sourcegitcommit: b6a3815a1ad17f3522ca35247a3fd5d7105e174e
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/22/2022
ms.locfileid: "66959005"
---
# <a name="create-a-dictionary-task-pane-add-in"></a>创建字典任务窗格加载项

本文中的示例展示了任务窗格加载项和随附 Web 服务，用于提供用户当前在 Word 2013 文档中选择的内容的字典定义或同义词库同义词。

字典 Office 外接程序基于标准任务窗格外接程序，它具有附加功能来支持在 Office 应用程序的 UI 中的其他位置查询和显示字典 XML Web 服务的定义。

在典型的字典任务窗格加载项中，用户在文档中选择某字词或短语，加载项依据的 JavaScript 逻辑将此选定内容传递给字典提供程序的 XML Web 服务。然后，字典提供程序的网页更新为，向用户显示选定内容的定义。XML Web 服务组件最多以 OfficeDefinitions XML 架构定义的格式返回三个定义，然后会在主机 Office 应用的 UI 中的其他位置向用户显示这些定义。图 1 展示了用户选择的内容，以及 Word 2013 中运行的必应品牌字典加载项显示的内容。

*图 1：显示选定字词的定义的字典加载项*

![显示定义的字典应用。](../images/dictionary-agave-01.jpg)

由你决定选择字典加载项的 HTML UI 中的 **“查看更多** ”链接是在任务窗格中显示详细信息，还是打开一个单独的浏览器窗口，指向所选单词或短语的完整网页。
图 2 显示上下文菜单中的 **“定义** ”命令，使用户能够快速启动已安装的字典。 图 3 至 5 显示了 Office 用户界面中使用字典 XML 服务提供 Word 2013 定义的位置。

*图 2.定义上下文菜单中的命令*

![定义上下文菜单。](../images/dictionary-agave-02.jpg)

*图 3.“拼写”和“语法”窗格中的定义*

![拼写和语法窗格中的定义。](../images/dictionary-agave-03.jpg)

*图 4.“同义词库”窗格中的定义*

![同义词库窗格中的定义。](../images/dictionary-agave-04.jpg)

*图 5.“阅读模式”中的定义*

![读取模式下的定义。](../images/dictionary-agave-05.jpg)

若要创建可提供字典查找的任务窗格外接程序，需创建两个主要组件：

- XML Web 服务，该服务从字典服务中查找定义，然后以字典加载项可以使用和显示的 XML 格式返回这些值。
- 任务窗格加载项，它将用户的当前选择提交至字典 Web 服务，显示定义，还可以选择将这些值插入文档。

以下各节提供了有关如何创建这些组件的示例。

## <a name="creating-a-dictionary-xml-web-service"></a>创建字典 XML Web 服务

XML Web 服务必须将对 Web 服务的查询作为符合 OfficeDefinitions XML 架构的 XML 返回。以下两节介绍了 OfficeDefinitions XML 架构，并提供有关如何对返回该 XML 格式查询的 XML Web 服务编码的示例。

### <a name="officedefinitions-xml-schema"></a>OfficeDefinitions XML 架构

以下代码显示用于 OfficeDefinitions XML 架构的 XSD。

```XML
<?xml version="1.0" encoding="utf-8"?>
<xs:schema
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xmlns:xs="https://www.w3.org/2001/XMLSchema"
  targetNamespace="http://schemas.microsoft.com/NLG/2011/OfficeDefinitions"
  xmlns="http://schemas.microsoft.com/NLG/2011/OfficeDefinitions">
  <xs:element name="Result">
    <xs:complexType>
      <xs:sequence>
        <xs:element name="SeeMoreURL" type="xs:anyURI"/>
        <xs:element name="Definitions" type="DefinitionListType"/>
      </xs:sequence>
    </xs:complexType>
  </xs:element>
  <xs:complexType name="DefinitionListType">
    <xs:sequence>
      <xs:element name="Definition" maxOccurs="3">
        <xs:simpleType>
          <xs:restriction base="xs:normalizedString">
            <xs:maxLength value="400"/>
          </xs:restriction>
        </xs:simpleType>
      </xs:element>
    </xs:sequence>
  </xs:complexType>
</xs:schema>
```

符合 OfficeDefinitions 架构的返回 XML 由根`Result`元素组成，该元素包含从零到三`Definition`个`Definitions`子元素的元素，每个元素包含长度不超过 400 个字符的定义。 此外，必须在元素中 `SeeMoreURL` 提供字典网站的完整页面的 URL。 以下示例演示返回的符合 OfficeDefinitions 架构的 XML 的结构。

```XML
<?xml version="1.0" encoding="utf-8"?>
<Result xmlns="http://schemas.microsoft.com/NLG/2011/OfficeDefinitions">
  <SeeMoreURL xmlns="">www.bing.com/dictionary/search?q=example</SeeMoreURL>
  <Definitions xmlns="">
    <Definition>Definition1</Definition>
    <Definition>Definition2</Definition>
    <Definition>Definition3</Definition>
  </Definitions>
 </Result>

```

### <a name="sample-dictionary-xml-web-service"></a>示例字典 XML Web 服务

以下 C# 代码提供了一个有关如何为 XML Web 服务编写代码的简单示例，该服务以 OfficeDefinitions XML 格式返回字典查询的结果。

```cs
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Services;
using System.Xml;
using System.Text;
using System.IO;
using System.Net;

/// <summary>
/// Summary description for _Default.
/// </summary>
[WebService(Namespace = "http://tempuri.org/")]
[WebServiceBinding(ConformsTo = WsiProfiles.BasicProfile1_1)]
// To allow this web service to be called from script, using ASP.NET AJAX, uncomment the following line. 
// [System.Web.Script.Services.ScriptService]
public class WebService : System.Web.Services.WebService {

    public WebService () {

        // Uncomment the following line if using designed components.
        // InitializeComponent(); 
    }

    // You can replace this method entirely with your own method that gets definitions
    // from your data source and then formats it into the OfficeDefinitions XML format. 
    // If you need a reference for constructing the returned XML, you can use this example as a basis.
    [WebMethod]
    public XmlDocument Define(string word)
    {

        StringBuilder sb = new StringBuilder();
        XmlWriter writer = XmlWriter.Create(sb);
        {
            writer.WriteStartDocument();
            
                writer.WriteStartElement("Result", "http://schemas.microsoft.com/NLG/2011/OfficeDefinitions");

                    // See More URL should be changed to the dictionary publisher's page for that word on their website.
                    writer.WriteElementString("SeeMoreURL", "http://www.bing.com/search?q=" + word);

                    writer.WriteStartElement("Definitions");
            
                        writer.WriteElementString("Definition", "Definition 1 of " + word);
                        writer.WriteElementString("Definition", "Definition 2 of " + word);
                        writer.WriteElementString("Definition", "Definition 3 of " + word);
                   
                    writer.WriteEndElement();

                writer.WriteEndElement();
            
            writer.WriteEndDocument();
        }
        writer.Close();

        XmlDocument doc = new XmlDocument();
        doc.LoadXml(sb.ToString());

        return doc;
    }
}
```

## <a name="creating-the-components-of-a-dictionary-add-in"></a>创建字典加载项的组件

字典加载项包含三个主要组件文件：

- 描述加载项的 XML 清单文件。
- 提供加载项 UI 的 HTML 文件。
- JavaScript 文件，用于提供从文档中获取用户选择的逻辑，将选择作为查询发送给 Web 服务，然后在外接程序的 UI 中显示返回的结果。

### <a name="creating-a-dictionary-add-ins-manifest-file"></a>创建字典加载项的清单文件

下面是字典加载项的示例清单文件。

```XML
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.0" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:type="TaskPaneApp">
  <Id>7164e750-dc86-49c0-b548-1bac57abdc7c</Id>
  <Version>15.0</Version>
  <ProviderName>Microsoft Office Demo Dictionary</ProviderName>
  <DefaultLocale>en-us</DefaultLocale>
  <!--DisplayName is the name that will appear in the user's list of applications.-->
  <DisplayName DefaultValue="Microsoft Office Demo Dictionary" />
  <!--Description is a 2-3 sentence description of this dictionary. -->
  <Description DefaultValue="The Microsoft Office Demo Dictionary is an example built to demonstrate how a publisher could create a dictionary that integrates with Office. It does not return real definitions." />
  <!--IconUrl is the URI for the icon that will appear in the user's list of applications.-->
  <IconUrl DefaultValue="http://officeimg.vo.msecnd.net/_layouts/images/general/office_logo.jpg" />
  <SupportUrl DefaultValue="[Insert the URL of a page that provides support information for the app]" />
  <!--Capabilities specifies the kind of Office application your dictionary add-in will support. You shouldn't have to modify this area.-->
  <Capabilities>
    <Capability Name="Workbook"/>
    <Capability Name="Document"/>
    <Capability Name="Project"/>
  </Capabilities>
  <DefaultSettings>
    <!--SourceLocation is the URL for your dictionary-->
    <SourceLocation DefaultValue="http://christophernlg/ExampleDictionary/DictionaryHome.html" />
  </DefaultSettings>
  <!--Permissions is the set of permissions a user will have to give your dictionary. If you need write access, such as to allow a user to replace the highlighted word with a synonym, use ReadWriteDocument. -->
  <Permissions>ReadDocument</Permissions>
  <Dictionary>
    <!--TargetDialects is the set of regional languages your dictionary contains. For example, if your dictionary applies to Spanish (Mexico) and Spanish (Peru), but not Spanish (Spain), you can specify that here. Do not put more than one language (for example, Spanish and English) here. Publish separate languages as separate dictionaries. -->
    <TargetDialects>
      <TargetDialect>EN-AU</TargetDialect>
      <TargetDialect>EN-BZ</TargetDialect>
      <TargetDialect>EN-CA</TargetDialect>
      <TargetDialect>EN-029</TargetDialect>
      <TargetDialect>EN-HK</TargetDialect>
      <TargetDialect>EN-IN</TargetDialect>
      <TargetDialect>EN-ID</TargetDialect>
      <TargetDialect>EN-IE</TargetDialect>
      <TargetDialect>EN-JM</TargetDialect>
      <TargetDialect>EN-MY</TargetDialect>
      <TargetDialect>EN-NZ</TargetDialect>
      <TargetDialect>EN-PH</TargetDialect>
      <TargetDialect>EN-SG</TargetDialect>
      <TargetDialect>EN-ZA</TargetDialect>
      <TargetDialect>EN-TT</TargetDialect>
      <TargetDialect>EN-GB</TargetDialect>
      <TargetDialect>EN-US</TargetDialect>
      <TargetDialect>EN-ZW</TargetDialect>
    </TargetDialects>
    <!--QueryUri is the address of this dictionary's XML web service (which is used to put definitions in additional contexts, such as the spelling checker.)-->
    <QueryUri DefaultValue="http://christophernlg/ExampleDictionary/WebService.asmx/Define?word="/>
    <!--Citation Text, Dictionary Name, and Dictionary Home Page will be combined to form the citation line (for example, this would produce "Examples by: Microsoft", where "Microsoft" is a hyperlink to http://www.microsoft.com).-->
    <CitationText DefaultValue="Examples by: " />
    <DictionaryName DefaultValue="Microsoft" />
    <DictionaryHomePage DefaultValue="http://www.microsoft.com" />
  </Dictionary>
</OfficeApp>
```

以下 `Dictionary` 部分介绍了特定于创建字典加载项清单文件的元素及其子元素。 有关清单文件中的其他元素的信息，请参阅 [Office 外接程序 XML 清单](../develop/add-in-manifests.md)。

### <a name="dictionary-element"></a>Dictionary 元素

指定用于字典外接程序的设置。

**父元素**

**\<OfficeApp\>**

**子元素**

**\<TargetDialects\>**, **\<QueryUri\>**, **\<CitationText\>**, **\<Name\>**, **\<DictionaryHomePage\>**

**备注**

创建字典加载项时，元素 `Dictionary` 及其子元素将添加到任务窗格加载项的清单中。

#### <a name="targetdialects-element"></a>TargetDialects 元素

指定此字典支持的区域语言集。对于字典加载项，此为必需元素。

**父元素**

**\<Dictionary\>**

**子元素**

**\<TargetDialect\>**

**备注**

该 `TargetDialects` 元素及其子元素指定字典包含的区域语言集。 例如，如果字典同时适用于西班牙语（墨西哥）和西班牙语（秘鲁），但不适用于西班牙语（西班牙），可以在此元素中进行指定。 请勿在此清单中指定多种语言（例如，西班牙语和英语）。 请将各种语言发布为单独的字典。

**示例**

```XML
<TargetDialects>
  <TargetDialect>EN-AU</TargetDialect>
  <TargetDialect>EN-BZ</TargetDialect>
  <TargetDialect>EN-CA</TargetDialect>
  <TargetDialect>EN-029</TargetDialect>
  <TargetDialect>EN-HK</TargetDialect>
  <TargetDialect>EN-IN</TargetDialect>
  <TargetDialect>EN-ID</TargetDialect>
  <TargetDialect>EN-IE</TargetDialect>
  <TargetDialect>EN-JM</TargetDialect>
  <TargetDialect>EN-MY</TargetDialect>
  <TargetDialect>EN-NZ</TargetDialect>
  <TargetDialect>EN-PH</TargetDialect>
  <TargetDialect>EN-SG</TargetDialect>
  <TargetDialect>EN-ZA</TargetDialect>
  <TargetDialect>EN-TT</TargetDialect>
  <TargetDialect>EN-GB</TargetDialect>
  <TargetDialect>EN-US</TargetDialect>
  <TargetDialect>EN-ZW</TargetDialect>
</TargetDialects>
```

#### <a name="targetdialect-element"></a>TargetDialect 元素

指定此字典支持的一种区域语言。对于字典加载项，此为必需元素。

**父元素**

**\<TargetDialects\>**

**备注**

以 RFC1766 `language` 标记格式中指定区域语言的值，如 EN-US。

**示例**

```XML
<TargetDialect>EN-US</TargetDialect>
```

#### <a name="queryuri-element"></a>QueryUri 元素

指定字典查询服务的终结点。对于字典加载项，此为必需元素。

**父元素**

**\<Dictionary\>**

**备注**

这是字典提供程序的 XML Web 服务的 URI。被正确转义的查询将被追加到此 URI。

**示例**

```XML
<QueryUri DefaultValue="http://msranlc-lingo1/proof.aspx?q="/>
```

#### <a name="citationtext-element"></a>CitationText 元素

指定要在引文中使用的文本。对于字典加载项，此为必需元素。

**父元素**

**\<Dictionary\>**

**备注**

此元素指定将在从 Web 服务返回的内容之下的行中显示的引文文本的开头（例如，“Results by:”或“Powered by:”）。

对于此元素，可以使用该元素指定其他区域设置的 `Override` 值。 例如，如果用户正在运行 Office 的西班牙语 SKU，但使用的是英语字典，则允许引文行读取“Resultados por: Bing”，而不是“Results by: Bing”。 有关如何指定其他区域设置的值的详细信息，请参阅 [Office 外接程序 XML 清单](../develop/add-in-manifests.md)中的“为不同区域设置提供设置”一节。

**示例**

```XML
<CitationText DefaultValue="Results by: " />
```

#### <a name="dictionaryname-element"></a>DictionaryName 元素

指定此字典的名称。对于字典加载项，此为必需元素。

**父元素**

**\<Dictionary\>**

**备注**

此元素指定引文文本中的链接文本。引文文本显示在从 Web 服务返回的内容之下的行中。

对于此元素，可以指定其他区域设置的值。

**示例**

```XML
<DictionaryName DefaultValue="Bing Dictionary" />
```

#### <a name="dictionaryhomepage-element"></a>DictionaryHomePage 元素

指定字典的主页 URL。对于字典加载项，此为必需元素。

**父元素**

**\<Dictionary\>**

**备注**

此元素指定引文文本中的链接 URL。引文文本显示在从 Web 服务返回的内容之下的行中。

对于此元素，可以指定其他区域设置的值。

**示例**

```XML
<DictionaryHomePage DefaultValue="http://www.bing.com" />
```

### <a name="creating-a-dictionary-add-ins-html-user-interface"></a>创建字典外接程序的 HTML 用户界面

以下两个示例演示用于演示字典外接程序的 UI 的 HTML 和 CSS 文件。若要查看 UI 在外接程序的任务窗格中如何显示，请参阅代码之后的图 6。若要查看 Dictionary.js 文件中 JavaScript 的实现如何为此 HTML UI 提供编程逻辑，请参阅本节后面紧接着的“编写 JavaScript 实现”。

```HTML
<!DOCTYPE html>
<html>

<head>
<meta http-equiv="X-UA-Compatible" content="IE=Edge"/>

<!--The title will not be shown but is supplied to ensure valid HTML.-->
<title>Example Dictionary</title>

<!--Required library includes.-->
<script type="text/javascript" src="http://ajax.microsoft.com/ajax/4.0/1/MicrosoftAjax.js"></script>
<script type="text/javascript" src="office.js"></script>

<!--Optional library includes.-->
<script type="text/javascript" src="http://ajax.aspnetcdn.com/ajax/jQuery/jquery-1.5.1.js"></script>

<!--App-specific CSS and JS.-->
<link rel="Stylesheet" type="text/css" href="style.css" />
<script type="text/ecmascript" src="dictionary.js"></script>
</head>

<body>
<div id="mainContainer">
    <div id="header">
        <span id="headword"></span>
        <span id="pronunciation">(<a id="pronunciationLink">Pronounce</a>)</span>
    </div>
    <ol id="definitions">
    </ol>
    <div id="SeeMore">
    <a id="SeeMoreLink">See More...</a>
    </div>
</div>
</body>

</html>
```

以下示例显示 Style.css 的内容。

```CSS
#mainContainer
{
    font-family: Segoe UI;
    font-size: 11pt;
}

#headword
{
    font-family: Segoe UI Semibold;
    color: #262626;
}

#pronunciation
{
    margin-left: 2px;
    margin-right: 2px;
}

#definitions
{
    font-size: 8.5pt;
}
a
{
    font-size: 8pt;
    color: #336699;
    text-decoration: none;
}
a:visited
{
    color: #993366;
}
a:hover, a:active
{  
    text-decoration: underline;
}
```

*图 6.演示词典 UI*

![演示字典 UI。](../images/dictionary-agave-06.jpg)

### <a name="writing-the-javascript-implementation"></a>编写 JavaScript 实现

以下示例显示 Dictionary.js 文件中的 JavaScript 实现（该文件从外接程序的 HTML 页面调用，以提供演示字典外接程序的编程逻辑）。 该脚本重新使用以前介绍的 XML Web 服务。 脚本作为示例 Web 服务被置于同一目录中时将从该服务获取定义。 它可与符合 OfficeDefinitions 的公共 XML Web 服务一起使用，方法是修改 `xmlServiceURL` 文件顶部的变量，然后将发音的必应 API 密钥替换为正确注册的发音。

从此实现调用的 Office JavaScript API (Office.js) 的主要成员如下所示：

- 对象的`Office`[初始化](/javascript/api/office)事件，在初始化外接程序上下文时引发，并提供对文[档](/javascript/api/office/office.document)对象实例的访问权限，该实例代表加载项正在与之交互的文档。
- 对象的 `Document` [addHandlerAsync](/javascript/api/office/office.document#office-office-document-addhandlerasync-member(1)) 方法，该方法在函数中`initialize`调用，用于为文档的 [SelectionChanged](/javascript/api/office/office.documentselectionchangedeventargs) 事件添加事件处理程序，以侦听用户选择更改。
- 对象[的 getSelectedDataAsync](/javascript/api/office/office.document#office-office-document-getselecteddataasync-member(1)) 方法，在引发事件处理程序以获取用户所选的单词或短语时`SelectionChanged`在函数中`tryUpdatingSelectedWord()`调用该方法，将其强制转换为纯文本，然后执行`selectedTextCallback`异步回调函数。`Document`
- 当作为方法的`selectTextCallback`*回调* 参数传递的`getSelectedDataAsync`异步回调函数执行时，它会在回调返回时获取所选文本的值。 它通过使用返回的对象的 [值属性](/javascript/api/office/office.asyncresult#office-office-asyncresult-status-member)`AsyncResult`从回调的 *selectedText* 参数 (中获取该值，该参数的类型为 [AsyncResult](/javascript/api/office/office.asyncresult)) 。
- `selectedTextCallback` 函数中剩余的代码查询定义的 XML Web 服务。它还调入 Microsoft Translator API，以提供具有所选字词拼音的 .wav 文件的 URL。
- Dictionary.js 中的其余代码会在外接程序的 HTML UI 中显示定义的列表和拼音链接。

```js
// The document the dictionary add-in is interacting with.
let _doc;
// The last looked-up word, which is also the currently displayed word.
let lastLookup;
// For demo purposes only!! Get an AppID if you intend to use the Pronunciation service for your feature.
const appID="3D8D4E1888B88B975484F0CA25CDD24AAC457ED8";

// The base URL for the OfficeDefinitions-conforming XML web service to query for definitions.
const xmlServiceUrl = "WebService.asmx/Define?Word=";

// Initialize the add-in.
// The initialize function is required for all add-ins.
Office.initialize = function (reason) {
    // Checks for the DOM to load using the jQuery ready method.
    $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    // Store a reference to the current document.
    _doc = Office.context.document;
    // Check whether text is already selected.
    tryUpdatingSelectedWord();
    // Add a handler to refresh when the user changes selection.
    _doc.addHandlerAsync("documentSelectionChanged", tryUpdatingSelectedWord);
    });
}

// Executes when event is raised on user's selection changes, and at initialization time. 
// Gets the current selection and passes that to asynchronous callback function.
function tryUpdatingSelectedWord() {
    _doc.getSelectedDataAsync(Office.CoercionType.Text, selectedTextCallback); 
}

// Async callback that executes when the add-in gets the user's selection.
// Determines whether anything should be done. If so, it makes requests that will be passed to various functions.
function selectedTextCallback(selectedText) {
    selectedText = $.trim(selectedText.value);
    // Be sure user has selected text. The SelectionChanged event is raised every time the user moves the cursor, even if no selection.
    if (selectedText != "") { 
        // Check whether user selected the same word the pane is currently displaying to avoid unnecessary web calls.
        if (selectedText != lastLookup) { 
            // Update the lastLookup variable.
            lastLookup = selectedText; 
            // Set the "headword" span to the word you looked up.
            $("#headword").text(selectedText); 
            // AJAX request to get definitions for the selected word; pass that to refreshDefinitions.
            $.ajax(xmlServiceUrl + selectedText, { dataType: 'xml', success: refreshDefinitions, error: errorHandler }); 
            // AJAX request to the Microsoft Translator APIs. Gets the URL of a WAV file with pronunciation, which is passed to refreshPronunciation. See http://www.microsofttranslator.com/dev for details.
            $.ajax("http://api.microsofttranslator.com/V2/Ajax.svc/Speak?oncomplete=refreshPronunciation&amp;appId=" + appID + "&amp;text=" + selectedText + "&amp;language=en-us", { dataType: 'script', success: null, error: errorHandler }); 
        }
    }
}

// This function is called when the add-in gets back the definitions target word.
// It removes the old definitions and replaces them with the definitions for the current word.
// It also sets the "See More" link.
function refreshDefinitions(data, textStatus, jqXHR) {
    $(".definition").remove();
    // Make a new list item for each returned definition that was returned, set the CSS class, and append it to the definitions div.
    $(data).find("Definition").each(function () {
        $(document.createElement("li")).text($(this).text()).addClass("definition").appendTo($("#definitions"));
    });
    // Change the "See More" link to direct to the correct URL.
    $("#SeeMoreLink").attr("href", $(data).find("SeeMoreURL").text());
}

// This function is called when the add-in gets back the link to the pronunciation
// to set the "Pronounce" link to the URL of the .WAV file.
function refreshPronunciation(data) {
    $("#pronunciationLink").attr("href", data);
}

// Basic error handler that writes to a div with id='message'.
function errorHandler(jqXHR, textStatus, errorThrown) {
    document.getElementById('message').innerText += errorThrown;
}
```
