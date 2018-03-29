---
title: 指定 Office 主机和 API 要求
description: ''
ms.date: 12/04/2017
---

# <a name="specify-office-hosts-and-api-requirements"></a>指定 Office 主机和 API 要求

你的 Office外接程序可能依赖于特定的 Office 主机、要求集、API 成员或 API 版本才能按预期运行。例如，你的外接程序可能：

- 在单个 Office 应用程序（Word 或 Excel）或多个应用程序中运行。
    
- 使用仅在 Office 的某些版本中可用的 JavaScript API。例如，可能会在运行在 Excel 2016 中的外接程序中使用适用于 Excel 的 JavaScript API。 
    
- 只能在 Office 的某些版本中运行，这些版本支持供你的外接程序使用的 API 成员。
    
本文可帮助你了解应选择的选项，以确保你的外接程序按预期运行，并遍及可能的最广泛的访问群体。

> [!NOTE]
> 若要概览 Office 加载项的当前受支持情况，请参阅 [Office 加载项主机和平台可用性](../overview/office-add-in-availability.md)页面。 

下表列出了本文中讨论的核心概念。

|**概念**|**说明**|
|:-----|:-----|
|Office 应用程序、Office 主机应用程序、Office 主机或主机|Office 应用程序、Office 主机应用程序、Office 主机或主机|
|平台|平台|
|要求集|命名的一组相关的 API 成员。外接程序使用要求集来确定 Office 主机是否支持你的外接程序使用的 API 成员。测试对要求集的支持比对单个的 API 成员的支持更为容易。要求集支持根据 Office 主机和 Office 主机的版本变化。 <br >要求集在清单文件中指定。当你在清单中指定要求集时，你可以设置 Office 主机必须提供的用于运行你的外接程序的最低级别的 API 支持。不支持在清单中指定的要求集的 Office 主机不能运行外接程序，并且外接程序不会显示在“<span class="ui">我的外接程序</span>”中。这限制了外接程序的使用位置。在使用运行时检查的代码中。有关要求集的完整列表，请参阅 [Office 外接程序要求集](https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets)。|
|运行时检查|在运行时执行的一种测试，用以确定运行外接程序的 Office 主机是否支持要求集或外接程序使用的方法。若要执行运行时检查，请使用包含 **isSetSupported** 方法、要求集或不属于要求集的方法名称的 **if** 语句。使用运行时检查以确保达到的客户数目最大。与要求集不同，运行时检查不指定 Office 主机必须提供的用于运行外接程序的最低级别的 API 支持。相反，使用 **if** 语句来确定是否支持某个 API 成员。如果支持，则可以在外接程序中提供其他功能。使用运行时检查时，外接程序将始终在“**我的外接程序**”中显示。|

## <a name="before-you-begin"></a>开始之前

您的外接程序必须使用最新版本的外接程序清单架构。如果您在外接程序中使用运行时检查，请确保您使用的是适用于 Office 的最新 JavaScript API (office.js) 库。

### <a name="specify-the-latest-add-in-manifest-schema"></a>指定最新的外接程序清单架构

外接程序清单必须使用外接程序清单架构版本 1.1。按照以下操作设置外接程序清单中的 **OfficeApp**。

```XML
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:type="TaskPaneApp">
```

### <a name="specify-the-latest-javascript-api-for-office-library"></a>指定最新的适用于 Office 的 JavaScript API 库

如果使用运行时检查，则请引用内容传送网络 (CDN) 中的最新版本的适用于 Office 的 JavaScript API 库。若要执行此操作，请将以下 `script` 标记添加到 HTML 中。使用 CDN URL 中的 `/1/` 可以确保引用的是最新版本的 Office.js。

```HTML
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
```

## <a name="options-to-specify-office-hosts-or-api-requirements"></a>指定 Office 主机或 API 要求的选项

指定 Office 主机或 API 要求时，有几个决策因素需要考虑。下图显示了如何确定要在外接程序中使用的技术。

![指定 Office 主机或 API 要求时，选择最适用于加载项的选项](../images/options-for-office-hosts.png)

- 如果加载项在 Office 主机中运行，请在清单中设置 **Hosts** 元素。有关详细信息，请参阅[设置 Hosts 元素](#set-the-hosts-element)。
    
- 若要设置 Office 主机必须支持的用以运行外接程序的最低要求集或 API 成员，请在清单中设置 **Requirements** 元素。有关详细信息，请参阅[在清单中设置 Requirements 元素](#set-the-requirements-element-in-the-manifest)。
    
- 如果特定要求集或 API 成员可在 Office 主机中使用，在这种情况下如果你想要提供其他功能，请在外接程序的 JavaScript 代码中执行运行时检查。例如，如果外接程序在 Excel 2016 中运行，你可能想要使用新的适用于 Excel 的 JavaScript API 中的 API 成员以提供附加功能。有关详细信息，请参阅[在你的 JavaScript 代码中使用运行时检查](#use-runtime-checks-in-your-javascript-code)。
    
## <a name="set-the-hosts-element"></a>设置 Hosts 元素

若要使外接程序运行在一个 Office 主机应用程序中，请使用清单中的 **Hosts** 和 **Host** 元素。如果未指定 **Hosts** 元素，你的外接程序将在所有主机中运行。

例如，以下 **Hosts** 和 **Host** 声明指定外接程序将使用任何版本的 Excel，其中包括适用于 Windows 的 Excel、Excel Online 和适用于 iPad 的 Excel。

```xml
<Hosts>
  <Host Name="Workbook" />
</Hosts>
```

**Hosts** 元素可以包含一个或多个 **Host** 元素。**Host** 元素指定外接程序要求的 Office 主机。**Name** 属性是必需的，且可以被设置为下列值之一。

| 名称          | Office 主机应用程序                      |
|:--------------|:----------------------------------------------|
| 数据库      | Access Web App                               |
| 文档      | Word for Windows、Word for Mac、Word for iPad 和 Word Online        |
| 邮箱       | Outlook for Windows、Outlook for Mac、Outlook Web 和 Outlook.com | 
| 演示文稿  | PowerPoint for Windows、PowerPoint for Mac、PowerPoint for iPad 和 PowerPoint Online  |
| 项目       | 项目                                       |
| 工作簿      | Excel for Windows、Excel for Mac、Excel for iPad 和 Excel Online           |

> [!NOTE]
> `Name` 属性指定可以运行加载项的 Office 主机应用。Office 主机受不同平台支持，且可在台式机、Web 浏览器、平板电脑和移动设备上运行。不能指定用于运行加载项的平台。例如，如果指定 `Mailbox`，那么 Outlook 和 Outlook Web App 都可以用来运行加载项。 


## <a name="set-the-requirements-element-in-the-manifest"></a>在清单中设置 Requirements 元素

**Requirements** 元素指定运行外接程序时 Office 主机需要支持的最小要求集或 API 成员。**Requirements** 元素可以指定要求集和外接程序中使用的各个方法。在 1.1 版外接程序清单架构中，除 Outlook 外接程序外，**Requirements** 元素对于所有外接程序均为可选项。

> [!WARNING]
> **Requirements** 元素只能用于指定加载项必须使用的关键要求集或 API 成员。如果 Office 主机或平台不支持在 **Requirements** 元素中指定的要求集或 API 成员，加载项将无法在相应主机或平台上运行，并且不会显示在“我的加载项”中。相反，建议让加载项适用于 Office 主机的所有平台，如 Excel for Windows、Excel Online 和 Excel for iPad。若要让加载项适用于_所有_ Office 主机和平台，请使用运行时检查，而不是 **Requirements** 元素。

下面的代码示例展示了在支持以下内容的所有 Office 主机应用中加载的加载项：

-  最低版本为 1.1 的 **TableBindings** 要求集。
    
-  最低版本为 1.1 的 **OOXML** 要求集。
    
-  **Document.getSelectedDataAsync** 方法。

```XML
<Requirements>
   <Sets DefaultMinVersion="1.1">
      <Set Name="TableBindings" MinVersion="1.1"/>
      <Set Name="OOXML" MinVersion="1.1"/>
   </Sets>
   <Methods>
      <Method Name="Document.getSelectedDataAsync"/>
   </Methods>
</Requirements>
```

- **Requirements** 元素包含 **Sets** 和 **Methods** 子元素。
    
- **Sets** 元素可以包含一个或多个 **Set** 元素。**DefaultMinVersion** 指定所有子 **Set** 元素的默认 **MinVersion** 值。
    
- **Set** 元素指定 Office 主机必须支持的用以外接程序的要求集。 **Name** 属性指定要求集名称。 **MinVersion** 指定最低要求集版本。 **MinVersion** 将覆盖 **DefaultMinVersion** 的值。有关您的 API 成员归属的要求集和要求集版本的详细信息，请参阅 [Office 外接程序要求集](https://dev.office.com/reference/add-ins/office-add-in-requirement-sets)。
    
- **Methods** 元素可以包含一个或多个 **Method** 元素。不能将 **Methods** 元素和 Outlook 外接程序结合使用。
    
- **Method** 元素指定在您的外接程序所运行 Office 主机中必须要支持的单个方法。 **Name** 属性为必需属性，并使用其父对象指定合格方法的名称。
    

## <a name="use-runtime-checks-in-your-javascript-code"></a>使用运行时签入您的 JavaScript 代码


如果 Office 主机支持某些要求集，你可能想要在你的外接程序中提供其他功能。例如，如果外接程序在 Word 2016 中运行，你可能想要在现有的外接程序中使用新的 Word JavaScript API Word。若要执行此操作，请结合要求集名称使用 **isSetSupported** 方法。**isSetSupported** 确定在运行时，运行外接程序的 Office 主机是否支持要求集。如果要求集受支持，则 **isSetSupported** 返回 **true** 并运行使用此要求集中 API 成员的其他代码。如果 Office 主机不支持此要求集，则 **isSetSupported** 返回 **false** 且不会运行其他代码。以下代码显示与 **isSetSupported** 结合使用的语法。


```js
if (Office.context.requirements.isSetSupported(RequirementSetName , VersionNumber))
{
   // Code that uses API members from RequirementSetName.
}

```


-  _RequirementSetName_（必填）是代表该要求集名称的字符串。有关可用要求集的详细信息，请参阅 [Office 外接程序要求集](https://dev.office.com/reference/add-ins/office-add-in-requirement-sets)。
    
-  _VersionNumber_（可选）是要求集的版本。
    
在 Excel 2016 或 Word 2016 中，将 **isSetSupported** 和 **ExcelAPI** 或 **WordAPI** 要求集结合使用。**isSetSupported** 方法和 **ExcelAPI** 及 **WordAPI** 要求集可从 CDN 的最新 Office.js 文件中获取。如果未使用 CDN 中的 Office.js，则外接程序可能产生异常，因为 **isSetSupported** 将属于未定义的内容。有关详细信息，请参阅 [指定最新的适用于 Office 的 JavaScript API 库](#specify-the-latest-javascript-api-for-office-library)。 


> [!NOTE]
> **isSetSupported** 不适用于 Outlook 或 Outlook Web App。若要在 Outlook 或 Outlook Web App 中使用运行时检查，请利用[使用不属于要求集的方法的运行时检查](#runtime-checks-using-methods-not-in-a-requirement-set)中所述的技术。

以下代码示例演示外接程序如何向支持不同要求集或 API 成员的不同 Office 主机提供不同功能。




```js
if (Office.context.requirements.isSetSupported('WordApi', 1.1))
{
    // Run code that provides additional functionality using the JavaScript API for Word when the add-in runs in Word 2016.
}
else if (Office.context.requirements.isSetSupported('CustomXmlParts'))
{
      // Run code that uses API members from the CustomXmlParts requirement set.
}
else 
{
    // Run additional code when the Office host is not Word 2016, and when the Office host does not support the CustomXmlParts requirement set.
}

```


## <a name="runtime-checks-using-methods-not-in-a-requirement-set"></a>使用不属于要求集的方法的运行时检查


部分 API 成员不属于要求集这仅适用于属于 [适用于 Office 的 JavaScript API](https://dev.office.com/reference/add-ins/javascript-api-for-office) 命名空间的 API 成员（Office 下的任何内容），而不适用于属于 Word JavaScript API（Word 中的任何内容）或 [Excel 外接程序 JavaScript API 引用](https://dev.office.com/reference/add-ins/excel/excel-add-ins-reference-overview)（Excel 中的任何内容）命名空间的 API 成员。当外接程序依赖于某个不属于要求集的方法时，可以使用运行时检查来确定 Office 主机是否支持此方法，方法如以下代码示例所示。有关不属于要求集的方法的完整列表，请参阅 [Office 外接程序要求集](https://dev.office.com/reference/add-ins/office-add-in-requirement-sets)。


> [!NOTE]
> 建议限制在加载项代码中使用此类型运行时检查。

下面的代码示例检查主机是否支持 **document.setSelectedDataAsync**。




```js
if (Office.context.document.setSelectedDataAsync)
{
    // Run code that uses document.setSelectedDataAsync.
}
```


## <a name="see-also"></a>另请参阅

- [Office 加载项 XML 清单](add-in-manifests.md)
- [Office 外接程序要求集](https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets)
- [Word-Add-in-Get-Set-EditOpen-XML](https://github.com/OfficeDev/Word-Add-in-Get-Set-EditOpen-XML)