---
title: 在加载项中请求获取 API 使用权限
description: 了解在内容或任务窗格外接程序的清单中声明的不同权限级别，以指定 JavaScript API 访问级别。
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: bdc23d1d8d5ff044a14306a868567864865c012d
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/23/2021
ms.locfileid: "53075969"
---
# <a name="requesting-permissions-for-api-use-in-add-ins"></a>在加载项中请求获取 API 使用权限

本文说明您可以在内容或任务窗格加载项清单中声明的不同权限级别，以指定加载项功能所需的 JavaScript API 访问的级别。 

## <a name="permissions-model"></a>权限模型

5 级 JavaScript API 访问权限模型为内容和任务窗格加载项的用户提供基本的隐私和安全功能。图 1 显示您可以在加载项清单中声明的 API 权限的 5 个级别。

*图 1：内容和任务窗格加载项的 5 级权限模型*

![任务窗格应用程序的权限级别。](../images/office15-app-sdk-task-pane-app-permission.png)

这些权限指定加载项运行时在用户插入然后激活（信任）加载项时允许内容或任务窗格加载项使用的 API 子集。若要声明内容或任务窗格加载项所需的权限级别，请在加载项清单的 [Permissions](../reference/manifest/permissions.md) 元素中指定任一权限文本值。以下示例要求 **WriteDocument** 权限，仅允许可以对文档进行写入（而非阅读）的方法。

```XML
<Permissions>WriteDocument</Permissions>
```

作为最佳做法，应该根据 _最小特权_ 原则请求权限。也就是说，应该请求仅可访问加载项正常运行所需的 API 最小子集的权限。例如，如果您加载项的功能只需要读取用户文档中的数据，应该请求的权限不应高于 **ReadDocument** 权限。

下表描述了每个权限级别启用的 JavaScript API 子集。

|**权限**|**启用的 API 子集**|
|:-----|:-----|
|**受限**|[Settings](/javascript/api/office/office.settings) 对象的方法和 [Document.getActiveViewAsync](/javascript/api/office/office.document#getactiveviewasync-options--callback-) 方法。这是内容或任务窗格加载项可以请求的最低级别权限。|
|**ReadDocument**|除了 Restricted 权限所允许的API 之外，还添加了对读取文档和管理绑定所需的 API 成员的访问权限。这包括使用：<br/><ul><li>
  <a href="/javascript/api/office/office.document#getselecteddataasync-coerciontype--options--callback-" target="_blank">Document.getSelectedDataAsync</a> 方法，用于获取所选文本、HTML（仅限 Word）或表格数据，但不可用于包含文档中所有数据的基础 Open Office XML (OOXML) 代码。</p></li><li><p><a href="/javascript/api/office/office.document#getfileasync-filetype--options--callback-" target="_blank">Document.getFileAsync</a> 方法，用于获取文档中的所有文本，而不是文档的基础 OOXML 二进制副本。</p></li><li><p><a href="/javascript/api/office/office.binding#getdataasync-options--callback-" target="_blank">Binding.getDataAsync</a> 方法，用于读取文档中的绑定数据。</p></li><li><p><a href="/javascript/api/office/office.bindings#addfromnameditemasync-itemname--bindingtype--options--callback-" target="_blank">Bindings</a> 对象的 <a href="/javascript/api/office/office.bindings#addfrompromptasync-bindingtype--options--callback-" target="_blank">addFromNamedItemAsync</a>、<a href="/javascript/api/office/office.bindings#addfromselectionasync-bindingtype--options--callback-" target="_blank">addFromPromptAsync</a>、<span class="keyword">addFromSelectionAsync</span> 方法，用于在文档中创建绑定。</p></li><li><p><a href="/javascript/api/office/office.bindings#getallasync-options--callback-" target="_blank">Bindings</a> 对象的 <a href="/javascript/api/office/office.bindings#getbyidasync-id--options--callback-" target="_blank">getAllAsync</a>、<a href="/javascript/api/office/office.bindings#releasebyidasync-id--options--callback-" target="_blank">getByIdAsync</a> 和 <span class="keyword">releaseByIdAsync</span> 方法，用于访问和删除文档中的绑定。</p></li><li><p><a href="/javascript/api/office/office.document#getfilepropertiesasync-options--callback-" target="_blank">Document.getFilePropertiesAsync</a> 方法，用于访问文档文件属性，例如文档的 URL。</p></li><li><p><a href="/javascript/api/office/office.document#gotobyidasync-id--gototype--options--callback-" target="_blank">Document.goToByIdAsync</a> 方法，用于导航到文档中的已命名对象和位置。</p></li><li><p>对于项目的任务窗格外接程序，<a href="/javascript/api/office/office.document" target="_blank">ProjectDocument</a> 对象的所有"get"方法。 </p></li></ul>|
|**ReadAllDocument**|除了 Restricted 和 **ReadDocument** 权限所允许的 API 之外，还允许以下附加访问文档数据： <br/><ul><li><p><span class="keyword">Document.getSelectedDataAsync</span> 和 <span class="keyword">Document.getFileAsync</span> 方法可以访问文档（文档中除了文本，还可能包含格式、链接、嵌入图片、注释、修订等）的基础 OOXML 代码。</p></li></ul>|
|**WriteDocument**|除了 Restricted 权限允许的API 之外，还添加了对以下 API 成员的访问权限：<br/><ul><li><p><a href="/javascript/api/office/office.document#setselecteddataasync-data--options--callback-" target="_blank">Document.setSelectedDataAsync</a> 方法，用于在文档中写入用户所选内容。</p></li></ul>|
|**ReadWriteDocument**|除了 Restricted、ReadDocument、ReadAllDocument 和 **WriteDocument** 权限所允许的 API 之外，还包括对内容和任务窗格加载项支持的所有剩余 API 的访问权限，包括订阅事件的方法。  必须声明 **ReadWriteDocument** 权限才能访问这些额外的 API 成员：<br/><ul><li><p><a href="/javascript/api/office/office.binding#setdataasync-data--options--callback-" target="_blank">Binding.setDataAsync</a> 方法，用于将内容写入到文档的绑定区域。</p></li><li><p><a href="/javascript/api/office/office.tablebinding#addrowsasync-rows--options--callback-" target="_blank">TableBinding.addRowsAsync</a> 方法，用于将行添加到绑定表格中。</p></li><li><p><a href="/javascript/api/office/office.tablebinding#addcolumnsasync-tabledata--options--callback-" target="_blank">TableBinding.addColumnsAsync</a> 方法，用于将列添加到绑定表格中。</p></li><li><p><a href="/javascript/api/office/office.tablebinding#deletealldatavaluesasync-options--callback-" target="_blank">TableBinding.deleteAllDataValuesAsync</a> 方法，用于删除绑定表格中的所有数据。</p></li><li><p><a href="/javascript/api/office/office.tablebinding#setformatsasync-cellformat--options--callback-" target="_blank">TableBinding</a> 对象的 <a href="/javascript/api/office/office.tablebinding#clearformatsasync-options--callback-" target="_blank">setFormatsAsync</a>、<a href="/javascript/api/office/office.tablebinding#settableoptionsasync-tableoptions--options--callback-" target="_blank">clearFormatsAsync</a> 和 <span class="keyword">setTableOptionsAsync</span> 方法，用于设置绑定表格中的格式和选项。</p></li><li><p><a href="/javascript/api/office/office.customxmlnode" target="_blank">CustomXmlNode</a>、<a href="/javascript/api/office/office.customxmlpart" target="_blank">CustomXmlPart</a>、<a href="/javascript/api/office/office.customxmlparts" target="_blank">CustomXmlParts</a> 和 <a href="/javascript/api/office/office.customxmlprefixmappings" target="_blank">CustomXmlPrefixMappings</a> 对象的所有成员。</p></li><li><p>内容和任务窗格加载项支持的所有订阅事件的方法，具体来说即 <span class="keyword">Binding</span>、<span class="keyword">CustomXmlPart</span>、<a href="/javascript/api/office/office.binding" target="_blank">Document</a>、<a href="/javascript/api/office/office.customxmlpart" target="_blank">ProjectDocument</a> 和 <a href="/javascript/api/office/office.document" target="_blank">Settings</a> 对象的 <a href="/javascript/api/office/office.document" target="_blank">addHandlerAsync</a> 和 <a href="/javascript/api/office/office.document#settings" target="_blank">removeHandlerAsync</a> 方法。</p></li></ul>|

## <a name="see-also"></a>另请参阅

- [Office 加载项的隐私和安全](../concepts/privacy-and-security.md)
