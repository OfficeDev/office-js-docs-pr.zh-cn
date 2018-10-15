---
title: 随文档自动打开任务窗格
description: ''
ms.date: 05/02/2018
ms.openlocfilehash: 2ebce1ce8bd95ee7802b5509d375f1986bb2877e
ms.sourcegitcommit: c53f05bbd4abdfe1ee2e42fdd4f82b318b363ad7
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/12/2018
ms.locfileid: "25505914"
---
# <a name="automatically-open-a-task-pane-with-a-document"></a>随文档自动打开任务窗格

可以使用 Office 加载项中的加载项指令， 通过将按钮加到 Office 功能区来扩展用户 UI ，当用户单击命令按钮时，会执行指定操作，如打开任务窗格。 

某些应用场景需要系统在文档打开时自动打开任务窗格，无需用户给予明确指令。可使用自动打开任务窗口功能，该功能在 AddInCommands 1.1 需求集中，可视情况自动打开任务窗格。 


## <a name="how-is-the-autoopen-feature-different-from-inserting-a-task-pane"></a>Autoopen 功能与插入任务窗格有什么不同？ 

如果用户启动不使用加载项命令的加载项（例如，在 Office 2013 中运行加载项），加载项会插入并保留在文档中。因此，当其他用户打开文档时，系统会提示他们安装加载项，随后会打开任务窗格。这种模型的问题是，通常用户不希望加载项保留在文档中。例如，在 Word 文档中使用字典插件的学生也许不希望他同学或老师在打开该文档时被提示安装该插件。  

使用 Autoopen 功能，可以明确定义或允许用户定义特定任务窗格加载项是否保留在特定文档中。 

## <a name="support-and-availability"></a>支持和有效性。
Autoopen 功能目前<!-- in **developer preview** and it is only -->在以下产品和平台中受支持。

|**产品**|**平台**|
|:-----------|:------------|
|<ul><li>Word</li><li>Excel</li><li>PowerPoint</li></ul>|所有产品支持的平台：<ul><li>Office for Windows Desktop（生成号 16.0.8121.1000+）</li><li>适用于 Mac 的 Office （生成号 15.34.17051500+）</li><li>Office Online</li></ul>|


## <a name="best-practices"></a>最佳做法

在使用 Autoopen 功能时建议如下操作：

- Autoopen 功能帮助加载项用户提高工作效率，如：
    - 当文档需要加载项才能正常工作时。例如，需要定期刷新的股票价格表，加载项应在文档打开始自动运行，以保证数据更新。 
    - 当用户使用某文档时，始终使用加载项。例如，某加载项可通过从后台系统提取信息来填写或更改文档中的数据。 
- 允许用户打开或关闭 Autoopen 功能。在 UI 中加入选项使用户可以选择 停止自动运行加载项任务窗口。  
- 使用需求集检测来判断 Autoopen 功能是否可用，若不可用则提供应变行为。
- 不要使用 Autoopen 功能来人为增加加载项的使用率。如果加载项在某些文档中自动无意义，则此功能就会打扰用户。 

    > [!NOTE]
    > 如果 Microsoft 检测到 Autoopen 功能被滥用，则加载项可能会从 AppSource 被迫下架。 

- 请勿使用此功能来固定多个任务窗格。每个文档只能自动打开一个加载项窗格。  

## <a name="implementation"></a>实施
要执行 Autoopen 功能：

- 指定要自动打开的任务窗格。
- 标记要自动打开任务窗格的文档。

> [!IMPORTANT]
> 只有在用户设备上已安装加载项时，被指定的窗格才回自动打开。如果在打开文档时用户未安装加载项，那么 Autoopen 功能将不起作用，并且设置也会被忽略。如果同时要求加载项与文档相对应，需要将 visibility 属性设置为 1；只能使用 OpenXML 完成此操作，本文稍后将提供示例。 

### <a name="step-1-specify-the-task-pane-to-open"></a>第 1 步：指定要打开的任务窗格
若要指定要自动打开的任务窗格，请将 [TaskpaneId](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/action?view=office-js#taskpaneid) 值设置为 **Office.AutoShowTaskpaneWithDocument**。只能在一个任务窗格上设置此值。如果在多个任务窗格上设置此值，将只识别第一个出现的值而忽略其他。 

在下面的示例中，TaskPaneId 的值被设置为 Office.AutoShowTaskpaneWithDocument。
          
```xml
<Action xsi:type="ShowTaskpane">
    <TaskpaneId>Office.AutoShowTaskpaneWithDocument</TaskpaneId>
    <SourceLocation resid="Contoso.Taskpane.Url" />
</Action>
```     

### <a name="step-2-tag-the-document-to-automatically-open-the-task-pane"></a>第 2 步：标记文档，自动打开任务窗格

可以通过如下两种方式触发 Autoopen 功能标记文档。选择最适合应用场景的方案。  


#### <a name="tag-the-document-on-the-client-side"></a>从客户端标记文档
使用 Office.js[settings.set](https://docs.microsoft.com/javascript/api/office/office.settings?view=office-js) 将 **Office.AutoShowTaskpaneWithDocument** 设置为**true**，如以下示例。   

```js
Office.context.document.settings.set("Office.AutoShowTaskpaneWithDocument", true);
Office.context.document.settings.saveAsync();
```

使用此方法将文档标记为加载项交互的一部分（例如，一旦用户创建了绑定，或选择自动弹出窗格）。

#### <a name="use-open-xml-to-tag-the-document"></a>使用 Open XML 标记文档
可以使用 Open XML 来创建或修改文档，并添加适当的 Open Office XML 标记来触发 Autoopen 功能。有关演示示例，请参阅 [Office-OOXML-EmbedAddin](https://github.com/OfficeDev/Office-OOXML-EmbedAddin)。 

向文档添加两个 Open XML 部件：

- Webextension 部件
- 任务窗格部件

以下示例演示如何添加 Webextension 部件。

```xml
<we:webextension xmlns:we="http://schemas.microsoft.com/office/webextensions/webextension/2010/11" id="[ADD-IN ID PER MANIFEST]">
  <we:reference id="[GUID or AppSource asset ID]" version="[your add-in version]" store="[Pointer to store or catalog]" storeType="[Store or catalog type]"/>
  <we:alternateReferences/>
  <we:properties>
    <we:property name="Office.AutoShowTaskpaneWithDocument" value="true"/>
  </we:properties>
  <we:bindings/>
  <we:snapshot xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
</we:webextension>
```

Webextension 部件包含一个属性包，必须将 **Office.AutoShowTaskpaneWithDocument** 属性设置为 `true` 。

Webextension 部件也包括对有 `id`、 `storeType`、 `store` 和 `version` 属性的应用商店或目录的引用。在 `storeType` 值中，只有四个值与 Autoopen 功能相关。其他三个属性的值取决于 `storeType` 的值，如下表所示。 

| **`storeType` 值** | **`id` 值**    |**`store` 值** | **`version` 值**|
|:---------------|:---------------|:---------------|:---------------|
|OMEX (AppSource)|加载项的 AppSource asset ID（请参阅 Note ）|AppSource 的区域；如，“en-us”。|AppSource 目录中的版本（请参阅 Note ）|
|FileSystem（网络共享）|加载项清单中加载项的 GUID 。|网络共享路径。例如，“\\\\MyComputer\\MySharedFolder”。|加载项清单中的版本。|
|EXCatalog（通过交换服务器部署） |加载项清单中加载项的 GUID 。|“EXCatalog”。EXCatalog 行是用于同加载项一同使用的行，该加载项使用 Office 365 管理中心的集中部署。|加载项清单中的版本。
|注册（系统注册）|加载项清单中加载项的 GUID 。|“开发者”|加载项清单中的版本。|

> [!NOTE]
> 若要查找 AppSource 的资产 ID 和加载项版本，请转到加载项的 AppSource 登陆页面。资产 ID 显示在浏览器的地址栏中。版本在页面的 **Details** 部分中列出。

若要详细了解 Webextension 标记，请参阅 [[MS-OWEXML] 2.2.5. WebExtensionReference](https://msdn.microsoft.com/library/hh695383(v=office.12).aspx)。

以下示例演示如何添加 Taskpane 部件。

```xml
<wetp:taskpane dockstate="right" visibility="0" width="350" row="4" xmlns:wetp="http://schemas.microsoft.com/office/webextensions/taskpanes/2010/11">
  <wetp:webextensionref xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1" />
</wetp:taskpane>
```

请注意，在本例中，`visibility` 属性设置为“0”。这表示在添加 Webextension 部件和任务窗格部件之后，第一次打开文档时，用户还须向功能区安装 **Add-in** 按钮的加载项。进而，加载项任务窗格将在打开该文件时自动弹出。此外，在将 `visibility` 设置为“0”时，可以使用 Office.js 允许用户打开或关闭 Autoopen 功能。具体来说，脚本会将 **Office.AutoShowTaskpaneWithDocument** 文档设置为 `true` 或 `false`。（有关详细信息，请参阅[在客户端上标记文档](#tag-the-document-on-the-client-side)。） 

如果 `visibility` 设置为“1”，任务窗格将在文件第一次打开时自动打开。系统会提示用户授权该加载项，授权后，将打开加载项。进而，加载项任务窗格将在打开该文件时自动打开。但是，当 `visibility` 设置为“1”时，则不能使用 Office.js 让用户打开或关闭 Autoopen 功能。 

当加载项和模板或文档内容紧密联系，用户不会选择退出 Autoopen 功能时，最好将 `visibility` 设置为“1”。 

> [!NOTE]
> 若要将加载项与文档一起发布，以便提示用户进行安装，必须将 Visibility 属性设置为 1。只能通过 Open XML 执行此操作。

编写 XML 的一个简单方法是首先运行加载项并使用[在客户端上标记文档](#tag-the-document-on-the-client-side)写入值，然后保存该文档并检查生成的 XML。Office 将检测并提供适当的属性值。还可以使用 [Open XML SDK 2.5 Productivity Tool](https://www.microsoft.com/download/details.aspx?id=30425) 工具生成 C# 代码以编程方式添加已生成的 XML 的标记。

## <a name="test-and-verify-opening-taskpanes"></a>测试并验证打开任务窗格
可以配置加载项的测试版本，加载项会使用 Office 365 管理中心的集中部署自动打开任务窗格。下面的示例演示如何从集中部署目录使用 EXCatalog 商店版本插入加载项。

```xml
<we:webextension xmlns:we="http://schemas.microsoft.com/office/webextensions/webextension/2010/11" id="{52811C31-4593-43B8-A697-EB873422D156}">
    <we:reference id="af8fa5ba-4010-4bcc-9e03-a91ddadf6dd3" version="1.0.0.0" store="EXCatalog" storeType="EXCatalog"/>
    <we:alternateReferences/>
    <we:properties/>
    <we:bindings/>
    <we:snapshot xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
</we:webextension>
```
若要测试上一例子，请考虑加入 [Office 365 开发者计划](https://docs.microsoft.com/office/developer-program/office-365-developer-program) 并注册 [Office 365 开发者帐户](https://developer.microsoft.com/office/dev-program) 如果还未订阅 Office 365 的话，可以对加载项是否如预期工作进行集中部署测试.


## <a name="see-also"></a>另请参阅

有关演示如何使用 Autoopen 功能的示例，请参阅  [Office 加载项命令示例](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/tree/master/AutoOpenTaskpane). [加入 Office 365 开发者计划](https://docs.microsoft.com/office/developer-program/office-365-developer-program). 

