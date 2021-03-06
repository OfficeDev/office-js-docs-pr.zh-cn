---
title: 随文档自动打开任务窗格
description: 了解如何将 Office 外接程序配置为在文档打开时自动打开。
ms.date: 07/07/2020
localization_priority: Normal
ms.openlocfilehash: 9c9d64594f98cc45c56b3251bf6750457bd5a37c
ms.sourcegitcommit: d28392721958555d6edea48cea000470bd27fcf7
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/13/2021
ms.locfileid: "49839955"
---
# <a name="automatically-open-a-task-pane-with-a-document"></a>随文档自动打开任务窗格

可以通过向 Office 应用程序功能区添加按钮，在 Office 外接程序中使用外接程序命令来扩展 Office UI。 当用户单击命令按钮时，会执行一个操作，如打开任务窗格。

某些情况下，需要在文档打开时自动打开一个任务窗格，而无需进行显式用户交互。 可以使用 Addincommand 1.1 要求集中引入的 AutoOpen 任务窗格功能，以在情况需要时自动打开一个任务窗格。


## <a name="how-is-the-autoopen-feature-different-from-inserting-a-task-pane"></a>AutoOpen 功能与插入任务窗格有何不同？

如果用户启动不使用外接程序命令的外接程序（例如，在 Office 2013 中运行的外接程序），外接程序会插入并保留在文档中。 因此，当其他用户打开文档时，系统会提示他们安装外接程序，随后会打开任务窗格。 此模型的挑战在于，在许多情况下，用户不希望外接程序在文档中保留。 例如，在 Word 文档中使用字典外接的学生可能不希望系统他们的同学或老师在打开该文档时提示他们安装该外接程序。

使用 Autoopen 功能，可以显式定义或允许用户定义特定任务窗格外接程序是否保留在特定文档中。

## <a name="support-and-availability"></a>支持和可用性

目前支持在以下产品和平台中 <!-- in **developer preview** and it is only --> 使用 Autoopen 功能。

|**产品**|**平台**|
|:-----------|:------------|
|<ul><li>Word</li><li>Excel</li><li>PowerPoint</li></ul>|所有产品的支持平台：<ul><li>Windows 桌面版 Office（内部版本 16.0.8121.1000+）</li><li>Mac 版 Office（生成号 15.34.17051500 及更高版本）</li><li>Office 网页版</li></ul>|


## <a name="best-practices"></a>最佳做法

在使用 Autoopen 功能时应用下面的最佳做法：

- 当 Autoopen 功能可帮助外接程序用户工作更高效时使用此功能，如：
  - 当文档需要外接程序才能正常工作时。例如，包括由外接程序定期刷新的股票值的电子表格。外接程序应在电子表格打开时自动打开，以保持值处于最新状态。
  - 当用户很可能始终将外接程序与某个特定文档一同使用时。例如，可帮助用户通过从后台系统中获取信息来填写或更改文档中数据的外接程序。
- 允许用户打开或关闭 Autoopen 功能。用户可以选择 UI 中包含的一个选项来停止自动打开外接程序任务窗格。  
- 使用要求集检测确定 Autoopen 功能是否可用，如果不可用，则提供回退行为。
- 不要使用 Autoopen 功能来人为地增加外接程序的使用率。 如果外接程序无法自动打开某些文档，则此功能可能会令用户厌烦。

    > [!NOTE]
    > 如果 Microsoft 检测到滥用 AutoOpen 功能，加载项可能会从 AppSource 下架。

- 请勿使用此功能来固定多个任务窗格。只能设置一个外接程序窗格随文档自动打开。  

## <a name="implementation"></a>实现

要实现 Autoopen 功能，请执行以下操作：

- 指定要自动打开的任务窗格。
- 标记要自动打开任务窗格的文档。

> [!IMPORTANT]
> 只有在用户设备上已安装加载项时，才能打开指定为自动打开的窗格。如果在打开文档时用户未安装加载项，那么 AutoOpen 功能将不起作用，而且设置也会被忽略。如果还要求加载项与文档一起分发，需要将“visibility”属性设置为 1；只能使用 OpenXML 完成此操作，本文稍后将提供示例。

### <a name="step-1-specify-the-task-pane-to-open"></a>第 1 步：指定要打开的任务窗格

若要指定要自动打开的任务窗格，请将 [TaskpaneId](../reference/manifest/action.md#taskpaneid) 值设置为 **Office.AutoShowTaskpaneWithDocument**。只能在一个任务窗格上设置此值。如果在多个任务窗格上设置此值，将识别值的第一个匹配项，而忽略其他。

在下面的示例中，TaskPaneId 值设置为 Office.AutoShowTaskpaneWithDocument。

```xml
<Action xsi:type="ShowTaskpane">
    <TaskpaneId>Office.AutoShowTaskpaneWithDocument</TaskpaneId>
    <SourceLocation resid="Contoso.Taskpane.Url" />
</Action>
```

### <a name="step-2-tag-the-document-to-automatically-open-the-task-pane"></a>第 2 步：将文档标记为自动打开任务窗格

可以通过下面的两种方法之一，将文档标记为触发自动打开功能。 选择最适合自己应用场景的备选方法。  


#### <a name="tag-the-document-on-the-client-side"></a>在客户端上标记文档

使用 Office.js [settings.set](/javascript/api/office/office.settings) 方法将 **Office.AutoShowTaskpaneWithDocument** 设置为“**true**”，如以下示例所示。

```js
Office.context.document.settings.set("Office.AutoShowTaskpaneWithDocument", true);
Office.context.document.settings.saveAsync();
```

如果需要将文档标记为外接程序交互的一部分（例如，在用户创建一个绑定，或选择一个选项来表示他们希望窗格自动打开时），则使用此方法。

#### <a name="use-open-xml-to-tag-the-document"></a>使用 Open XML 标记文档

可以使用 Open XML 来创建或修改文档，并添加适当的 Open Office XML 标记来触发 Autoopen 功能。有关演示如何执行此操作的示例，请参阅 [Office-OOXML-EmbedAddin](https://github.com/OfficeDev/Office-OOXML-EmbedAddin)。

向文档添加两个 Open XML 部件：

- 一个 `webextension` 部件
- 一个 `taskpane` 部件

以下示例演示如何添加 `webextension` 部件。

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

`webextension` 部件包含一个属性包，以及必须设置为 `true` 的 **Office.AutoShowTaskpaneWithDocument** 属性。

`webextension` 部件还包括对具有 `id`、`storeType`、`store` 和 `version` 的属性的应用商店或目录的引用。 在 `storeType` 值中，只有四个与 AutoOpen 功能相关。 其他三个属性的值取决于 `storeType` 的值，如下表所示。

| **`storeType` 值** | **`id` 值**    |**`store` 值** | **`version` 值**|
|:---------------|:---------------|:---------------|:---------------|
|OMEX (AppSource)|加载项的 AppSource 资产 ID（请参阅“注意”）|AppSource 的区域设置；例如，“en-us”。|AppSource 目录中的版本（请参阅“注意”）|
|FileSystem（网络共享）|外接程序清单中外接程序的 GUID。|网络共享路径。例如，“\\\\MyComputer\\MySharedFolder”。|外接程序清单中的版本。|
|EXCatalog（通过 Exchange 服务器部署） |外接程序清单中外接程序的 GUID。|“EXCatalog”。 EXCatalog 行是一行，用于 Microsoft 365 管理中心内使用集中部署的加载项。|外接程序清单中的版本。
|Registry（系统注册表）|外接程序清单中外接程序的 GUID。|“developer”|加载项清单中的版本。|

> [!NOTE]
> 若要查找 AppSource 中加载项的资产 ID 和版本，请转到加载项的 AppSource 登陆页面。资产 ID 显示在浏览器的地址栏中。版本在页面的“详细信息”部分中列出。

若要详细了解 webextension 标记，请参阅 [[MS-OWEXML] 2.2.5. WebExtensionReference](/openspecs/office_standards/ms-owexml/d4081e0b-5711-45de-b708-1dfa1b943ad1)。

以下示例演示如何添加 `taskpane` 部件。

```xml
<wetp:taskpane dockstate="right" visibility="0" width="350" row="4" xmlns:wetp="http://schemas.microsoft.com/office/webextensions/taskpanes/2010/11">
  <wetp:webextensionref xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1" />
</wetp:taskpane>
```

请注意，在本例中，`visibility` 属性设置为“0”。 这意味着在添加 webextension 部件和 `taskpane` 部件之后，第一次打开文档时，用户还必须从功能区上的“外接程序”按钮安装该外接程序。 此后，外接程序任务窗格将在打开该文件时自动打开。 此外，在将 `visibility` 设置为“0”时，可以使用 Office.js 让用户打开或关闭 AutoOpen 功能。 具体来说，脚本会将 **Office.AutoShowTaskpaneWithDocument** 文档设置为 `true` 或 `false`。 （有关详细信息，请参阅[在客户端上标记文档](#tag-the-document-on-the-client-side)。）

如果 `visibility` 设置为“1”，任务窗格将在文件第一次打开时自动打开。系统会提示用户信任该外接程序，授予信任后，将打开外接程序。此后，外接程序任务窗格将在打开该文件时自动打开。但是，当 `visibility` 设置为“1”时，则不能使用 Office.js 让用户打开或关闭 Autoopen 功能。

当外接程序和模板或文档内容紧密集成以致用户不会选择退出 Autoopen 功能时，将 `visibility` 设置为“1”是一个不错的选择。

> [!NOTE]
> 若要将加载项与文档一起分发，以便提示用户进行安装，必须将“visibility”属性设置为 1。只能通过 Open XML 执行此操作。

编写 XML 的一个简单方法是首先运行外接程序并[标记客户端上的文档](#tag-the-document-on-the-client-side)以写入值，然后保存该文档并检查生成的 XML。Office 将检测并提供适当的属性值。还可以使用 [Open XML SDK 2.5 Productivity Tool](https://www.microsoft.com/download/details.aspx?id=30425) 工具生成 C# 代码以编程方式添加基于生成的 XML 的标记。

## <a name="test-and-verify-opening-task-panes"></a>对打开任务窗格进行测试和验证

可以通过 Microsoft 365 管理中心部署将自动使用集中部署打开任务窗格的外接程序的测试版本。 以下示例演示如何使用 EXCatalog 应用商店版本从集中部署目录插入外接程序。

```xml
<we:webextension xmlns:we="http://schemas.microsoft.com/office/webextensions/webextension/2010/11" id="{52811C31-4593-43B8-A697-EB873422D156}">
    <we:reference id="af8fa5ba-4010-4bcc-9e03-a91ddadf6dd3" version="1.0.0.0" store="EXCatalog" storeType="EXCatalog"/>
    <we:alternateReferences/>
    <we:properties/>
    <we:bindings/>
    <we:snapshot xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
</we:webextension>
```

可以使用 Microsoft 365 订阅测试上一个示例，以试用集中部署并验证加载项是否按预期工作。 如果你还没有 Microsoft 365 订阅，可以通过加入 Microsoft 365 开发人员计划获得为期 90 天的免费可续订 [Microsoft 365 订阅](https://developer.microsoft.com/office/dev-program)。

## <a name="see-also"></a>另请参阅

有关演示如何使用 AutoOpen 功能的示例，请参阅 [Office 外接程序命令示例](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/tree/master/AutoOpenTaskpane)。
[加入 Microsoft 365 开发人员计划](/office/developer-program/office-365-developer-program)。