---
title: 随文档自动打开任务窗格
description: 了解如何将 Office 加载项配置为在文档打开时自动打开。
ms.date: 09/01/2022
ms.localizationpriority: medium
ms.openlocfilehash: 125e6bcccceb9fe0ced6992ba04a954695235ed4
ms.sourcegitcommit: 3abcf7046446e7b02679c79d9054843088312200
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/02/2022
ms.locfileid: "68810188"
---
# <a name="automatically-open-a-task-pane-with-a-document"></a>随文档自动打开任务窗格

可以通过向 Office 应用功能区添加按钮，在 Office 外接程序中使用加载项命令来扩展 Office UI。 当用户单击命令按钮时，会执行一个操作，如打开任务窗格。

某些情况下，需要在文档打开时自动打开一个任务窗格，而无需进行显式用户交互。 可以使用 [AddInCommands 1.1 要求集中](/javascript/api/requirement-sets/common/add-in-commands-requirement-sets)引入的自动打开任务窗格功能，在方案需要时自动打开任务窗格。

> [!NOTE]
> 若要将任务窗格配置为在安装加载项时立即打开，但不一定在以后打开文档时打开，请参阅 [在安装外接程序时自动打开任务窗格](automatically-open-on-installation.md)。

## <a name="how-is-the-autoopen-feature-different-from-inserting-a-task-pane"></a>AutoOpen 功能与插入任务窗格有何不同？

如果用户启动不使用外接程序命令的外接程序（例如，在 Office 2013 中运行的外接程序），外接程序会插入并保留在文档中。 因此，当其他用户打开文档时，系统会提示他们安装外接程序，随后会打开任务窗格。 此模型的挑战在于，在许多情况下，用户不希望外接程序保留在文档中。 例如，在 Word 文档中使用字典外接的学生可能不希望系统他们的同学或老师在打开该文档时提示他们安装该外接程序。

使用 Autoopen 功能，可以显式定义或允许用户定义特定任务窗格外接程序是否保留在特定文档中。

## <a name="support-and-availability"></a>支持和可用性

以下产品和平台当前支持自动打开功能。

|产品|平台|
|:-----------|:------------|
|<ul><li>Word</li><li>Excel</li><li>PowerPoint</li></ul>|所有受支持产品支持的平台：<ul><li>Windows 桌面上的 Office。 内部版本 16.0.8121.1000+</li><li>Office on Mac. Build 15.34.17051500+</li><li>Office 网页版</li></ul>|

## <a name="best-practices"></a>最佳做法

使用自动打开功能时，请应用以下最佳做法。

- 当 Autoopen 功能可帮助外接程序用户工作更高效时使用此功能，如：
  - When the document needs the add-in in order to function properly. For example, a spreadsheet that includes stock values that are periodically refreshed by an add-in. The add-in should open automatically when the spreadsheet is opened to keep the values up to date.
  - When the user will most likely always use the add-in with a particular document. For example, an add-in that helps users fill in or change data in a document by pulling information from a backend system.
- Allow users to turn on or turn off the autoopen feature. Include an option in your UI for users to choose to no longer automatically open the add-in task pane.  
- 使用要求集检测来确定自动打开功能是否可用，如果没有，则提供回退行为。
- 不要使用 Autoopen 功能来人为地增加外接程序的使用率。 如果加载项使用某些文档自动打开没有意义，此功能可能会惹恼用户。

    > [!NOTE]
    > 如果 Microsoft 检测到滥用 AutoOpen 功能，加载项可能会从 AppSource 下架。

- Don't use this feature to pin multiple task panes. You can only set one pane of your add-in to open automatically with a document.  

## <a name="implement-the-autoopen-feature"></a>实现自动打开功能

- 指定要自动打开的任务窗格。
- 标记要自动打开任务窗格的文档。

> [!IMPORTANT]
> The pane that you designate to open automatically will only open if the add-in is already installed on the user's device. If the user does not have the add-in installed when they open a document, the autoopen feature will not work and the setting will be ignored. If you also require the add-in to be distributed with the document you need to set the visibility property to 1; this can only be done using OpenXML, an example is provided later in this article.

### <a name="step-1-specify-the-task-pane-to-open"></a>第 1 步：指定要打开的任务窗格

To specify the task pane to open automatically, set the [TaskpaneId](/javascript/api/manifest/action#taskpaneid) value to **Office.AutoShowTaskpaneWithDocument**. You can only set this value on one task pane. If you set this value on multiple task panes, the first occurrence of the value will be recognized and the others will be ignored.

在下面的示例中，TaskPaneId 值设置为 Office.AutoShowTaskpaneWithDocument。

```xml
<Action xsi:type="ShowTaskpane">
    <TaskpaneId>Office.AutoShowTaskpaneWithDocument</TaskpaneId>
    <SourceLocation resid="Contoso.Taskpane.Url" />
</Action>
```

### <a name="step-2-tag-the-document-to-automatically-open-the-task-pane"></a>第 2 步：将文档标记为自动打开任务窗格

You can tag the document to trigger the autoopen feature in one of two ways. Pick the alternative that works best for your scenario.  

#### <a name="tag-the-document-on-the-client-side"></a>在客户端上标记文档

使用 Office.js [settings.set](/javascript/api/office/office.settings) 方法将 **Office.AutoShowTaskpaneWithDocument** 设置为 `true`，如以下示例所示。

```js
Office.context.document.settings.set("Office.AutoShowTaskpaneWithDocument", true);
Office.context.document.settings.saveAsync();
```

如果需要将文档标记为外接程序交互的一部分（例如，在用户创建一个绑定，或选择一个选项来表示他们希望窗格自动打开时），则使用此方法。

#### <a name="use-open-xml-to-tag-the-document"></a>使用 Open XML 标记文档

You can use Open XML to create or modify a document and add the appropriate Open Office XML markup to trigger the autoopen feature. For a sample that shows you how to do this, see [Office-OOXML-EmbedAddin](https://github.com/OfficeDev/Office-OOXML-EmbedAddin).

向文档添加两个 Open XML 部件。

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

|`storeType` 值|`id` 值|`store` 值|`version` 值|
|:---------------|:---------------|:---------------|:---------------|
|OMEX (AppSource)|加载项的 AppSource 资产 ID (请参阅注释) 。|AppSource 的区域设置；例如，“en-us”。|AppSource 目录中的版本 (请参阅注释) 。|
|WOPICatalog (合作伙伴 [WOPI](/microsoft-365/cloud-storage-partner-program/online/) 主机) | 加载项的 AppSource 资产 ID (请参阅注释) 。 | “wopicatalog”。 对于在应用源中发布并安装在 WOPI 主机中的加载项，请使用此值。 有关详细信息，请参阅 [与 Office Online 集成](/microsoft-365/cloud-storage-partner-program/online/overview)。 | 加载项清单中的版本。|
|FileSystem（网络共享）|外接程序清单中外接程序的 GUID。|网络共享路径。例如，“\\\\MyComputer\\MySharedFolder”。|外接程序清单中的版本。|
|EXCatalog（通过 Exchange 服务器部署） |外接程序清单中外接程序的 GUID。|“EXCatalog”。 EXCatalog 行是用于在Microsoft 365 管理中心中使用集中部署的加载项的行。|外接程序清单中的版本。|
|Registry（系统注册表）|外接程序清单中外接程序的 GUID。|“developer”|加载项清单中的版本。|

> [!NOTE]
> To find the asset ID and version of an add-in in AppSource, go to the AppSource landing page for the add-in. The asset ID appears in the address bar in the browser. The version is listed in the **Details** section of the page.

若要详细了解 webextension 标记，请参阅 [[MS-OWEXML] 2.2.5. WebExtensionReference](/openspecs/office_standards/ms-owexml/d4081e0b-5711-45de-b708-1dfa1b943ad1)。

以下示例演示如何添加 `taskpane` 部件。

```xml
<wetp:taskpane dockstate="right" visibility="0" width="350" row="4" xmlns:wetp="http://schemas.microsoft.com/office/webextensions/taskpanes/2010/11">
  <wetp:webextensionref xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1" />
</wetp:taskpane>
```

请注意，在本例中，`visibility` 属性设置为“0”。 这意味着在添加 webextension 部件和 `taskpane` 部件之后，第一次打开文档时，用户还必须从功能区上的“外接程序”按钮安装该外接程序。 此后，外接程序任务窗格将在打开该文件时自动打开。 此外，在将 `visibility` 设置为“0”时，可以使用 Office.js 让用户打开或关闭 AutoOpen 功能。 具体来说，脚本会将 **Office.AutoShowTaskpaneWithDocument** 文档设置为 `true` 或 `false`。 （有关详细信息，请参阅[在客户端上标记文档](#tag-the-document-on-the-client-side)。）

If `visibility` is set to "1", the task pane opens automatically the first time the document is opened. The user is prompted to trust the add-in, and when trust is granted, the add-in opens. Thereafter, the add-in task pane opens automatically when the file is opened. However, when `visibility` is set to "1", you can't use Office.js to enable users to turn on or turn off the autoopen feature.

当外接程序和模板或文档内容紧密集成以致用户不会选择退出 Autoopen 功能时，将 `visibility` 设置为“1”是一个不错的选择。

> [!NOTE]
> If you want to distribute your add-in with the document, so that users are prompted to install it, you must set the visibility property to 1. You can only do this via Open XML.

编写 XML 的一种简单方法是首先运行加载项， [并在客户端标记文档](#tag-the-document-on-the-client-side) 以写入值，然后保存文档并检查生成的 XML。Office 将检测并提供相应的属性值。 还可以使用 [Open XML SDK Productivity Tool](https://www.nuget.org/packages/Open-XML-SDK) 生成 C# 代码，以编程方式添加基于生成的 XML 的标记。

## <a name="test-and-verify-opening-task-panes"></a>对打开任务窗格进行测试和验证

可以部署外接程序的测试版本，该版本将通过 Microsoft 365 管理中心使用集中部署自动打开任务窗格。 以下示例演示如何使用 EXCatalog 应用商店版本从集中部署目录插入外接程序。

```xml
<we:webextension xmlns:we="http://schemas.microsoft.com/office/webextensions/webextension/2010/11" id="{52811C31-4593-43B8-A697-EB873422D156}">
    <we:reference id="af8fa5ba-4010-4bcc-9e03-a91ddadf6dd3" version="1.0.0.0" store="EXCatalog" storeType="EXCatalog"/>
    <we:alternateReferences/>
    <we:properties/>
    <we:bindings/>
    <we:snapshot xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
</we:webextension>
```

可以使用 Microsoft 365 订阅来测试前面的示例，以尝试集中部署并验证加载项是否按预期工作。 如果还没有 Microsoft 365 订阅，可以通过加入 Microsoft 365 开发人员计划来获取免费的 90 天可续订 [Microsoft 365](https://developer.microsoft.com/office/dev-program) 订阅。

## <a name="see-also"></a>另请参阅

- 有关演示如何使用 AutoOpen 功能的示例，请参阅 [Office 外接程序命令示例](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/tree/master/AutoOpenTaskpane)。
- [安装加载项时自动打开任务窗格](automatically-open-on-installation.md)
- [加入 Microsoft 365 开发人员计划。](/office/developer-program/office-365-developer-program)