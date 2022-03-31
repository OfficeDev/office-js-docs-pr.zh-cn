---
title: 指定 Office 主机和 API 要求
description: 了解如何指定Office应用和 API 要求，使加载项按预期运行。
ms.date: 01/26/2022
ms.localizationpriority: medium
ms.openlocfilehash: 9f08a4c5f52d52022b33285faf3d3914056a03e2
ms.sourcegitcommit: f32123f2b7254e76965dc95c21108f081507feed
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/30/2022
ms.locfileid: "64536543"
---
# <a name="specify-office-applications-and-api-requirements"></a>指定 Office 应用程序和 API 要求

Office加载项可能依赖于特定的 Office 应用程序 (也称为 Office 主机) 或 Office JavaScript API (office.js) 。 例如，你的外接程序可能：

- 在单个 Office 应用程序（如，Word 或 Excel）或多个应用程序中运行。
- 使用仅在 Office某些版本的 JavaScript API Office。 例如，Excel 2016的一次购买版本不支持 Excel JavaScript 库中Office API。

在这些情况下，需要确保你的外接程序永远不会安装在Office或Office无法运行的版本中。

还有一些方案，您希望根据用户的外接程序应用程序和版本来控制哪些外接程序功能Office用户Office。 两个示例是：

- 外接程序具有在 Word 和 PowerPoint 中都有用的功能，例如文本操作，但它具有仅在 PowerPoint 中有意义的一些附加功能，如幻灯片管理功能。 当外接程序在 Word PowerPoint时，您需要隐藏仅支持这些功能的功能。
- 您的外接程序具有一项需要 Office JavaScript API 方法的功能，该方法在某些版本的 Office 应用程序（如订阅 Excel）中受支持，但在其他版本中不受支持，例如一次购买 Excel 2016。 但是，加载项具有其他功能，只需Office支持的其他 JavaScript API Excel 2016。 在此方案中，需要在 Excel 2016 上安装外接程序，但需要不受支持的方法的功能应隐藏给 Excel 2016。

本文可帮助你了解应选择的选项，以确保你的外接程序按预期运行，并遍及可能的最广泛的访问群体。

> [!NOTE]
> 有关当前支持Office外接程序的高级别视图，请参阅 Office 客户端应用程序和 Office [外接程序的可用性](/javascript/api/requirement-sets)页面。

> [!TIP]
> 在使用工具（如 Office 加载项的 [Yeoman](yeoman-generator-overview.md) 生成器或 Visual Studio 中的 Office 加载项模板之一）创建加载项项目时，本文中介绍的许多任务都全部或部分完成。 在这种情况下，请将任务解释为应验证任务已完成的含义。

## <a name="use-the-latest-office-javascript-api-library"></a>使用最新 Office JavaScript API 库

加载项应从内容交付网络Office JavaScript API 库的最新版本 (CDN) 。 为此，请确保外接程序打开的第一个 `script` HTML 文件中具有以下标记。 使用 CDN URL 中的 `/1/` 可以确保引用的是最新版本的 Office.js。

```HTML
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
```

## <a name="specify-which-office-applications-can-host-your-add-in"></a>指定Office哪些应用程序可以托管外接程序

默认情况下，外接程序可安装在指定外接程序类型 (（即"邮件"、"任务窗格"或"内容") ）支持的所有 Office 应用程序中。 例如，默认情况下，任务窗格外接程序可安装在 Access、Excel、OneNote、PowerPoint、Project 和 Word 上。 

若要确保您的外接程序可安装在一组 Office应用程序中，请使用清单中的 [Hosts](/javascript/api/manifest/hosts) 和 [Host](/javascript/api/manifest/host) 元素。

例如，以下 **Hosts** 和 **Host** 声明指定外接程序可以安装在任何 Excel 版本（包括 Excel web 版、Windows 和 iPad）上，但不能安装在任何其他 Office 应用程序上。

```xml
<Hosts>
  <Host Name="Workbook" />
</Hosts>
```

**Hosts** 元素可以包含一个或多个 **Host** 元素。 对于 **应安装外接程序** 的每个 Office 应用程序，都应有一个单独的 Host 元素。 属性 `Name` 是必需的，可以设置为下列值之一。

| 名称          | Office客户端应用程序                     | 可用的外接程序类型 |
|:--------------|:-----------------------------------------------|:-----------------------|
| 数据库      | Access Web App                                | 任务窗格              |
| 文档      | Word web 版、Windows、Mac、iPad            | 任务窗格              |
| 邮箱       | Outlook 网页版、Windows、Mac、Android、iOS | 邮件                   |
| 笔记本      | OneNote 网页版                             | 任务窗格、内容     |
| 演示文稿  | PowerPoint web 版、Windows、Mac、iPad      | 任务窗格、内容     |
| 项目       | Windows 版 Project                             | 任务窗格              |
| 工作簿      | Excel web 版、Windows、Mac、iPad           | 任务窗格、内容     |

> [!NOTE]
> Office应用程序支持在不同的平台上运行，并且这些应用程序在桌面、Web 浏览器、平板电脑和移动设备上运行。 通常无法指定可用于运行外接程序的平台。 例如，如果指定 ，`Workbook`Excel web 版 和 Windows 都可用于运行加载项。 但是，如果指定 `Mailbox`，外接程序将不会在移动Outlook运行，除非您定义[移动扩展点](/javascript/api/manifest/extensionpoint#mobilemessagereadcommandsurface)。

> [!NOTE]
> 外接程序清单不能应用于多个类型：邮件、任务窗格或内容。 这意味着，如果您希望您的外接程序可以在 Outlook 和其他 Office 应用程序之一上安装，则必须创建两个外接程序，一个外接程序具有邮件类型清单，另一个外接程序具有任务窗格或内容类型清单。

> [!IMPORTANT]
> 我们不建议在 SharePoint 中创建和使用 Access Web 应用和数据库。 作为一种替代方法，我们建议你使用 [Microsoft PowerApps](https://powerapps.microsoft.com/) 生成适用于 Web 和移动设备的无代码业务解决方案。

## <a name="specify-which-office-versions-and-platforms-can-host-your-add-in"></a>指定Office哪些版本和平台可以托管外接程序

无法显式指定 Office 版本和内部版本或外接程序应安装在其中的平台，您不希望安装，因为只要对外接程序使用的外接程序功能的支持扩展到新版本或平台，您就必须修改清单。 相反，在清单中指定外接程序所需的 API。 Office阻止在不支持 API 的 Office 版本和平台的组合上安装外接程序，并确保外接程序不会显示在"我的外接程序"**中**。

> [!IMPORTANT]
> 仅使用基本清单指定外接程序必须具有任何重要值的 API 成员。 如果你的外接程序将 API 用于某些功能，但具有其他不需要 API 的有用功能，则应该设计外接程序，以便它可以安装在不支持 API 但提供这些组合体验降低的平台和 Office 版本组合上。 有关详细信息，请参阅 [设计备用体验](#design-for-alternate-experiences)。

### <a name="requirement-sets"></a>要求集

若要简化指定外接程序所需的 API 的过程，Office要求集将大多数 API *组合在一起*。 通用 API 对象模型中 [的 API](understanding-the-javascript-api-for-office.md#api-models) 按它们支持的开发功能进行分组。 例如，连接到表绑定的所有 API 均在名为"TableBindings 1.1"的要求集内。 应用程序特定对象 [模型中](understanding-the-javascript-api-for-office.md#api-models) 的 API 在发布以用于生产外接程序时进行分组。

要求集已进行版本控制。 例如，支持对话框的 [API 在要求](../design/dialog-boxes.md) 集 DialogApi 1.1 中。 当释放支持从任务窗格到对话框的消息的其他 API 时，它们与 DialogApi 1.1 中所有 API 一起分组到 DialogApi 1.2 中。 *要求集的每个版本都是所有早期版本的超集。*

要求集支持Office应用程序、Office版本及其运行平台的不同而不同。 例如，在 Office 2021 之前，Office 的一次购买版本不支持 DialogApi 1.2，但返回到 Office 2013 的所有一次购买版本均支持 DialogApi 1.1。 您希望您的外接程序可安装在支持其使用的 API 的每个平台和 Office 版本组合上，因此应始终在清单中指定外接程序要求的每个要求集的最低版本。 本文稍后将详细介绍如何进行此操作。

> [!TIP]
> 有关要求集版本控制详细信息，请参阅 [Office](office-versions-and-requirement-sets.md#office-requirement-sets-availability) 要求集可用性，有关每个要求集和 API 的完整列表，请从 [Office 外接程序](/javascript/api/requirement-sets/common/office-add-in-requirement-sets)要求集开始。 大多数 API 的参考Office.js还指定它们所属的要求集 (（如果有) ）。

> [!NOTE]
> 某些要求集还具有与其关联的清单元素。 有关 [此事实何时与](#specify-requirements-in-a-versionoverrides-element) 外接程序设计相关的信息，请参阅在 VersionOverrides 元素中指定要求。

#### <a name="apis-not-in-a-requirement-set"></a>不在要求集的 API

应用程序特定模型的所有 API 均在要求集内，但通用 API 模型中的一些 API 不在要求集内。 还有一种方法可以在加载项需要时在清单中指定其中一个无设置 API。 详细信息请参见本文稍后部分。

### <a name="requirements-element"></a>Requirements 元素

使用 [Requirements](/javascript/api/manifest/requirements) 元素及其子元素 [Sets](/javascript/api/manifest/sets) 和 [Methods](/javascript/api/manifest/methods) 指定安装外接程序时 Office 应用程序必须支持的最低要求集或 API 成员。 

如果 Office 应用程序或平台不支持 **Requirements** 元素中指定的要求集或 API 成员，外接程序将不会在该应用程序或平台中运行，也不会显示在"我的外接程序"**中**。

> [!NOTE]
> **Requirements** 元素对于所有外接程序都是可选的，但Outlook外接程序除外。当根`xsi:type`元素的 属性`OfficeApp``MailBox`为 时，必须存在 **一个 Requirements** 元素，该元素指定外接程序所需的 MailBox 要求集的最低版本。 有关详细信息，请参阅 Outlook [JavaScript API 要求集](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets)。

以下代码示例演示如何配置一个加载项，该加载项可安装在Office支持以下内容的所有应用程序：

-  `TableBindings` 要求集，最低版本为"1.1"。
-  `OOXML` 要求集，最低版本为"1.1"。
-  `Document.getSelectedDataAsync` 方法。

```XML
<OfficeApp ... >
  ...
  <Requirements>
     <Sets DefaultMinVersion="1.1">
        <Set Name="TableBindings" MinVersion="1.1"/>
        <Set Name="OOXML" MinVersion="1.1"/>
     </Sets>
     <Methods>
        <Method Name="Document.getSelectedDataAsync"/>
     </Methods>
  </Requirements>
    ...
</OfficeApp>
```
关于此示例，请注意以下事项。

- **Requirements 元素** 包含 **Sets** 和 **Methods** 子元素。
- **Sets 元素** 可以包含一个或多个 **Set** 元素。 `DefaultMinVersion` 指定所有子 `MinVersion` **Set** 元素的默认值。
- [Set](/javascript/api/manifest/set) 元素指定要求集，Office应用程序必须支持该要求集，使外接程序可安装。 属性 `Name` 指定要求集的名称。 指定 `MinVersion` 要求集的最低版本。 `MinVersion` 替代父 `DefaultMinVersion` **Sets 中的 属性的值**。
- **Methods 元素** 可以包含一个或多个 [Method](/javascript/api/manifest/method) 元素。 不能将 **Methods** 元素和 Outlook 外接程序结合使用。
- **Method** 元素指定应用程序必须支持的Office方法，使加载项可安装。 属性 `Name` 是必需的，并指定使用其父对象限定的方法的名称。

## <a name="design-for-alternate-experiences"></a>备用体验设计

加载项平台提供的Office扩展性功能可以分为三类：

- 安装外接程序后立即可用的扩展性功能。 可以通过在清单中配置 [VersionOverrides](/javascript/api/manifest/versionoverrides) 元素来利用此类功能。 外接程序命令是此类功能 [的示例](../design/add-in-commands.md)，这些命令是自定义功能区按钮和菜单。
- 仅在外接程序正在运行并且使用 JavaScript API 实现的扩展性Office.js可用;例如 [，对话框](../design/dialog-boxes.md)。
- 仅在运行时可用的扩展性功能，但通过结合使用 Office.js JavaScript 和 **VersionOverrides** 元素中的配置来实施这些功能。 这些示例包括Excel[、](../excel/custom-functions-overview.md)[单一登录](sso-in-office-add-ins.md)和自定义[上下文选项卡](../design/contextual-tabs.md)。

如果您的外接程序将其特定扩展性功能用于其某些功能，但具有其他不需要扩展性功能的有用功能，您应该设计外接程序，以便它可安装在不支持扩展性功能的平台和 Office 版本组合上。 它可以为这些组合提供有价值的但减少的体验。 

根据扩展性功能实现方式的不同，您实现此设计的方式有所不同： 

- 有关完全使用 JavaScript 实现的功能，请参阅运行时 [检查方法和要求集支持](#runtime-checks-for-method-and-requirement-set-support)。
- 有关需要配置 **VersionOverrides** 元素的功能，请参阅在 [VersionOverrides 元素中指定要求](#specify-requirements-in-a-versionoverrides-element)。

### <a name="runtime-checks-for-method-and-requirement-set-support"></a>方法和要求集支持的运行时检查 

可以在运行时进行测试，以发现用户的Office [isSetSupported](/javascript/api/office/office.requirementsetsupport#office-office-requirementsetsupport-issetsupported-member(1)) 方法是否支持要求集。 将要求集的名称和最低版本作为参数传递。 如果要求集受支持，则返回 `isSetSupported` **true**。 以下代码是一个示例。

```js
if (Office.context.requirements.isSetSupported('WordApi', '1.1'))
{
   // Code that uses API members from the WordApi 1.1 requirement set.
} else {
   // Provide diminished experience here. E.g., run alternate code when the user's Word is one-time purchase Word 2013 (which does not support WordApi 1.1).
}
```
关于此代码，请注意以下几点：

- 第一个参数是必需的。 它是表示要求集名称的字符串。 有关可用要求集的详细信息，请参阅 [Office 加载项要求集](/javascript/api/requirement-sets/common/office-add-in-requirement-sets)。
- 第二个参数是可选的。 它是一个字符串，指定 Office `if` 应用程序必须支持的最低要求集版本，以便语句中的代码运行 (例如"**1.9**") 。 如果未使用，则假定版本"1.1"。

> [!WARNING]
> 调用 方法 `isSetSupported` 时，如果指定了 (，则第二) 参数的值应为字符串而不是数字。 JavaScript 分析程序无法区分数值（如 1.1 和 1.10），而它可以用于字符串值，如"1.1"和"1.10"。

下表显示了特定于应用程序的 API 模型的要求集名称。

|Office 应用程序|RequirementSetName|
|---|---|
|Excel|ExcelApi|
|OneNote|OneNoteApi|
|Outlook|Mailbox|
|PowerPoint|PowerPointApi|
|Word|WordApi|

下面是将 方法与通用 API 模型要求集之一使用的示例。

```js
if (Office.context.requirements.isSetSupported('CustomXmlParts'))
{
    // Run code that uses API members from the CustomXmlParts requirement set.
}
else
{
    // Run alternate code when the user's Word doesn't support the CustomXmlParts requirement set.
}
```

> [!NOTE] 
> 这些`isSetSupported`应用程序的 方法和要求集可在 Office.js 上的最新 CDN。 如果您不使用 `isSetSupported` Office.js 中的CDN，则当您使用的是未定义的旧版本的库时，外接程序可能会生成异常。 有关详细信息，请参阅使用最新 [Office JavaScript API 库](#use-the-latest-office-javascript-api-library)。

当加载项依赖于不是要求集一部分的方法时，请使用运行时检查来确定 Office 应用程序是否支持该方法，如以下代码示例所示。 有关不属于要求集的方法的完整列表，请参阅 [Office 加载项要求集](/javascript/api/requirement-sets/common/office-add-in-requirement-sets#methods-that-arent-part-of-a-requirement-set)。

> [!NOTE]
> 建议限制在加载项代码中使用此类型运行时检查。

下面的代码示例检查应用程序Office是否支持 `document.setSelectedDataAsync`。

```js
if (Office.context.document.setSelectedDataAsync)
{
    // Run code that uses `document.setSelectedDataAsync`.
}
```

### <a name="specify-requirements-in-a-versionoverrides-element"></a>在 VersionOverrides 元素中指定要求

[VersionOverrides](/javascript/api/manifest/versionoverrides) 元素主要（但并非独占）添加到清单架构中，以支持在安装外接程序后立即可用的功能，例如外接程序命令 (自定义功能区按钮和菜单) 。 Office分析外接程序清单时，必须了解这些功能。 

假设您的外接程序使用这些功能之一，但外接程序非常有价值，并且应该可安装，即使在不支持该功能的 Office 版本上。 在此方案中，使用 [Requirements](/javascript/api/manifest/requirements) 元素 (及其子 [Sets](/javascript/api/manifest/sets) 和 [Methods](/javascript/api/manifest/methods) 元素) （作为 **VersionOverrides** `OfficeApp` 元素本身的子元素，而不是作为基元素的子元素包含）标识功能。 这样做的效果是，Office 将允许安装外接程序，但 Office 将忽略不支持该功能的 Office 版本上的 **VersionOverrides** 元素的某些子元素。

具体而言，将忽略替代基本清单中的元素的 **VersionOverrides** 的子元素，如 **Hosts** 元素，并改为使用基本清单的相应元素。 但是， **VersionOverrides** 中的子元素可以实际实现其他功能，而不是替代基本清单中的设置。 两个示例是 和 `WebApplicationInfo` `EquivalentAddins`。 假定版本的平台和版本支持相应的功能，Office **VersionOverrides** 的这些部分不会被忽略。  

有关 **Requirements** 元素的后代元素的信息，请参阅本文前面介绍的 [Requirements](#requirements-element) 元素。

示例如下。

```XML
<VersionOverrides ... >
   ...
   <Requirements>
      <Sets DefaultMinVersion="1.1">
         <Set Name="WordApi" MinVersion="1.2"/>
      </Sets>
   </Requirements>
   <Hosts>

      <!-- ALL MARKUP INSIDE THE HOSTS ELEMENT IS IGNORED WHEREVER WordApi 1.2 IS NOT SUPPORTED -->

      <Host xsi:type="Workbook">
         <!-- markup for custom add-in commands -->
      </Host>
   </Hosts>
</VersionOverrides>
```

> [!WARNING]
> 在 **VersionOverrides** 中使用 **Requirements** 元素之前要谨慎，因为在不支持该要求的平台和版本组合上，不会安装任何外接程序命令，甚至调用不需要该要求的功能的命令。  例如，请考虑具有两个自定义功能区按钮的外接程序。 其中一个Office调用要求集 **ExcelApi 1.4** (及更高版本中可用的 JavaScript) 。 其他调用仅在 **ExcelApi 1.9** (及更高版本) 。 如果在 **VersionOverrides** **中对 ExcelApi 1.9** 提出要求，则当不支持 1.9 时，功能区上将不会显示任何按钮。 此方案中更好的策略是使用运行时检查方法和要求集支持 [中所述的技术](#runtime-checks-for-method-and-requirement-set-support)。 第二个按钮调用的代码首先 `isSetSupported` 用于检查 **是否支持 ExcelApi 1.9**。 如果不支持此功能，则代码会向用户显示一条消息，指出外接程序的此功能在外接程序的 Office。 

> [!TIP]
> 在基本清单中已经出现的 **VersionOverrides** 中，无需重复 Requirement 元素。 如果要求在基本清单中指定，则外接程序无法安装在不支持该要求的位置，因此Office甚至不会分析 **VersionOverrides** 元素。 

## <a name="see-also"></a>另请参阅

- [Office 加载项 XML 清单](add-in-manifests.md)
- [Office 加载项要求集](/javascript/api/requirement-sets/common/office-add-in-requirement-sets)
- [Word-Add-in-Get-Set-EditOpen-XML](https://github.com/OfficeDev/Word-Add-in-Get-Set-EditOpen-XML)
