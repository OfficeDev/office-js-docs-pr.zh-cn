---
title: 指定 Office 主机和 API 要求
description: 了解如何指定外接程序按预期运行的 Office 应用程序和 API 要求。
ms.date: 05/19/2022
ms.localizationpriority: medium
ms.openlocfilehash: 7b1520160e75c0e67eddfae8f8413bc929f35f7f
ms.sourcegitcommit: df7964b6509ee6a807d754fbe895d160bc52c2d3
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/20/2022
ms.locfileid: "66889364"
---
# <a name="specify-office-applications-and-api-requirements"></a>指定 Office 应用程序和 API 要求

Office 加载项可能依赖于特定的 Office 应用程序 (也称为 Office 主机) 或 Office JavaScript API (office.js) 的特定成员。 例如，你的外接程序可能：

- 在单个 Office 应用程序（如，Word 或 Excel）或多个应用程序中运行。
- 使用仅在某些版本的 Office 中可用的 Office JavaScript API。 例如，一次性购买版本的Excel 2016不支持 Office JavaScript 库中与 Excel 相关的所有 API。

在这些情况下，需要确保您的外接程序永远不会安装在无法运行的 Office 应用程序或 Office 版本上。

在某些情况下，你还希望根据用户的 Office 应用程序和 Office 版本控制加载项的哪些功能可见。 两个示例包括：

- 加载项具有在 Word 和 PowerPoint 中都有用的功能，例如文本操作，但它具有一些仅在 PowerPoint 中有意义的附加功能，例如幻灯片管理功能。 加载项在 Word 中运行时，需要隐藏仅限 PowerPoint 的功能。
- 外接程序具有一项功能，该功能需要 Office JavaScript API 方法，该方法在某些版本的 Office 应用程序（如订阅 Excel）中受支持，但在其他版本中不受支持，例如一次性购买Excel 2016。 但加载项具有其他仅需要 Office JavaScript API 方法的功能，这些方法在Excel 2016中 *受* 支持。 在此方案中，需要在Excel 2016上安装加载项，但需要不受支持的方法的功能应对Excel 2016用户隐藏。

本文可帮助你了解应选择的选项，以确保你的外接程序按预期运行，并遍及可能的最广泛的访问群体。

> [!NOTE]
> 有关当前支持 Office 外接程序的高级别视图，请参阅 Office 外接程序的 [Office 客户端应用程序和平台可用性](/javascript/api/requirement-sets) 页面。

> [!TIP]
> 本文中所述的许多任务都是在使用工具（例如 Office 外接程序的 [Yeoman 生成器](yeoman-generator-overview.md) 或 Visual Studio 中的某个 Office 外接程序模板）创建外接程序项目时为你完成的， 在这种情况下，请将该任务解释为应验证任务是否已完成。

## <a name="use-the-latest-office-javascript-api-library"></a>使用最新的 Office JavaScript API 库

外接程序应从内容分发网络加载最新版本的 Office JavaScript API 库 (CDN) 。 为此，请确保加载项打开的第一个 HTML 文件中有以下 `script` 标记。 使用 CDN URL 中的 `/1/` 可以确保引用的是最新版本的 Office.js。

```HTML
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
```

## <a name="specify-which-office-applications-can-host-your-add-in"></a>指定哪些 Office 应用程序可以托管加载项

默认情况下，加载项可安装在指定加载项类型 (（即邮件、任务窗格或内容) ）支持的所有 Office 应用程序中。 例如，默认情况下，可在 Access、Excel、OneNote、PowerPoint、Project 和 Word 上安装任务窗格加载项。

若要确保外接程序可安装在 Office 应用程序的子集中，请使用清单中的 [主机](/javascript/api/manifest/hosts) 和 [主机](/javascript/api/manifest/host) 元素。

例如，以下 **\<Hosts\>** 声明指定 **\<Host\>** 外接程序可以安装在任何版本的 Excel（包括 Excel web 版、Windows 和 iPad）上，但不能安装在任何其他 Office 应用程序上。

```xml
<Hosts>
  <Host Name="Workbook" />
</Hosts>
```

该 **\<Hosts\>** 元素可以包含一个或多个 **\<Host\>** 元素。 每个 Office 应用程序应有一个单独 **\<Host\>** 的元素，加载项应可在其上安装。 该 `Name` 属性是必需的，可以设置为以下值之一。

| 名称          | Office 客户端应用程序                     | 可用加载项类型 |
|:--------------|:-----------------------------------------------|:-----------------------|
| 数据库      | Access Web App                                | 任务窗格              |
| 文档      | Web 上的 Word、Windows、Mac、iPad            | 任务窗格              |
| 邮箱       | Outlook 网页版、Windows、Mac、Android、iOS | 邮件                   |
| 笔记本      | OneNote 网页版                             | 任务窗格，内容     |
| 演示文稿  | PowerPoint web 版、Windows、Mac、iPad      | 任务窗格，内容     |
| 项目       | Windows 版 Project                             | 任务窗格              |
| 工作簿      | Excel web 版、Windows、Mac、iPad           | 任务窗格，内容     |

> [!NOTE]
> Office 应用程序在不同的平台上受支持，并在桌面、Web 浏览器、平板电脑和移动设备上运行。 通常无法指定哪些平台可用于运行外接程序。 例如，如果指定`Workbook`，Excel web 版和 Windows 上都可用于运行加载项。 但是，如果指定 `Mailbox`外接程序，除非定义 [移动扩展点](/javascript/api/manifest/extensionpoint#mobilemessagereadcommandsurface)，否则您的外接程序不会在 Outlook 移动客户端上运行。

> [!NOTE]
> 外接程序清单无法应用于多个类型：邮件、任务窗格或内容。 这意味着，如果希望外接程序可在 Outlook 和其他一个 Office 应用程序上安装，则必须创建 *两* 个加载项，一个加载项包含邮件类型清单，另一个加载项包含任务窗格或内容类型清单。

> [!IMPORTANT]
> 我们不建议在 SharePoint 中创建和使用 Access Web 应用和数据库。 作为一种替代方法，我们建议你使用 [Microsoft PowerApps](https://powerapps.microsoft.com/) 生成适用于 Web 和移动设备的无代码业务解决方案。

## <a name="specify-which-office-versions-and-platforms-can-host-your-add-in"></a>指定哪些 Office 版本和平台可以托管加载项

无法显式指定 Office 版本和内部版本或外接程序应安装的平台，并且不希望这样做，因为每当对外接程序使用的外接程序功能的支持扩展到新版本或平台时，你都必须修改清单。 而是在清单中指定外接程序需要的 API。 Office 阻止在不支持 API 的 Office 版本和平台的组合上安装外接程序，并确保外接程序不会显示在 **“我的外接程序”中**。

> [!IMPORTANT]
> 仅使用基清单来指定加载项必须具有任何重要值的 API 成员。 如果外接程序对某些功能使用 API，但具有其他不需要 API 的有用功能，则应设计外接程序，以便在不支持 API 但提供这些组合体验的平台和 Office 版本组合上安装加载项。 有关详细信息，请参阅 [“设计”以获取备用体验](#design-for-alternate-experiences)。

### <a name="requirement-sets"></a>要求集

为了简化指定外接程序需要的 API 的过程，Office 会在 *要求集中将* 大多数 API 组合在一起。 [公共 API 对象模型中的 API](understanding-the-javascript-api-for-office.md#api-models) 按其支持的开发功能进行分组。 例如，连接到表绑定的所有 API 都位于名为“TableBindings 1.1”的要求集中。 [应用程序特定对象模型](understanding-the-javascript-api-for-office.md#api-models)中的 API 在发布以用于生产外接程序时分组。

要求集已进行版本控制。 例如，支持 [对话框](../develop/dialog-api-in-office-add-ins.md) 的 API 位于要求集 DialogApi 1.1 中。 当释放启用从任务窗格到对话框的消息的其他 API 时，它们与 DialogApi 1.1 中的所有 API 分组到 DialogApi 1.2 中。 *要求集的每个版本都是所有早期版本的超集。*

要求集支持因 Office 应用程序、Office 应用程序的版本及其运行的平台而异。 例如，在Office 2021之前，Office 的一次性购买版本不支持 DialogApi 1.2，但在返回 Office 2013 的所有一次性购买版本上都支持 DialogApi 1.1。 你希望您的外接程序可安装在支持其使用的 API 的平台和 Office 版本的每个组合上，因此应始终在清单中指定加载项所需的每个要求集 *的最低* 版本。 本文稍后将介绍有关如何执行此操作的详细信息。

> [!TIP]
> 有关要求集版本控制的详细信息，请参阅 [Office 要求集可用性](office-versions-and-requirement-sets.md#office-requirement-sets-availability)，以及每个要求集的完整列表以及有关每个 API 的信息，请从 [Office 外接程序要求集](/javascript/api/requirement-sets/common/office-add-in-requirement-sets)开始。 大多数Office.js API 的参考主题还指定它们属于 (（如果有) ）的要求集。

> [!NOTE]
> 某些要求集还具有与其关联的清单元素。 有关此事实何时与外接程序设计相关的信息，请参阅 [VersionOverrides 元素中的指定要求](#specify-requirements-in-a-versionoverrides-element) 。

#### <a name="apis-not-in-a-requirement-set"></a>不在要求集中的 API

应用程序特定模型中的所有 API 都位于要求集中，但通用 API 模型中的某些 API 不在要求集中。 此外，在加载项需要加载项时，还可以在清单中指定这些无设置 API 之一。 详细信息请参见本文稍后部分。

### <a name="requirements-element"></a>Requirements 元素

使用 [Requirements](/javascript/api/manifest/requirements) 元素及其子元素 [集](/javascript/api/manifest/sets) 和 [方法](/javascript/api/manifest/methods) 指定必须由 Office 应用程序支持的最低要求集或 API 成员才能安装外接程序。

如果 Office 应用程序或平台不支持元素中 **\<Requirements\>** 指定的要求集或 API 成员，则外接程序不会在该应用程序或平台中运行，也不会显示在 **“我的外接程序”中**。

> [!NOTE]
> 除 **\<Requirements\>** Outlook 外接程序外，该元素对于所有外接程序都是可选的。`xsi:type`根元素的属性为`MailApp`根`OfficeApp`元素时，必须有一个 **\<Requirements\>** 元素指定加载项所需的邮箱要求集的最小版本。 有关详细信息，请参阅 [Outlook JavaScript API 要求集](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets)。

以下代码示例演示如何配置在支持以下内容的所有 Office 应用程序中可安装的加载项：

- `TableBindings` 要求集，其最低版本为“1.1”。
- `OOXML` 要求集，其最低版本为“1.1”。
- `Document.getSelectedDataAsync` 方法。

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

请注意以下有关此示例的信息。

- 该 **\<Requirements\>** 元素包含元素 **\<Sets\>** 和 **\<Methods\>** 子元素。
- 该 **\<Sets\>** 元素可以包含一个或多个 **\<Set\>** 元素。 `DefaultMinVersion`指定所有子 **\<Set\>** 元素的默认`MinVersion`值。
- [Set](/javascript/api/manifest/set) 元素指定 Office 应用程序必须支持的要求集才能使外接程序可安装。 该 `Name` 属性指定要求集的名称。 指定 `MinVersion` 要求集的最小版本。 `MinVersion`重写父 **\<Sets\>** 属性的`DefaultMinVersion`值。
- 该 **\<Methods\>** 元素可以包含一个或多个 [方法](/javascript/api/manifest/method) 元素。 不能将该 **\<Methods\>** 元素与 Outlook 加载项配合使用。
- 该 **\<Method\>** 元素指定 Office 应用程序必须支持的单个方法才能使加载项可安装。 该 `Name` 属性是必需的，并指定使用其父对象限定的方法的名称。

## <a name="design-for-alternate-experiences"></a>为备用体验设计

Office 外接程序平台提供的扩展性功能可以分为三种类型：

- 安装加载项后立即可用的扩展性功能。 可以通过在清单中配置 [VersionOverrides](/javascript/api/manifest/versionoverrides) 元素来使用此类功能。 此类功能的一个示例是 [外接程序命令](../design/add-in-commands.md)，它是自定义功能区按钮和菜单。
- 扩展性功能，仅当加载项正在运行且使用 Office.js JavaScript API 实现时才可用;例如， [对话框](../develop/dialog-api-in-office-add-ins.md)。
- 可扩展性功能仅在运行时可用，但通过元素中的 **\<VersionOverrides\>** Office.js JavaScript 和配置的组合实现。 这些示例包括 [Excel 自定义函数](../excel/custom-functions-overview.md)、 [单一登录](sso-in-office-add-ins.md)和 [自定义上下文选项卡](../design/contextual-tabs.md)。

如果外接程序对其某些功能使用特定的扩展性功能，但具有其他不需要扩展性功能的有用功能，则应设计外接程序，使其可安装在不支持扩展功能的平台和 Office 版本组合上。 它可以提供一个有价值的，虽然减少，在这些组合的经验。

根据扩展性功能的实现方式，以不同的方式实现此设计：

- 有关完全使用 JavaScript 实现的功能，请参阅 [运行时检查方法和要求集支持](#runtime-checks-for-method-and-requirement-set-support)。
- 有关需要配置元素的 **\<VersionOverrides\>** 功能，请参阅 [VersionOverrides 元素中的指定要求](#specify-requirements-in-a-versionoverrides-element)。

### <a name="runtime-checks-for-method-and-requirement-set-support"></a>运行时检查方法和要求集支持

在运行时进行测试，以发现用户的 Office 是否支持使用 [isSetSupported](/javascript/api/office/office.requirementsetsupport#office-office-requirementsetsupport-issetsupported-member(1)) 方法的要求集。 将要求集的名称和最低版本作为参数传递。 如果支持要求集，则 `isSetSupported` 返回 `true`。 以下代码是一个示例。

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
- 第二个参数是可选的。 它是一个字符串，指定 Office 应用程序必须支持的最低要求集版本，以便语句中的 `if` 代码运行 (例如“**1.9**”) 。 如果未使用，则假定版本“1.1”。

> [!WARNING]
> 调用 `isSetSupported` 该方法时，如果指定的) 应为字符串而不是数字，则 (第二个参数的值。 JavaScript 分析器无法区分数字值（如 1.1 和 1.10），而对于字符串值（如“1.1”和“1.10”）则可以区分。

下表显示了特定于应用程序的 API 模型的要求集名称。

|Office 应用程序|RequirementSetName|
|---|---|
|Excel|ExcelApi|
|OneNote|OneNoteApi|
|Outlook|Mailbox|
|PowerPoint|PowerPointApi|
|Word|WordApi|

下面是将该方法与公共 API 模型要求集之一配合使用的示例。

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
> CDN `isSetSupported` 上的最新Office.js文件中提供了这些应用程序的方法和要求集。 如果不使用 CDN 中的Office.js，则如果使用未定义的库 `isSetSupported` 的旧版本，外接程序可能会生成异常。 有关详细信息，请参阅 [使用最新的 Office JavaScript API 库](#use-the-latest-office-javascript-api-library)。

当外接程序依赖于不属于要求集的方法时，请使用运行时检查来确定该方法是否受 Office 应用程序的支持，如以下代码示例所示。 有关不属于要求集的方法的完整列表，请参阅 [Office 加载项要求集](/javascript/api/requirement-sets/common/office-add-in-requirement-sets#methods-that-arent-part-of-a-requirement-set)。

> [!NOTE]
> 建议限制在加载项代码中使用此类型运行时检查。

以下代码示例检查 Office 应用程序是否支持 `document.setSelectedDataAsync`。

```js
if (Office.context.document.setSelectedDataAsync)
{
    // Run code that uses `document.setSelectedDataAsync`.
}
```

### <a name="specify-requirements-in-a-versionoverrides-element"></a>在 VersionOverrides 元素中指定要求

[VersionOverrides](/javascript/api/manifest/versionoverrides) 元素主要添加到清单架构中，但不只支持安装加载项后必须立即提供的功能，例如外接程序命令 (自定义功能区按钮和菜单) 。 Office 在分析外接程序清单时必须了解这些功能。

假设您的外接程序使用这些功能之一，但外接程序很有价值，应该可安装，即使在不支持该功能的 Office 版本上也可安装。 在此方案中，使用 [Requirements](/javascript/api/manifest/requirements) 元素 (及其子 [集](/javascript/api/manifest/sets) 和 [方法](/javascript/api/manifest/methods) 元素来标识该功能，) 作为元素本身的 **\<VersionOverrides\>** 子元素而非作为基 `OfficeApp` 元素的子元素包含的元素。 这样做的效果是 Office 将允许安装外接程序，但 Office 将忽略 Office 版本中不支持该功能的元素的 **\<VersionOverrides\>** 某些子元素。

具体而言，将忽略替代基清单中的元素（如 **\<Hosts\>** 元素）的子元素 **\<VersionOverrides\>**，并改用基清单的相应元素。 但是，在实际实现其他功能，而不是覆盖基清单中的 **\<VersionOverrides\>** 设置时，可以有子元素。 两个示例是和 `WebApplicationInfo` `EquivalentAddins`。 假设 Office 的 **\<VersionOverrides\>** 平台和版本支持相应的功能，则 *不会* 忽略这些部分。  

有关元素的 **\<Requirements\>** 后代元素的信息，请参阅本文前面的 [Requirements 元素](#requirements-element) 。

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
> 在使用 **\<Requirements\>** 某个元素 **\<VersionOverrides\>** 之前请谨慎使用，因为在不支持要求的平台和版本组合上， *不会安装任何* 外接程序命令， *即使是那些调用不需要要求的功能的命令*。 例如，请考虑具有两个自定义功能区按钮的加载项。 其中一个调用在要求集 **ExcelApi 1.4** (及更高版本) 中可用的 Office JavaScript API。 其他调用仅在 **ExcelApi 1.9** (及更高版本) 中可用的 API。 如果在其中对 **ExcelApi 1.9** **\<VersionOverrides\>** 提出了要求，则在不支持 1.9 时，功能区上不会显示 *任何* 按钮。 在此方案中，更好的策略是使用运行时检查中所述 [的方法和要求集支持](#runtime-checks-for-method-and-requirement-set-support)。 第二个按钮调用的代码首先用于 `isSetSupported` 检查 **ExcelApi 1.9** 的支持。 如果不支持，代码会向用户发送一条消息，指出加载项的此功能在其 Office 版本上不可用。

> [!TIP]
> 在基清单中 **\<VersionOverrides\>** 已显示的 **Requirement** 元素是没意义的。 如果要求在基清单中指定，则加载项无法安装不支持要求的位置，因此 Office 甚至不分析该 **\<VersionOverrides\>** 元素。

## <a name="see-also"></a>另请参阅

- [Office 加载项 XML 清单](add-in-manifests.md)
- [Office 加载项要求集](/javascript/api/requirement-sets/common/office-add-in-requirement-sets)
- [Word-Add-in-Get-Set-EditOpen-XML](https://github.com/OfficeDev/Word-Add-in-Get-Set-EditOpen-XML)
