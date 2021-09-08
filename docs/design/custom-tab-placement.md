---
title: 在功能区上定位自定义选项卡
description: 了解如何控制自定义选项卡在功能区的Office位置，以及默认情况下它是否具有焦点。
ms.date: 02/25/2021
localization_priority: Normal
ms.openlocfilehash: 6718a69191d1d84d96512c01b2544094ce276ab6
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/08/2021
ms.locfileid: "58939323"
---
# <a name="position-a-custom-tab-on-the-ribbon"></a>在功能区上定位自定义选项卡

通过使用外接程序清单中的标记，Office加载项的自定义选项卡显示在应用程序功能区上。

> [!NOTE]
> 本文假定您熟悉文章 [Basic concepts for add-in commands](add-in-commands.md)。 如果你最近未这样做，请查看它。

> [!IMPORTANT]
>
> - 本文中介绍的加载项功能与标记仅在 PowerPoint web 版 *中提供*。
> - 本文中介绍的标记仅适用于支持要求集 **AddinCommands 1.3 的平台**。 请参阅 [下面的不受支持的平台上](#behavior-on-unsupported-platforms) 的行为。

通过标识希望自定义选项卡位于哪个内置 Office 选项卡旁边并指定它应位于内置选项卡的左侧还是右侧，来指定自定义选项卡的显示位置。在外接程序清单的[CustomTab](../reference/manifest/customtab.md)元素中添加[InsertBefore](../reference/manifest/customtab.md#insertbefore) () 或[InsertAfter](../reference/manifest/customtab.md#insertafter) (right) 元素，以创建这些规范。  (不能同时具有这两个元素) 

在下面的示例中，自定义选项卡 *配置为显示在"* 审阅"**选项卡的正** 之后。请注意，元素的值是内置"属性"选项卡 `<InsertAfter>` Office ID。 

```xml
<ExtensionPoint xsi:type="ContosoRibbonTab">
  <CustomTab id="TabCustom1">
    <Group id="myCustomTab.group1">
       <!-- additional markup omitted -->
    </Group>
    <Label resid="customTabLabel1" />
    <InsertAfter>TabReview</InsertAfter>
  </CustomTab>
</ExtensionPoint>
```

请记住以下几点。

- 和  `<InsertBefore>`  `<InsertAfter>` 元素是可选的。 如果两者均不使用，则自定义选项卡将显示为功能区最右边的选项卡。
- 和  `<InsertBefore>`  `<InsertAfter>` 元素相互排斥。 不能同时使用这两者。
- 如果用户安装了多个自定义选项卡配置为同一位置的外接程序（例如，在"审阅"选项卡之后，则最近安装的外接程序的选项卡将位于该位置）。 以前安装的加载项的选项卡将移动到一处。 例如，用户按该顺序安装加载项 A、B 和 C，并且所有加载项均配置为在"审阅"选项卡后插入选项卡，选项卡将按以下顺序显示：Review、AddinCTab、AddinBTab、AddinATab。    
- 用户可以在应用程序内自定义Office功能区。 例如，用户可以移动或隐藏外接程序的选项卡。您无法阻止此情况或检测到它已发生。
- 如果用户移动其中一个内置选项卡，则Office根据内置选项卡的默认位置解释 和 `<InsertBefore>` `<InsertAfter>` *元素*。例如，如果用户将"审阅"选项卡移到功能区的右端，Office 会将以上示例中的标记解释为"将自定义选项卡放在"审阅"选项卡默认位置的右侧。 **

## <a name="specifying-which-tab-has-focus-when-the-document-opens"></a>指定文档打开时哪个选项卡具有焦点

Office始终将默认焦点放在紧接"文件"选项卡右边 **的选项卡上**。默认情况下，这是"主页 **"** 选项卡。如果将自定义选项卡配置为在"开始"选项卡之前，则当文档打开时，自定义选项卡将 `<InsertBefore>TabHome</InsertBefore>` 具有焦点。

> [!IMPORTANT]
> 过分强调加载项的不便，并惹恼用户和管理员。 不要将自定义选项卡定位到"主页"选项卡之前，除非外接程序是用户将与文档交互的主要方式。

## <a name="behavior-on-unsupported-platforms"></a>不受支持的平台上的行为

如果外接程序安装在不支持要求集 [AddinCommands 1.3](../reference/requirement-sets/add-in-commands-requirement-sets.md)的平台上，则本文中描述的标记将被忽略，您的自定义选项卡将显示为功能区最右边的选项卡。 若要防止加载项安装在不支持标记的平台上，请添加对清单部分的要求集 `<Requirements>` 的引用。 有关说明，请参阅 [在清单中设置 Requirements 元素](../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest)。 或者，您可以将外接程序设计成在 **AddinCommands 1.3** 不受支持时提供备用体验，如在 [JavaScript](../develop/specify-office-hosts-and-api-requirements.md#use-runtime-checks-in-your-javascript-code)代码中使用运行时检查中所述。 例如，如果您的外接程序包含假定自定义选项卡位于您需要的位置的说明，则您可能具有假定该选项卡最右边的备用版本。
