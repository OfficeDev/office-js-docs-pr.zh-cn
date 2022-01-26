---
title: 在功能区上定位自定义选项卡
description: 了解如何控制自定义选项卡在功能区的Office位置，以及默认情况下它是否具有焦点。
ms.date: 01/22/2022
ms.localizationpriority: medium
ms.openlocfilehash: bced5bf5506d0366b29d8e2ad6803b0ddfaad80b
ms.sourcegitcommit: ae3a09d905beb4305a6ffcbc7051ad70745f79f9
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/26/2022
ms.locfileid: "62222092"
---
# <a name="position-a-custom-tab-on-the-ribbon"></a>在功能区上定位自定义选项卡

您可以使用外接程序清单中的标记指定您希望外接程序的自定义选项卡Office应用程序功能区上显示在哪里。

> [!NOTE]
> 本文假定您熟悉文章 [Basic concepts for add-in commands](add-in-commands.md)。 如果你最近未这样做，请查看它。

> [!IMPORTANT]
>
> - 本文中介绍的加载项功能与标记仅在 PowerPoint web 版 *中提供*。
> - 本文中介绍的标记仅适用于支持要求集 **AddinCommands 1.3 的平台**。 请参阅 [下面的不受支持的平台上](#behavior-on-unsupported-platforms) 的行为。

通过标识希望自定义选项卡位于哪个内置 Office 选项卡旁边并指定该选项卡应位于内置选项卡的左侧还是右侧，指定自定义选项卡的显示位置。在外接程序清单的[CustomTab](../reference/manifest/customtab.md)元素中加入[InsertBefore](../reference/manifest/customtab.md#insertbefore) ( (left) 或[InsertAfter](../reference/manifest/customtab.md#insertafter) (right) 元素，从而指定这些规范。  (不能同时具有这两个元素。) 

在下面的示例中，自定义选项卡配置为显示在"审阅 *"***选项卡的正** 之后。请注意 **，InsertAfter** 元素的值是内置"属性"选项卡Office ID。 

```xml
<ExtensionPoint xsi:type="ContosoRibbonTab">
  <CustomTab id="Contoso.TabCustom2">
    <Group id="Contoso.TabCustom2.group2">
       <!-- additional markup omitted -->
    </Group>
    <Label resid="customTabLabel1" />
    <InsertAfter>TabReview</InsertAfter>
  </CustomTab>
</ExtensionPoint>
```

请记住以下几点。

- **InsertBefore** 和 **InsertAfter** 元素是可选的。 如果两者均不使用，则自定义选项卡将显示为功能区最右边的选项卡。
- **InsertBefore** 和 **InsertAfter** 元素相互排斥。 不能同时使用这两者。
- 如果用户安装了多个自定义选项卡配置为同一位置的外接程序（例如，在"审阅"选项卡之后，则最近安装的外接程序的选项卡将位于该位置）。 以前安装的加载项的选项卡将移动到一处。 例如，用户按该顺序安装加载项 A、B 和 C，且所有加载项均配置为在"审阅"选项卡后插入选项卡，选项卡将按以下顺序显示：Review、AddinCTab、AddinBTab、AddinATab。    
- 用户可以在应用程序应用程序中自定义Office功能区。 例如，用户可以移动或隐藏外接程序的选项卡。您无法阻止此情况或检测到它已发生。
- 如果用户移动其中一个内置选项卡，则Office根据内置选项卡的默认位置解释 **InsertBefore** 和 **InsertAfter** *元素*。例如，如果用户将"审阅"选项卡移到功能区的右端，Office 会将上一示例中的标记解释为"将自定义选项卡放在"审阅"选项卡默认位置的右侧。 **

## <a name="specify-which-tab-has-focus-when-the-document-opens"></a>指定文档打开时哪个选项卡具有焦点

Office始终将默认焦点放在紧接"文件"选项卡右边 **的选项卡上**。默认情况下，这是"主页 **"** 选项卡。如果将自定义选项卡配置为在"开始"选项卡之前，则当文档打开时，自定义选项卡将 `<InsertBefore>TabHome</InsertBefore>` 具有焦点。

> [!IMPORTANT]
> 过分强调加载项的不便，并惹恼用户和管理员。 不要将自定义选项卡定位到"主页"选项卡之前，除非外接程序是用户将与文档交互的主要方式。

## <a name="behavior-on-unsupported-platforms"></a>不受支持的平台上的行为

如果加载项安装在不支持要求集 [AddinCommands 1.3](../reference/requirement-sets/add-in-commands-requirement-sets.md)的平台上，则本文中介绍的标记将被忽略，自定义选项卡将显示为功能区最右边的选项卡。 若要防止外接程序安装在不支持标记的平台上，请添加对清单的"要求"部分的要求集的引用。  有关说明，请参阅[指定Office哪些版本和平台可以托管你的外接程序](../develop/specify-office-hosts-and-api-requirements.md#specify-which-office-versions-and-platforms-can-host-your-add-in)。 或者，当 **AddinCommands 1.3** 不受支持时，将外接程序设计为具有备用体验，如设计 [备用体验中所述](../develop/specify-office-hosts-and-api-requirements.md#design-for-alternate-experiences)。 例如，如果您的外接程序包含假定自定义选项卡位于您需要的位置的说明，则您可能具有假定该选项卡最右边的备用版本。
