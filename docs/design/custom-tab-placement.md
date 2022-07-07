---
title: 在功能区上定位自定义选项卡
description: 了解如何控制自定义选项卡在 Office 功能区上显示的位置，以及它默认是否具有焦点。
ms.date: 01/22/2022
ms.localizationpriority: medium
ms.openlocfilehash: 42445898623e082c3c85e756625307dc5a237c28
ms.sourcegitcommit: 4ba5f750358c139c93eb2170ff2c97322dfb50df
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/06/2022
ms.locfileid: "66659813"
---
# <a name="position-a-custom-tab-on-the-ribbon"></a>在功能区上定位自定义选项卡

可以通过在加载项清单中使用标记来指定要在 Office 应用程序功能区上显示外接程序的自定义选项卡的位置。

> [!NOTE]
> 本文假定你熟悉 [加载项命令的基本概念](add-in-commands.md)一文。 如果你最近没有这样做，请查看它。

> [!IMPORTANT]
>
> - 本文中所述的加载项功能和标记 *仅在PowerPoint web 版中可用*。
> - 本文中所述的标记仅适用于支持要求集 **AddinCommands 1.3** 的平台。 请参阅下面 [不受支持的平台上的行为](#behavior-on-unsupported-platforms) 。

通过标识希望自定义选项卡位于其旁边的内置 Office 选项卡，并指定它应位于内置选项卡的左侧还是右侧，来指定要显示自定义选项卡的位置。通过在外接程序清单的 [CustomTab](/javascript/api/manifest/customtab) 元素中包括 [InsertBefore](/javascript/api/manifest/customtab#insertbefore) (左) 或 [InsertAfter](/javascript/api/manifest/customtab#insertafter) (右) 元素来制定这些规范。  (不能同时包含这两个元素。) 

在以下示例中，自定义选项卡配置为在 **“审阅**”选项卡 *之后* 显示。请注意，该元素的 **\<InsertAfter\>** 值是内置 Office 选项卡的 ID。 

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

- 这些 **\<InsertBefore\>** 元素和 **\<InsertAfter\>** 元素是可选的。 如果两者都不使用，则自定义选项卡将显示为功能区最右侧的选项卡。
- **\<InsertAfter\>** 和 **\<InsertBefore\>** 元素是相互排斥的。 不能同时使用这两者。
- 如果用户安装多个自定义选项卡配置为同一位置的加载项，请在 **“审阅** ”选项卡后说，那么最近安装的加载项的选项卡将位于该位置。 以前安装的加载项的选项卡将移到一个位置。 例如，用户按该顺序安装加载项 A、B 和 C，并且都配置为在 **“审阅** ”选项卡后插入选项卡，然后按以下顺序显示选项卡： **Review**、 **AddinCTab**、 **AddinBTab**、 **AddinATab**。
- 用户可以在 Office 应用程序中自定义功能区。 例如，用户可以移动或隐藏加载项的选项卡。无法阻止此操作或检测它是否已发生。
- 如果用户移动其中一个内置选项卡，则 Office 会根据 *内置选项卡的默认位置* 解释 **\<InsertBefore\>** 和 **\<InsertAfter\>** 元素。例如，如果用户将 **“审阅”** 选项卡移到功能区右端，Office 会将上一示例中的标记解释为“默认情况下将自定义选项卡放在 ***”审阅**“选项卡* 的右侧。

## <a name="specify-which-tab-has-focus-when-the-document-opens"></a>指定打开文档时哪个选项卡具有焦点

Office 始终为紧靠“ **文件”** 选项卡右侧的选项卡提供默认焦点。默认情况下，这是 **“开始** ”选项卡。如果将自定义选项卡配置为位于 **“开始** ”选项卡之前， `<InsertBefore>TabHome</InsertBefore>`则自定义选项卡在打开文档时将具有焦点。

> [!IMPORTANT]
> 过分强调加载项的不便，并惹恼用户和管理员。 除非外接程序是用户与文档交互的主要方式，否则不要在 **“开始** ”选项卡之前放置自定义选项卡。

## <a name="behavior-on-unsupported-platforms"></a>不受支持的平台上的行为

如果外接程序安装在不支持 [要求集 AddinCommands 1.3](/javascript/api/requirement-sets/common/add-in-commands-requirement-sets) 的平台上，则忽略本文中所述的标记，自定义选项卡将显示为功能区最右侧的选项卡。 若要防止在不支持标记的平台上安装加载项，请在清单部分中 **\<Requirements\>** 添加对要求集的引用。 有关说明，请参阅 [指定哪些 Office 版本和平台可以托管外接程序](../develop/specify-office-hosts-and-api-requirements.md#specify-which-office-versions-and-platforms-can-host-your-add-in)。 或者，在不支持 **AddinCommands 1.3** 时，设计外接程序以获得备用体验，如“设计”中所述 [的备用体验](../develop/specify-office-hosts-and-api-requirements.md#design-for-alternate-experiences)。 例如，如果外接程序包含的说明假定自定义选项卡是你想要的，则可以使用假定选项卡最右侧的备用版本。
