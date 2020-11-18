---
title: 在功能区上定位自定义选项卡
description: 了解如何在默认情况下控制自定义选项卡在 Office 功能区上显示的位置以及它是否有焦点。
ms.date: 11/01/2020
localization_priority: Normal
ms.openlocfilehash: 2c1e2ae66805212e78868cf7c07a0e5c14cd4025
ms.sourcegitcommit: 3189c4bd62dbe5950b19f28ac2c1314b6d304dca
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/17/2020
ms.locfileid: "49088165"
---
# <a name="position-a-custom-tab-on-the-ribbon-preview"></a>将自定义选项卡放置在功能区上 (预览) 

您可以使用外接程序清单中的标记来指定您希望外接程序的自定义选项卡在 Office 应用程序的功能区上显示的位置。

> [!NOTE]
> 本文假定您熟悉文章 [外接程序命令的基本概念](add-in-commands.md)。 如果你最近未执行此操作，请查看它。

> [!IMPORTANT]
>
> - 本文中介绍的加载项功能和标记位于预览中， *仅适用于 web 上的 PowerPoint*。 我们建议您仅在测试和开发环境中尝试标记。 请勿在生产环境中或在业务关键型文档中使用预览标记。
> - 本文中所述的标记仅适用于支持要求集 **addincommand 1.3** 的平台。 请参阅下面 [有关不受支持的平台的行为](#behavior-on-unsupported-platforms) 。

指定要显示自定义选项卡的位置，具体方法是确定您希望它在其旁边的内置 "Office" 选项卡，并指定它是在内置选项卡的左侧还是右侧。通过在外接程序清单的[CustomTab](../reference/manifest/customtab.md)元素中包括一个[InsertBefore](../reference/manifest/customtab.md#insertbefore) (left) 或[InsertAfter](../reference/manifest/customtab.md#insertafter) (right) 元素来设置这些规范。  (不能同时具有这两个元素 ) 

在以下示例中，将自定义选项卡配置为 *恰好* 显示在 " **审阅** " 选项卡的后面。请注意，该元素的值 `<InsertAfter>` 是内置 "Office" 选项卡的 ID。 

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

- `<InsertBefore>`和 `<InsertAfter>` 元素是可选的。 如果不使用这两种方式，则自定义选项卡将显示为功能区上最右边的选项卡。
- `<InsertBefore>`和 `<InsertAfter>` 元素相互排斥。 您不能同时使用这两种。
- 如果用户安装了多个加载项，其自定义选项卡配置为相同位置，则在 " **审阅** " 选项卡之后，最近安装的加载项的选项卡将位于该位置。 之前安装的外接程序的选项卡将移动到一个位置。 例如，用户在该顺序中安装外接程序 A、B 和 C，并将所有配置为在 " **审阅** " 选项卡上将其配置为插入一个选项卡，然后选项卡将按如下顺序显示： " **审阅**"、" **AddinCTab**"、" **AddinBTab**"、" **AddinATab**"。
- 用户可以在 Office 应用程序中自定义功能区。 例如，用户可以移动或隐藏外接程序的选项卡。您不能阻止此情况，也不能检测到此问题。
- 如果用户移动了其中一个内置选项卡，则 Office 将 `<InsertBefore>` `<InsertAfter>` 根据 *内置选项卡的默认位置* 来解释和元素。例如，如果用户将 "**审阅**" 选项卡移到功能区的右端，则 Office 会将上面示例中的标记解释为 "将自定义选项卡放在 ***审阅** 选项卡的默认位置* 的右侧"。

## <a name="specifying-which-tab-has-focus-when-the-document-opens"></a>指定在文档打开时哪个选项卡具有焦点

Office 始终向紧靠 " **文件** " 选项卡右侧的选项卡提供默认焦点。默认情况下，这是 " **主页** " 选项卡。如果将自定义选项卡配置为在 " **主页** " 选项卡之前使用 `<InsertBefore>TabHome</InsertBefore>` ，则在打开文档时，自定义选项卡将获得焦点。

> [!IMPORTANT]
> 为您的外接程序 inconveniences 和 annoys 用户和管理员提供了更多的突出。 不要将自定义选项卡放置在 " **主页** " 选项卡之前，除非您的外接程序是用户将与文档进行交互的主要方式。

## <a name="behavior-on-unsupported-platforms"></a>不受支持的平台上的行为

如果您的外接程序安装在不支持 [要求集 addincommand 1.3](../reference/requirement-sets/add-in-commands-requirement-sets.md)的平台上，则会忽略本文中所述的标记，并且您的自定义选项卡将显示为功能区上最右边的选项卡。 若要防止外接程序安装在不支持标记的平台上，请在清单的部分中添加对要求集的引用 `<Requirements>` 。 有关说明，请参阅 [在清单中设置需求元素](../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest)。 此外，还可以设计外接程序，使其在 **addincommand 1.3** 不受支持时具有备用体验，如 [JavaScript 代码中的使用运行时检查](../develop/specify-office-hosts-and-api-requirements.md#use-runtime-checks-in-your-javascript-code)中所述。 例如，如果您的外接程序包含假设自定义选项卡位于您所需的位置的说明，则可以使用该选项卡的备选版本，该选项卡位于最右边。
