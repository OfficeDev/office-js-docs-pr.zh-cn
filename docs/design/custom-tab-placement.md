---
title: 在功能区上定位自定义选项卡
description: 了解如何控制自定义选项卡在 Office 功能区上的显示位置以及默认情况下是否具有焦点。
ms.date: 02/25/2021
localization_priority: Normal
ms.openlocfilehash: 6718a69191d1d84d96512c01b2544094ce276ab6
ms.sourcegitcommit: e7009c565b18c607fe0868db2e26e250ad308dce
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/05/2021
ms.locfileid: "50505204"
---
# <a name="position-a-custom-tab-on-the-ribbon"></a>在功能区上定位自定义选项卡

可以使用加载项清单中的标记指定希望外接程序的自定义选项卡在 Office 应用程序功能区上的显示位置。

> [!NOTE]
> 本文假定您熟悉外接程序 [命令的基本概念一文](add-in-commands.md)。 如果你最近没有这样做，请查看它。

> [!IMPORTANT]
>
> - 本文中介绍的外接程序功能及标记 *仅在 PowerPoint 网页中可用*。
> - 本文中介绍的标记仅适用于支持要求集 **AddinCommands 1.3 的平台**。 请参阅 [下面的不受支持平台上的行为](#behavior-on-unsupported-platforms) 。

通过标识希望自定义选项卡位于哪个内置 Office 选项卡旁边并指定自定义选项卡应位于内置选项卡的左侧还是右侧，来指定自定义选项卡的显示位置。在外接程序清单的[CustomTab](../reference/manifest/customtab.md)元素中添加[InsertBefore](../reference/manifest/customtab.md#insertbefore) () 或[InsertAfter](../reference/manifest/customtab.md#insertafter) (right) 元素，以创建这些规范。  (不能同时具有这两个元素。) 

在下面的示例中，自定义选项卡配置为显示在"审阅"*选项卡***的正** 后。请注意，该元素 `<InsertAfter>` 的值是内置 Office 选项卡的 ID。 

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

- And  `<InsertBefore>`  `<InsertAfter>` 元素是可选的。 如果两者均不使用，则自定义选项卡将显示为功能区最右边的选项卡。
- 和  `<InsertBefore>`  `<InsertAfter>` 元素相互排斥。 不能同时使用这两者。
- 如果用户安装了多个自定义选项卡配置为同一位置的外接程序（例如，在"审阅"选项卡之后，则最近安装的外接程序的选项卡将位于该位置）。 以前安装的加载项的选项卡将移动到一处。 例如，用户按该顺序安装外接程序 A、B 和 C，并且所有加载项均配置为在"审阅"选项卡后插入选项卡，然后选项卡将按以下顺序显示：Review、AddinCTab、AddinBTab、AddinATab。    
- 用户可以在 Office 应用程序中自定义功能区。 例如，用户可以移动或隐藏加载项的选项卡。无法阻止此情况或检测到已发生此情况。
- 如果用户移动其中一个内置选项卡，则 Office 根据内置选项卡的默认位置解释 and `<InsertBefore>` `<InsertAfter>` *元素*。例如，如果用户将"审阅"选项卡移到功能区的右端，Office 会将上述示例中的标记解释为"将自定义选项卡放在"审阅"选项卡默认位置的右侧。" **

## <a name="specifying-which-tab-has-focus-when-the-document-opens"></a>指定文档打开时哪个选项卡具有焦点

Office 始终为紧接在"文件"选项卡右边的选项卡提供 **默认** 焦点。默认情况下，这是"主页 **"** 选项卡。如果将自定义选项卡配置为在"开始"选项卡之前，则打开文档时您的自定义选项卡 `<InsertBefore>TabHome</InsertBefore>` 将具有焦点。

> [!IMPORTANT]
> 过分强调加载项的不便，并惹恼用户和管理员。 除非外接程序是用户与文档交互的主要方式，否则不要将自定义选项卡定位到"主页"选项卡之前。

## <a name="behavior-on-unsupported-platforms"></a>不受支持的平台上的行为

如果加载项安装在不支持要求集 [AddinCommands 1.3](../reference/requirement-sets/add-in-commands-requirement-sets.md)的平台上，则本文中描述的标记将被忽略，并且您的自定义选项卡将显示为功能区最右边的选项卡。 若要防止加载项安装在不支持标记的平台上，请添加对清单部分中的要求集 `<Requirements>` 的引用。 有关说明，请参阅 [清单中的设置 Requirements 元素](../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest)。 或者，可以将外接程序设计成在 **不支持 AddinCommands 1.3** 时具有备用体验，如 [JavaScript](../develop/specify-office-hosts-and-api-requirements.md#use-runtime-checks-in-your-javascript-code)代码中的"使用运行时检查"中所述。 例如，如果您的外接程序包含假定自定义选项卡位于需要它的说明，则您可能有一个备用版本，假定该选项卡最右侧。
