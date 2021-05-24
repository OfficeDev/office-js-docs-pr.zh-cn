---
title: 清单文件中运行时
description: Runtime 元素将外接程序配置为将共享 JavaScript 运行时用于其各种组件，例如功能区、任务窗格、自定义函数。
ms.date: 05/19/2021
localization_priority: Normal
ms.openlocfilehash: cd09abe31ff57eac629c6c61c873c5c886f73f9c
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/21/2021
ms.locfileid: "52590910"
---
# <a name="runtime-element"></a>运行时元素

将外接程序配置为使用共享的 JavaScript 运行时，以便各种组件都在同一运行时中运行。 元素的 [`<Runtimes>`](runtimes.md) 子元素。

**外接程序类型：** 任务窗格、邮件

[!include[Runtimes support](../../includes/runtimes-note.md)]

## <a name="syntax"></a>语法

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a>包含于

- [运行时](runtimes.md)

## <a name="child-elements"></a>子元素

|  元素 |  必需  |  说明  |
|:-----|:-----|:-----|
| [Override](override.md) | 否 | **Outlook**：指定 Desktop 为 [LaunchEvent](../../reference/manifest/extensionpoint.md#launchevent)扩展点处理程序Outlook JavaScript 文件的 URL 位置。 **重要** 提示：目前只能定义一 `<Override>` 个元素，并且必须为 类型 `javascript` 。|

## <a name="attributes"></a>属性

|  属性  |  必需  |  说明  |
|:-----|:-----|:-----|
|  **resid**  |  是  | 指定外接程序的 HTML 页面的 URL 位置。 `resid`不能超过 32 个字符，并且必须与 元素中的 `id` `Url` 元素的 属性 `Resources` 匹配。 |
|  **lifetime**  |  否  | 的默认值是 `lifetime` `short` ，不需要指定。 Outlook加载项只能使用 `short` 值。 如果要在加载项中Excel运行时，请显式将值设置为 `long` 。 |

## <a name="see-also"></a>另请参阅

- [运行时](runtimes.md)
- [将 Office 加载项配置为使用共享 JavaScript 运行时](../../develop/configure-your-add-in-to-use-a-shared-runtime.md)
- [配置Outlook加载项进行基于事件的激活](../../outlook/autolaunch.md)
