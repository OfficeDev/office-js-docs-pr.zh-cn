---
title: 清单文件中运行时
description: Runtime 元素将外接程序配置为将共享 JavaScript 运行时用于其各种组件，例如功能区、任务窗格、自定义函数。
ms.date: 03/22/2022
ms.localizationpriority: medium
ms.openlocfilehash: 38920dc43349be8da629785167d03252578f2a42
ms.sourcegitcommit: 64942cdd79d7976a0291c75463d01cb33a8327d8
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/25/2022
ms.locfileid: "64404672"
---
# <a name="runtime-element"></a>运行时元素

将外接程序配置为使用共享的 JavaScript 运行时，以便各种组件都在同一运行时中运行。 元素的 [`<Runtimes>`](runtimes.md) 子元素。

**外接程序类型：** 任务窗格、邮件

**仅在以下 VersionOverrides 架构中有效**：

 - 任务窗格 1.0
 - 邮件 1.1

有关详细信息，请参阅清单 [中的版本替代](../../develop/add-in-manifests.md#version-overrides-in-the-manifest)。

**与以下要求集相关联**：

- [SharedRuntime 1.1](../requirement-sets/shared-runtime-requirement-sets.md) (仅在任务窗格外接程序中使用时) 

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
| [Override](override.md) | 否 | **Outlook**：指定 Desktop 为 [LaunchEvent](../../reference/manifest/extensionpoint.md#launchevent) 扩展点处理程序Outlook JavaScript 文件的 URL 位置。 **重要** 提示：目前，只能定义一个元素 `<Override>` ，并且必须为 类型 `javascript`。|

## <a name="attributes"></a>属性

|  属性  |  必需  |  说明  |
|:-----|:-----|:-----|
|  **resid**  |  是  | 指定外接程序的 HTML 页面的 URL 位置。 不能 `resid` 超过 32 个字符，并且必须与 `id` `Url` 元素中的 元素的属性 `Resources` 匹配。 |
|  [lifetime](#lifetime-attribute)  |  否  | 的默认值 `lifetime` 是 `short` ，不需要指定。 Outlook事件激活加载项只能使用 `short` 值。 如果要在加载项中Excel运行时，请显式将值设置为 `long`。 |

### <a name="lifetime-attribute"></a>lifetime 属性

可选。 表示允许外接程序运行的时间长度。

**可用值**

`short`：默认值。 仅用于Outlook事件的激活加载项。外接程序激活后，它将按平台指定的最长时间运行。 目前，大约 5 分钟。 这是唯一受 Outlook。

`long`：仅在配置共享 [JavaScript 运行时时使用](../../develop/configure-your-add-in-to-use-a-shared-runtime.md)。 加载项可以在打开文档时启动并无限期运行。 例如，即使用户关闭任务窗格，任务窗格代码仍将继续运行。 这是共享运行时支持的唯一值。

## <a name="see-also"></a>另请参阅

- [运行时](runtimes.md)
- [将 Office 加载项配置为使用共享 JavaScript 运行时](../../develop/configure-your-add-in-to-use-a-shared-runtime.md)
- [配置Outlook加载项进行基于事件的激活](../../outlook/autolaunch.md)
