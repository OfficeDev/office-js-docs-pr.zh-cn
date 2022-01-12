---
title: 清单文件中 Event 元素
description: 定义外接程序中的事件处理程序。
ms.date: 01/11/2022
ms.localizationpriority: medium
ms.openlocfilehash: fac920fc91abd908d3d159877c0c414bd7fae244
ms.sourcegitcommit: 33824aa3995a2e0bcc6d8e67ada46f296c224642
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/12/2022
ms.locfileid: "61765890"
---
# <a name="event-element"></a>Event 元素

定义外接程序中的事件处理程序。

> [!NOTE]
> 有关支持和使用情况的信息，请参阅[On-send feature for Outlook add-ins](../../outlook/outlook-on-send-addins.md)。

**外接程序类型：** 邮件

**仅在以下 VersionOverrides 架构中有效**：

- 邮件 1.0
- 邮件 1.1

有关详细信息，请参阅清单 [中的版本替代](../../develop/add-in-manifests.md#version-overrides-in-the-manifest)。

## <a name="attributes"></a>属性

|  属性  |  必需  |  说明  |
|:-----|:-----|:-----|
|  [类型](#type-attribute)  |  是  | 指定要处理的事件。 |
|  [FunctionExecution](#functionexecution-attribute)  |  是  | 指定事件处理程序的执行风格、异步或同步。目前仅支持同步事件处理程序。 |
|  [FunctionName](#functionname-attribute)  |  是  | 指定事件处理程序的函数名称。 |

### <a name="type-attribute"></a>类型属性

必需。指定哪些事件会调用此事件处理程序。此属性的可能值在下表中指定。

|  事件类型  |  说明  |
|:-----|:-----|
|  `ItemSend`  |  在用户发送邮件或会议邀请时将调用此事件处理程序。  |

### <a name="functionexecution-attribute"></a>FunctionExecution 属性

必需。必须设置为 `synchronous`。

### <a name="functionname-attribute"></a>FunctionName 属性

必需。指定事件处理程序的函数名称。该值必须与外接程序的[函数文件](functionfile.md)中的函数名称相匹配。

```xml
<Event Type="ItemSend" FunctionExecution="synchronous" FunctionName="itemSendHandler" />
```
