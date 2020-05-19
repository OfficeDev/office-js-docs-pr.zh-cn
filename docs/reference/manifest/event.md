---
title: 清单文件中的 Event 元素
description: 定义外接程序中的事件处理程序。
ms.date: 05/15/2020
localization_priority: Normal
ms.openlocfilehash: 80f21d1819e3d7e335389070ccac0db583026045
ms.sourcegitcommit: 54e2892c0c26b9ad1e4dba8aba48fea39f853b6c
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/18/2020
ms.locfileid: "44275705"
---
# <a name="event-element"></a>Event 元素

定义外接程序中的事件处理程序。

> [!NOTE]
> 有关支持和使用的信息，请参阅[Outlook 外接程序的 "发送时功能"](../../outlook/outlook-on-send-addins.md)。

## <a name="attributes"></a>属性

|  属性  |  必需  |  说明  |
|:-----|:-----|:-----|
|  [Type](#type-attribute)  |  是  | 指定要处理的事件。 |
|  [FunctionExecution](#functionexecution-attribute)  |  是  | 指定事件处理程序的执行风格、异步或同步。目前仅支持同步事件处理程序。 |
|  [FunctionName](#functionname-attribute)  |  是  | 指定事件处理程序的函数名称。 |

### <a name="type-attribute"></a>类型属性

必需。指定哪些事件会调用此事件处理程序。此属性的可能值在下表中指定。

|  事件类型  |  说明  |
|:-----|:-----|
|  `ItemSend`  |  在用户发送邮件或会议邀请时将调用此事件处理程序。  |

### <a name="functionexecution-attribute"></a>FunctionExecution 属性

必需。 必须设置为 `synchronous`。

### <a name="functionname-attribute"></a>FunctionName 属性

必需。指定事件处理程序的函数名称。该值必须与外接程序的[函数文件](functionfile.md)中的函数名称相匹配。

```xml
<Event Type="ItemSend" FunctionExecution="synchronous" FunctionName="itemSendHandler" />
```
