---
title: 清单文件中 Event 元素
description: 定义外接程序中的事件处理程序。
ms.date: 09/24/2021
ms.localizationpriority: medium
ms.openlocfilehash: 095023a8f2d8cd5a01835e09cd50ae7289c98c01
ms.sourcegitcommit: 517786511749c9910ca53e16eb13d0cee6dbfee6
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/29/2021
ms.locfileid: "59990660"
---
# <a name="event-element"></a>Event 元素

定义外接程序中的事件处理程序。

> [!NOTE]
> 有关支持和使用情况的信息，请参阅加载项的Outlook[功能](../../outlook/outlook-on-send-addins.md)。

**外接程序类型：** 邮件

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

必需。必须设置为 `synchronous`。

### <a name="functionname-attribute"></a>FunctionName 属性

必需。指定事件处理程序的函数名称。该值必须与外接程序的[函数文件](functionfile.md)中的函数名称相匹配。

```xml
<Event Type="ItemSend" FunctionExecution="synchronous" FunctionName="itemSendHandler" />
```
