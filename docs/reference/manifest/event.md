---
title: 清单文件中的 Event 元素
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: eda895b01e106d67eef70f199be64086e9372bef
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432737"
---
# <a name="event-element"></a>Event 元素

定义外接程序中的事件处理程序。

> [!NOTE] 
> 目前仅 Office 365 中的 Outlook 网页版支持 `Event` 元素。

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