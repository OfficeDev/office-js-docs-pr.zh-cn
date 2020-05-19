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
# <a name="event-element"></a><span data-ttu-id="77912-103">Event 元素</span><span class="sxs-lookup"><span data-stu-id="77912-103">Event element</span></span>

<span data-ttu-id="77912-104">定义外接程序中的事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="77912-104">Defines an event handler in an add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="77912-105">有关支持和使用的信息，请参阅[Outlook 外接程序的 "发送时功能"](../../outlook/outlook-on-send-addins.md)。</span><span class="sxs-lookup"><span data-stu-id="77912-105">For information about support and usage, see [On-send feature for Outlook add-ins](../../outlook/outlook-on-send-addins.md).</span></span>

## <a name="attributes"></a><span data-ttu-id="77912-106">属性</span><span class="sxs-lookup"><span data-stu-id="77912-106">Attributes</span></span>

|  <span data-ttu-id="77912-107">属性</span><span class="sxs-lookup"><span data-stu-id="77912-107">Attribute</span></span>  |  <span data-ttu-id="77912-108">必需</span><span class="sxs-lookup"><span data-stu-id="77912-108">Required</span></span>  |  <span data-ttu-id="77912-109">说明</span><span class="sxs-lookup"><span data-stu-id="77912-109">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="77912-110">Type</span><span class="sxs-lookup"><span data-stu-id="77912-110">Type</span></span>](#type-attribute)  |  <span data-ttu-id="77912-111">是</span><span class="sxs-lookup"><span data-stu-id="77912-111">Yes</span></span>  | <span data-ttu-id="77912-112">指定要处理的事件。</span><span class="sxs-lookup"><span data-stu-id="77912-112">Specifies the event to handle.</span></span> |
|  [<span data-ttu-id="77912-113">FunctionExecution</span><span class="sxs-lookup"><span data-stu-id="77912-113">FunctionExecution</span></span>](#functionexecution-attribute)  |  <span data-ttu-id="77912-114">是</span><span class="sxs-lookup"><span data-stu-id="77912-114">Yes</span></span>  | <span data-ttu-id="77912-p101">指定事件处理程序的执行风格、异步或同步。目前仅支持同步事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="77912-p101">Specifies the execution style for the event handler, asynchronous or synchronous. Currently only synchronous event handlers are supported.</span></span> |
|  [<span data-ttu-id="77912-117">FunctionName</span><span class="sxs-lookup"><span data-stu-id="77912-117">FunctionName</span></span>](#functionname-attribute)  |  <span data-ttu-id="77912-118">是</span><span class="sxs-lookup"><span data-stu-id="77912-118">Yes</span></span>  | <span data-ttu-id="77912-119">指定事件处理程序的函数名称。</span><span class="sxs-lookup"><span data-stu-id="77912-119">Specifies the function name for the event handler.</span></span> |

### <a name="type-attribute"></a><span data-ttu-id="77912-120">类型属性</span><span class="sxs-lookup"><span data-stu-id="77912-120">Type attribute</span></span>

<span data-ttu-id="77912-p102">必需。指定哪些事件会调用此事件处理程序。此属性的可能值在下表中指定。</span><span class="sxs-lookup"><span data-stu-id="77912-p102">Required. Specifies which event will invoke the event handler. The possible values for this attribute are specified in the following table.</span></span>

|  <span data-ttu-id="77912-124">事件类型</span><span class="sxs-lookup"><span data-stu-id="77912-124">Event type</span></span>  |  <span data-ttu-id="77912-125">说明</span><span class="sxs-lookup"><span data-stu-id="77912-125">Description</span></span>  |
|:-----|:-----|
|  `ItemSend`  |  <span data-ttu-id="77912-126">在用户发送邮件或会议邀请时将调用此事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="77912-126">The event handler will be invoked when the user sends a message or meeting invitation.</span></span>  |

### <a name="functionexecution-attribute"></a><span data-ttu-id="77912-127">FunctionExecution 属性</span><span class="sxs-lookup"><span data-stu-id="77912-127">FunctionExecution attribute</span></span>

<span data-ttu-id="77912-128">必需。</span><span class="sxs-lookup"><span data-stu-id="77912-128">Required.</span></span> <span data-ttu-id="77912-129">必须设置为 `synchronous`。</span><span class="sxs-lookup"><span data-stu-id="77912-129">MUST be set to `synchronous`.</span></span>

### <a name="functionname-attribute"></a><span data-ttu-id="77912-130">FunctionName 属性</span><span class="sxs-lookup"><span data-stu-id="77912-130">FunctionName attribute</span></span>

<span data-ttu-id="77912-p104">必需。指定事件处理程序的函数名称。该值必须与外接程序的[函数文件](functionfile.md)中的函数名称相匹配。</span><span class="sxs-lookup"><span data-stu-id="77912-p104">Required. Specifies the function name of the event handler. This value must match a function name in the add-in's [function file](functionfile.md).</span></span>

```xml
<Event Type="ItemSend" FunctionExecution="synchronous" FunctionName="itemSendHandler" />
```
