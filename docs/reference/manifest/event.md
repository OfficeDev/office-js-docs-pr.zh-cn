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
# <a name="event-element"></a><span data-ttu-id="6c237-102">Event 元素</span><span class="sxs-lookup"><span data-stu-id="6c237-102">Event element</span></span>

<span data-ttu-id="6c237-103">定义外接程序中的事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="6c237-103">Defines an event handler in an add-in.</span></span>

> [!NOTE] 
> <span data-ttu-id="6c237-104">目前仅 Office 365 中的 Outlook 网页版支持 `Event` 元素。</span><span class="sxs-lookup"><span data-stu-id="6c237-104">Note: The `Event` element is currently only supported by Outlook on the web in Office 365.</span></span>

## <a name="attributes"></a><span data-ttu-id="6c237-105">属性</span><span class="sxs-lookup"><span data-stu-id="6c237-105">Attributes</span></span>

|  <span data-ttu-id="6c237-106">属性</span><span class="sxs-lookup"><span data-stu-id="6c237-106">Attribute</span></span>  |  <span data-ttu-id="6c237-107">必需</span><span class="sxs-lookup"><span data-stu-id="6c237-107">Required</span></span>  |  <span data-ttu-id="6c237-108">说明</span><span class="sxs-lookup"><span data-stu-id="6c237-108">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="6c237-109">Type</span><span class="sxs-lookup"><span data-stu-id="6c237-109">Type</span></span>](#type-attribute)  |  <span data-ttu-id="6c237-110">是</span><span class="sxs-lookup"><span data-stu-id="6c237-110">Yes</span></span>  | <span data-ttu-id="6c237-111">指定要处理的事件。</span><span class="sxs-lookup"><span data-stu-id="6c237-111">Specifies the event to handle.</span></span> |
|  [<span data-ttu-id="6c237-112">FunctionExecution</span><span class="sxs-lookup"><span data-stu-id="6c237-112">FunctionExecution</span></span>](#functionexecution-attribute)  |  <span data-ttu-id="6c237-113">是</span><span class="sxs-lookup"><span data-stu-id="6c237-113">Yes</span></span>  | <span data-ttu-id="6c237-p101">指定事件处理程序的执行风格、异步或同步。目前仅支持同步事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="6c237-p101">Specifies the execution style for the event handler, asynchronous or synchronous. Currently only synchronous event handlers are supported.</span></span> |
|  [<span data-ttu-id="6c237-116">FunctionName</span><span class="sxs-lookup"><span data-stu-id="6c237-116">FunctionName</span></span>](#functionname-attribute)  |  <span data-ttu-id="6c237-117">是</span><span class="sxs-lookup"><span data-stu-id="6c237-117">Yes</span></span>  | <span data-ttu-id="6c237-118">指定事件处理程序的函数名称。</span><span class="sxs-lookup"><span data-stu-id="6c237-118">Specifies the function name for the event handler.</span></span> |

### <a name="type-attribute"></a><span data-ttu-id="6c237-119">类型属性</span><span class="sxs-lookup"><span data-stu-id="6c237-119">Type attribute</span></span>

<span data-ttu-id="6c237-p102">必需。指定哪些事件会调用此事件处理程序。此属性的可能值在下表中指定。</span><span class="sxs-lookup"><span data-stu-id="6c237-p102">Required. Specifies which event will invoke the event handler. The possible values for this attribute are specified in the following table.</span></span>

|  <span data-ttu-id="6c237-123">事件类型</span><span class="sxs-lookup"><span data-stu-id="6c237-123">Event type</span></span>  |  <span data-ttu-id="6c237-124">说明</span><span class="sxs-lookup"><span data-stu-id="6c237-124">Description</span></span>  |
|:-----|:-----|
|  `ItemSend`  |  <span data-ttu-id="6c237-125">在用户发送邮件或会议邀请时将调用此事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="6c237-125">The event handler will be invoked when the user sends a message or meeting invitation.</span></span>  |

### <a name="functionexecution-attribute"></a><span data-ttu-id="6c237-126">FunctionExecution 属性</span><span class="sxs-lookup"><span data-stu-id="6c237-126">FunctionExecution attribute</span></span>

<span data-ttu-id="6c237-p103">必需。必须设置为 `synchronous`。</span><span class="sxs-lookup"><span data-stu-id="6c237-p103">Required. MUST be set to `synchronous`.</span></span>

### <a name="functionname-attribute"></a><span data-ttu-id="6c237-129">FunctionName 属性</span><span class="sxs-lookup"><span data-stu-id="6c237-129">FunctionName attribute</span></span>

<span data-ttu-id="6c237-p104">必需。指定事件处理程序的函数名称。该值必须与外接程序的[函数文件](functionfile.md)中的函数名称相匹配。</span><span class="sxs-lookup"><span data-stu-id="6c237-p104">Required. Specifies the function name of the event handler. This value must match a function name in the add-in's [function file](functionfile.md).</span></span>

```xml
<Event Type="ItemSend" FunctionExecution="synchronous" FunctionName="itemSendHandler" /> 
```