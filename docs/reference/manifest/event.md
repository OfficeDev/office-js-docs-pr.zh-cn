---
title: 清单文件中的 Event 元素
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 51bbcd5a3d5abe60b850e88e4063e6bbc2da37bc
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450588"
---
# <a name="event-element"></a><span data-ttu-id="77167-102">Event 元素</span><span class="sxs-lookup"><span data-stu-id="77167-102">Event element</span></span>

<span data-ttu-id="77167-103">定义外接程序中的事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="77167-103">Defines an event handler in an add-in.</span></span>

> [!NOTE] 
> <span data-ttu-id="77167-104">目前`Event` , Outlook 在 Office 365 中的网站仅支持该元素。</span><span class="sxs-lookup"><span data-stu-id="77167-104">The `Event` element is currently only supported by Outlook on the web in Office 365.</span></span>

## <a name="attributes"></a><span data-ttu-id="77167-105">属性</span><span class="sxs-lookup"><span data-stu-id="77167-105">Attributes</span></span>

|  <span data-ttu-id="77167-106">属性</span><span class="sxs-lookup"><span data-stu-id="77167-106">Attribute</span></span>  |  <span data-ttu-id="77167-107">必需</span><span class="sxs-lookup"><span data-stu-id="77167-107">Required</span></span>  |  <span data-ttu-id="77167-108">说明</span><span class="sxs-lookup"><span data-stu-id="77167-108">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="77167-109">Type</span><span class="sxs-lookup"><span data-stu-id="77167-109">Type</span></span>](#type-attribute)  |  <span data-ttu-id="77167-110">是</span><span class="sxs-lookup"><span data-stu-id="77167-110">Yes</span></span>  | <span data-ttu-id="77167-111">指定要处理的事件。</span><span class="sxs-lookup"><span data-stu-id="77167-111">Specifies the event to handle.</span></span> |
|  [<span data-ttu-id="77167-112">FunctionExecution</span><span class="sxs-lookup"><span data-stu-id="77167-112">FunctionExecution</span></span>](#functionexecution-attribute)  |  <span data-ttu-id="77167-113">是</span><span class="sxs-lookup"><span data-stu-id="77167-113">Yes</span></span>  | <span data-ttu-id="77167-p101">指定事件处理程序的执行风格、异步或同步。目前仅支持同步事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="77167-p101">Specifies the execution style for the event handler, asynchronous or synchronous. Currently only synchronous event handlers are supported.</span></span> |
|  [<span data-ttu-id="77167-116">FunctionName</span><span class="sxs-lookup"><span data-stu-id="77167-116">FunctionName</span></span>](#functionname-attribute)  |  <span data-ttu-id="77167-117">是</span><span class="sxs-lookup"><span data-stu-id="77167-117">Yes</span></span>  | <span data-ttu-id="77167-118">指定事件处理程序的函数名称。</span><span class="sxs-lookup"><span data-stu-id="77167-118">Specifies the function name for the event handler.</span></span> |

### <a name="type-attribute"></a><span data-ttu-id="77167-119">类型属性</span><span class="sxs-lookup"><span data-stu-id="77167-119">Type attribute</span></span>

<span data-ttu-id="77167-p102">必需。指定哪些事件会调用此事件处理程序。此属性的可能值在下表中指定。</span><span class="sxs-lookup"><span data-stu-id="77167-p102">Required. Specifies which event will invoke the event handler. The possible values for this attribute are specified in the following table.</span></span>

|  <span data-ttu-id="77167-123">事件类型</span><span class="sxs-lookup"><span data-stu-id="77167-123">Event type</span></span>  |  <span data-ttu-id="77167-124">说明</span><span class="sxs-lookup"><span data-stu-id="77167-124">Description</span></span>  |
|:-----|:-----|
|  `ItemSend`  |  <span data-ttu-id="77167-125">在用户发送邮件或会议邀请时将调用此事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="77167-125">The event handler will be invoked when the user sends a message or meeting invitation.</span></span>  |

### <a name="functionexecution-attribute"></a><span data-ttu-id="77167-126">FunctionExecution 属性</span><span class="sxs-lookup"><span data-stu-id="77167-126">FunctionExecution attribute</span></span>

<span data-ttu-id="77167-127">必需。</span><span class="sxs-lookup"><span data-stu-id="77167-127">Required.</span></span> <span data-ttu-id="77167-128">必须设置为 `synchronous`。</span><span class="sxs-lookup"><span data-stu-id="77167-128">MUST be set to `synchronous`.</span></span>

### <a name="functionname-attribute"></a><span data-ttu-id="77167-129">FunctionName 属性</span><span class="sxs-lookup"><span data-stu-id="77167-129">FunctionName attribute</span></span>

<span data-ttu-id="77167-p104">必需。指定事件处理程序的函数名称。该值必须与外接程序的[函数文件](functionfile.md)中的函数名称相匹配。</span><span class="sxs-lookup"><span data-stu-id="77167-p104">Required. Specifies the function name of the event handler. This value must match a function name in the add-in's [function file](functionfile.md).</span></span>

```xml
<Event Type="ItemSend" FunctionExecution="synchronous" FunctionName="itemSendHandler" /> 
```
