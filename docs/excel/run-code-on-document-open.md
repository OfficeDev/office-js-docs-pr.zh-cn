---
title: 在文档打开时，在 Excel 外接程序中运行代码
description: 在文档打开时，在 Excel 外接程序中运行代码。
ms.date: 05/15/2020
localization_priority: Normal
ms.openlocfilehash: c0e8493b5437a1a8b59d689bbbdd794fa98d89a2
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/08/2020
ms.locfileid: "44610340"
---
# <a name="run-code-in-your-excel-add-in-when-the-document-opens"></a><span data-ttu-id="fb144-103">在文档打开时，在 Excel 外接程序中运行代码</span><span class="sxs-lookup"><span data-stu-id="fb144-103">Run code in your Excel add-in when the document opens</span></span>

<span data-ttu-id="fb144-104">您可以将 Excel 加载项配置为在文档打开后立即加载和运行代码。</span><span class="sxs-lookup"><span data-stu-id="fb144-104">You can configure your Excel add-in to load and run code as soon as the document is opened.</span></span> <span data-ttu-id="fb144-105">如果需要注册事件处理程序、任务窗格的预加载数据、同步 UI 或在外接程序可见之前执行其他任务，这将非常有用。</span><span class="sxs-lookup"><span data-stu-id="fb144-105">This is useful if you need to register event handlers, pre-load data for the task pane, synchronize UI, or perform other tasks before the add-in is visible.</span></span>

[!include[Excel shared runtime note](../includes/note-requires-shared-runtime.md)]

## <a name="configure-your-add-in-to-load-when-the-document-opens"></a><span data-ttu-id="fb144-106">将外接程序配置为在文档打开时加载</span><span class="sxs-lookup"><span data-stu-id="fb144-106">Configure your add-in to load when the document opens</span></span>

<span data-ttu-id="fb144-107">下面的代码将加载项配置为在文档打开时加载并开始运行。</span><span class="sxs-lookup"><span data-stu-id="fb144-107">The following code configures your add-in to load and start running when the document is opened.</span></span>

```JavaScript
Office.addin.setStartupBehavior(Office.StartupBehavior.load);
```

> [!NOTE]
> <span data-ttu-id="fb144-108">`setStartupBehavior`方法是异步的。</span><span class="sxs-lookup"><span data-stu-id="fb144-108">The `setStartupBehavior` method is asynchronous.</span></span>

## <a name="configure-your-add-in-for-no-load-behavior-on-document-open"></a><span data-ttu-id="fb144-109">为打开的文档配置无加载行为的外接程序</span><span class="sxs-lookup"><span data-stu-id="fb144-109">Configure your add-in for no load behavior on document open</span></span>

<span data-ttu-id="fb144-110">以下代码将外接程序配置为在文档打开时启动。</span><span class="sxs-lookup"><span data-stu-id="fb144-110">The following code configures your add-in not to start when the document is opened.</span></span> <span data-ttu-id="fb144-111">而是在用户以某种方式（例如，选择功能区按钮或打开任务窗格）时启动。</span><span class="sxs-lookup"><span data-stu-id="fb144-111">Instead it will start when the user engages it in some way (such as choosing a ribbon button, or opening the task pane.)</span></span>

```JavaScript
Office.addin.setStartupBehavior(Office.StartupBehavior.none);
```

## <a name="get-the-current-load-behavior"></a><span data-ttu-id="fb144-112">获取当前加载行为</span><span class="sxs-lookup"><span data-stu-id="fb144-112">Get the current load behavior</span></span>

<span data-ttu-id="fb144-113">若要确定当前启动行为是什么，请运行以下函数，该函数将返回 StartupBehavior 对象。</span><span class="sxs-lookup"><span data-stu-id="fb144-113">To determine what the current startup behavior is, run the following function, which returns an Office.StartupBehavior object.</span></span>

```JavaScript
let behavior = await Office.addin.getStartupBehavior();
```

## <a name="how-to-run-code-when-the-document-opens"></a><span data-ttu-id="fb144-114">如何在文档打开时运行代码</span><span class="sxs-lookup"><span data-stu-id="fb144-114">How to run code when the document opens</span></span>

<span data-ttu-id="fb144-115">将外接程序配置为在打开文档时加载时，它将立即运行。</span><span class="sxs-lookup"><span data-stu-id="fb144-115">When your add-in is configured to load on document open, it will run immediately.</span></span> <span data-ttu-id="fb144-116">`Office.initialize`将调用事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="fb144-116">The `Office.initialize` event handler will be called.</span></span> <span data-ttu-id="fb144-117">将启动代码放在 `Office.initialize` 事件处理程序中。</span><span class="sxs-lookup"><span data-stu-id="fb144-117">Place your startup code in the `Office.initialize` event handler.</span></span>

<span data-ttu-id="fb144-118">下面的代码演示如何为活动工作表中的更改事件注册事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="fb144-118">The following code shows how to register an event handler for change events from the active worksheet.</span></span> <span data-ttu-id="fb144-119">如果将加载项配置为在打开文档时加载，此代码将在文档打开时注册事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="fb144-119">If you configure your add-in to load on document open, this code will register the event handler when the document is opened.</span></span> <span data-ttu-id="fb144-120">您可以在打开任务窗格之前处理更改事件。</span><span class="sxs-lookup"><span data-stu-id="fb144-120">You can handle change events before the task pane is opened.</span></span>


```JavaScript
//This is called as soon as the document opens.
//Put your startup code here.
Office.initialize = () => {
  // Add the event handler
  Excel.run(async context => {
    let sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.onChanged.add(onChange);

    await context.sync();
    console.log("A handler has been registered for the onChanged event.");
  });
};

/**
 * Handle the changed event from the worksheet.
 *
 * @param event The event information from Excel
 */
async function onChange(event) {
  return Excel.run(function(context) {
    return context.sync().then(function() {
      console.log("Change type of event: " + event.changeType);
      console.log("Address of event: " + event.address);
      console.log("Source of event: " + event.source);
    });
  });
}

```

## <a name="see-also"></a><span data-ttu-id="fb144-121">另请参阅</span><span class="sxs-lookup"><span data-stu-id="fb144-121">See also</span></span>

- [<span data-ttu-id="fb144-122">在 Excel 自定义函数和任务窗格教程之间共享数据和事件教程</span><span class="sxs-lookup"><span data-stu-id="fb144-122">Share data and events between Excel custom functions and task pane tutorial</span></span>](../tutorials/share-data-and-events-between-custom-functions-and-the-task-pane-tutorial.md)