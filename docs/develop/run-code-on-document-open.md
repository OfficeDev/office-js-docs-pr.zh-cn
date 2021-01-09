---
title: 文档打开时在 Office 外接程序中运行代码
description: 了解如何在文档打开时在 Office 外接程序中运行代码。
ms.date: 12/28/2020
localization_priority: Normal
ms.openlocfilehash: 1655c053a4fa6f92aae95f2155991fa4f7f7a5a7
ms.sourcegitcommit: 545888b08f57bb1babb05ccfd83b2b3286bdad5c
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/08/2021
ms.locfileid: "49789217"
---
# <a name="run-code-in-your-office-add-in-when-the-document-opens"></a><span data-ttu-id="72b61-103">文档打开时在 Office 外接程序中运行代码</span><span class="sxs-lookup"><span data-stu-id="72b61-103">Run code in your Office Add-in when the document opens</span></span>

[!include[Shared JavaScript runtime requirements](../includes/shared-runtime-requirements-note.md)]

<span data-ttu-id="72b61-104">您可以将 Office 外接程序配置为在文档打开后加载和运行代码。</span><span class="sxs-lookup"><span data-stu-id="72b61-104">You can configure your Office Add-in to load and run code as soon as the document is opened.</span></span> <span data-ttu-id="72b61-105">如果你需要在加载项可见之前注册事件处理程序、预加载任务窗格数据、同步 UI 或执行其他任务，这将非常有用。</span><span class="sxs-lookup"><span data-stu-id="72b61-105">This is useful if you need to register event handlers, pre-load data for the task pane, synchronize UI, or perform other tasks before the add-in is visible.</span></span>

[!include[Shared runtime note](../includes/note-requires-shared-runtime.md)]

## <a name="configure-your-add-in-to-load-when-the-document-opens"></a><span data-ttu-id="72b61-106">将加载项配置为在文档打开时加载</span><span class="sxs-lookup"><span data-stu-id="72b61-106">Configure your add-in to load when the document opens</span></span>

<span data-ttu-id="72b61-107">以下代码将外接程序配置为在打开文档时加载并开始运行。</span><span class="sxs-lookup"><span data-stu-id="72b61-107">The following code configures your add-in to load and start running when the document is opened.</span></span>

```JavaScript
Office.addin.setStartupBehavior(Office.StartupBehavior.load);
```

> [!NOTE]
> <span data-ttu-id="72b61-108">方法是 `setStartupBehavior` 异步的。</span><span class="sxs-lookup"><span data-stu-id="72b61-108">The `setStartupBehavior` method is asynchronous.</span></span>

## <a name="configure-your-add-in-for-no-load-behavior-on-document-open"></a><span data-ttu-id="72b61-109">配置加载项以在文档打开时不加载行为</span><span class="sxs-lookup"><span data-stu-id="72b61-109">Configure your add-in for no load behavior on document open</span></span>

<span data-ttu-id="72b61-110">下面的代码将加载项配置为在打开文档时不启动。</span><span class="sxs-lookup"><span data-stu-id="72b61-110">The following code configures your add-in not to start when the document is opened.</span></span> <span data-ttu-id="72b61-111">相反，它将在用户以某种方式参与时启动，例如选择功能区按钮或打开任务窗格。</span><span class="sxs-lookup"><span data-stu-id="72b61-111">Instead, it will start when the user engages it in some way, such as choosing a ribbon button or opening the task pane.</span></span>

```JavaScript
Office.addin.setStartupBehavior(Office.StartupBehavior.none);
```

## <a name="get-the-current-load-behavior"></a><span data-ttu-id="72b61-112">获取当前加载行为</span><span class="sxs-lookup"><span data-stu-id="72b61-112">Get the current load behavior</span></span>

<span data-ttu-id="72b61-113">若要确定当前的启动行为是什么，请运行以下函数，该函数将返回 `Office.StartupBehavior` 一个对象。</span><span class="sxs-lookup"><span data-stu-id="72b61-113">To determine what the current startup behavior is, run the following function, which returns an `Office.StartupBehavior` object.</span></span>

```JavaScript
let behavior = await Office.addin.getStartupBehavior();
```

## <a name="how-to-run-code-when-the-document-opens"></a><span data-ttu-id="72b61-114">如何在文档打开时运行代码</span><span class="sxs-lookup"><span data-stu-id="72b61-114">How to run code when the document opens</span></span>

<span data-ttu-id="72b61-115">当加载项配置为在文档打开时加载时，它将立即运行。</span><span class="sxs-lookup"><span data-stu-id="72b61-115">When your add-in is configured to load on document open, it will run immediately.</span></span> <span data-ttu-id="72b61-116">将 `Office.initialize` 调用事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="72b61-116">The `Office.initialize` event handler will be called.</span></span> <span data-ttu-id="72b61-117">将启动代码放在 `Office.initialize` 或 `Office.onReady` 事件处理程序中。</span><span class="sxs-lookup"><span data-stu-id="72b61-117">Place your startup code in the `Office.initialize` or `Office.onReady` event handler.</span></span>

<span data-ttu-id="72b61-118">以下 Excel 加载项代码显示如何为活动工作表中的更改事件注册事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="72b61-118">The following Excel add-in code shows how to register an event handler for change events from the active worksheet.</span></span> <span data-ttu-id="72b61-119">如果将加载项配置为在文档打开时加载，则此代码将在文档打开时注册事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="72b61-119">If you configure your add-in to load on document open, this code will register the event handler when the document is opened.</span></span> <span data-ttu-id="72b61-120">您可以在打开任务窗格之前处理更改事件。</span><span class="sxs-lookup"><span data-stu-id="72b61-120">You can handle change events before the task pane is opened.</span></span>

```JavaScript
// This is called as soon as the document opens.
// Put your startup code here.
Office.initialize = () => {
  // Add the event handler.
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

<span data-ttu-id="72b61-121">以下 PowerPoint 加载项代码显示如何为 PowerPoint 文档中的选择更改事件注册事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="72b61-121">The following PowerPoint add-in code shows how to register an event handler for selection change events from the PowerPoint document.</span></span> <span data-ttu-id="72b61-122">如果将加载项配置为在文档打开时加载，则此代码将在文档打开时注册事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="72b61-122">If you configure your add-in to load on document open, this code will register the event handler when the document is opened.</span></span> <span data-ttu-id="72b61-123">您可以在打开任务窗格之前处理更改事件。</span><span class="sxs-lookup"><span data-stu-id="72b61-123">You can handle change events before the task pane is opened.</span></span>

```JavaScript
// This is called as soon as the document opens.
// Put your startup code here.
Office.onReady(info => {
  if (info.host === Office.HostType.PowerPoint) {
    Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, onChange);
    console.log("A handler has been registered for the onChanged event.");
  }
});

/**
 * Handle the changed event from the PowerPoint document.
 *
 * @param event The event information from PowerPoint
 */
async function onChange(event) {
  console.log("Change type of event: " + event.type);
}
```

## <a name="see-also"></a><span data-ttu-id="72b61-124">另请参阅</span><span class="sxs-lookup"><span data-stu-id="72b61-124">See also</span></span>

- [<span data-ttu-id="72b61-125">将 Office 外接程序配置为使用共享的 JavaScript 运行时</span><span class="sxs-lookup"><span data-stu-id="72b61-125">Configure your Office Add-in to use a shared JavaScript runtime</span></span>](configure-your-add-in-to-use-a-shared-runtime.md)
- [<span data-ttu-id="72b61-126">在 Excel 自定义函数和任务窗格教程之间共享数据和事件</span><span class="sxs-lookup"><span data-stu-id="72b61-126">Share data and events between Excel custom functions and task pane tutorial</span></span>](../tutorials/share-data-and-events-between-custom-functions-and-the-task-pane-tutorial.md)
- [<span data-ttu-id="72b61-127">使用 Excel JavaScript API 处理事件</span><span class="sxs-lookup"><span data-stu-id="72b61-127">Work with Events using the Excel JavaScript API</span></span>](../excel/excel-add-ins-events.md)
