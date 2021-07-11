---
title: 避免在循环中使用 context.sync
description: 了解如何使用拆分循环和相关对象模式避免在循环中调用 context.sync。
ms.date: 07/29/2020
localization_priority: Normal
ms.openlocfilehash: 64cfd5cd350746ba07e1a98986a4bd7811431475
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/09/2021
ms.locfileid: "53349138"
---
# <a name="avoid-using-the-contextsync-method-in-loops"></a><span data-ttu-id="68220-103">避免在循环中使用 context.sync</span><span class="sxs-lookup"><span data-stu-id="68220-103">Avoid using the context.sync method in loops</span></span>

> [!NOTE]
> <span data-ttu-id="68220-104">本文假定你已超出使用批处理系统与 Office 文档交互的四个特定于应用程序的 Office JavaScript API（适用于 &mdash; Excel、Word、OneNote 和 Visio）的开始阶段。 &mdash;</span><span class="sxs-lookup"><span data-stu-id="68220-104">This article assumes that you're beyond the beginning stage of working with at least one of the four application-specific Office JavaScript APIs&mdash;for Excel, Word, OneNote, and Visio&mdash;that use a batch system to interact with the Office document.</span></span> <span data-ttu-id="68220-105">特别是，你应了解调用功能， `context.sync` 并且应了解集合对象是什么。</span><span class="sxs-lookup"><span data-stu-id="68220-105">In particular, you should know what a call of `context.sync` does and you should know what a collection object is.</span></span> <span data-ttu-id="68220-106">如果你未处于该阶段，请从了解[JavaScript API Office，](../develop/understanding-the-javascript-api-for-office.md)以及该文章中"特定于应用程序"下的链接文档开始。</span><span class="sxs-lookup"><span data-stu-id="68220-106">If you're not at that stage, please start with [Understanding the Office JavaScript API](../develop/understanding-the-javascript-api-for-office.md) and the documentation linked to under "application-specific" in that article.</span></span>

<span data-ttu-id="68220-107">对于 Office 外接程序中对 Excel、Word、OneNote 和 Visio) 使用应用程序特定的 API 模型 (之一的一些编程方案，代码需要读取、写入或处理集合对象的每个成员中的一些属性。</span><span class="sxs-lookup"><span data-stu-id="68220-107">For some programming scenarios in Office Add-ins that use one of the application-specific API models (for Excel, Word, OneNote, and Visio), your code needs to read, write, or process some property from every member of a collection object.</span></span> <span data-ttu-id="68220-108">例如，Excel需要获取特定表格列内每个单元格的值的加载项，或需要突出显示文档中每个字符串实例的 Word 加载项。</span><span class="sxs-lookup"><span data-stu-id="68220-108">For example, an Excel add-in that needs to get the values of every cell in a particular table column or a Word add-in that needs to highlight every instance of a string in the document.</span></span> <span data-ttu-id="68220-109">您需要循环访问集合对象的 属性中的成员;但出于性能原因，您需要避免在循环的每次迭代中 `items` `context.sync` 调用。</span><span class="sxs-lookup"><span data-stu-id="68220-109">You need to iterate over the members in the `items` property of the collection object; but, for performance reasons, you need to avoid calling `context.sync` in every iteration of the loop.</span></span> <span data-ttu-id="68220-110">每次调用 都是从加载项到文档Office `context.sync` 行程。</span><span class="sxs-lookup"><span data-stu-id="68220-110">Every call of `context.sync` is a round trip from the add-in to the Office document.</span></span> <span data-ttu-id="68220-111">重复的往返会损害性能，尤其是在外接程序运行在 Office web 版，因为往返行程通过 Internet。</span><span class="sxs-lookup"><span data-stu-id="68220-111">Repeated round trips hurt performance, especially if the add-in is running in Office on the web because the round trips go across the internet.</span></span>

> [!NOTE]
> <span data-ttu-id="68220-112">本文中所有示例都使用循环，但所介绍的做法适用于可以循环访问数组的任何循环语句， `for` 包括：</span><span class="sxs-lookup"><span data-stu-id="68220-112">All examples in this article use `for` loops but the practices described apply to any loop statement that can iterate through an array, including the following:</span></span>
>
> - `for`
> - `for of`
> - `while`
> - `do while`
> 
> <span data-ttu-id="68220-113">它们还适用于函数传递给并应用于数组中的项目的任何数组方法，包括：</span><span class="sxs-lookup"><span data-stu-id="68220-113">They also apply to any array method to which a function is passed and applied to the items in the array, including the following:</span></span>
>
> - `Array.every`
> - `Array.forEach`
> - `Array.filter`
> - `Array.find`
> - `Array.findIndex`
> - `Array.map`
> - `Array.reduce`
> - `Array.reduceRight`
> - `Array.some`

## <a name="writing-to-the-document"></a><span data-ttu-id="68220-114">写入文档</span><span class="sxs-lookup"><span data-stu-id="68220-114">Writing to the document</span></span>

<span data-ttu-id="68220-115">在最简单的情况下，您只写入集合对象的成员，而不是读取它们的属性。</span><span class="sxs-lookup"><span data-stu-id="68220-115">In the simplest case, you are only writing to members of a collection object, not reading their properties.</span></span> <span data-ttu-id="68220-116">例如，以下代码在 Word 文档中以黄色突出显示每个"the"实例。</span><span class="sxs-lookup"><span data-stu-id="68220-116">For example, the following code highlights in yellow every instance of "the" in a Word document.</span></span>

> [!NOTE]
> <span data-ttu-id="68220-117">通常，在应用程序方法的结束"}"字符（如 、等）之前 (一个 `context.sync` `run` `Excel.run` `Word.run` 最终) 。</span><span class="sxs-lookup"><span data-stu-id="68220-117">It is generally a good practice to put have a final `context.sync` just before the closing "}" character of the application `run` method (such as `Excel.run`, `Word.run`, etc.).</span></span> <span data-ttu-id="68220-118">这是因为该方法在（并且仅在存在尚未同步的已排队命令）时执行最后一项操作时进行隐藏 `run` `context.sync` 调用。</span><span class="sxs-lookup"><span data-stu-id="68220-118">This is because the `run` method makes a hidden call of `context.sync` as the last thing it does if, and only if, there are queued commands that have not yet been synchronized.</span></span> <span data-ttu-id="68220-119">隐藏此调用这一事实可能会令人困惑，因此，我们通常建议添加显式 `context.sync` 。</span><span class="sxs-lookup"><span data-stu-id="68220-119">The fact that this call is hidden can be confusing, so we generally recommend that you add the explicit `context.sync`.</span></span> <span data-ttu-id="68220-120">但是，鉴于本文将调用最小化，添加一个完全不必要的最终 ，实际上会更 `context.sync` 令人困惑 `context.sync` 。</span><span class="sxs-lookup"><span data-stu-id="68220-120">However, given that this article is about minimizing calls of `context.sync`, it is actually more confusing to add an entirely unnecessary final `context.sync`.</span></span> <span data-ttu-id="68220-121">因此，在本文中，当 末尾没有未同步的命令时，我们会不进行介绍 `run` 。</span><span class="sxs-lookup"><span data-stu-id="68220-121">So, in this article, we leave it out when there are no unsynchronized commands at the end of the `run`.</span></span>

```javascript
Word.run(async function (context) {
    let startTime, endTime;
    const docBody = context.document.body;

    // search() returns an array of Ranges.
    const searchResults = docBody.search('the', { matchWholeWord: true });
    context.load(searchResults, 'items');
    await context.sync();

    // Record the system time.
    startTime = performance.now();

    for (var i = 0; i < searchResults.items.length; i++) {
      searchResults.items[i].font.highlightColor = '#FFFF00';

      await context.sync(); // SYNCHRONIZE IN EACH ITERATION
    }
    
    // await context.sync(); // SYNCHRONIZE AFTER THE LOOP

    // Record the system time again then calculate how long the operation took.
    endTime = performance.now();
    console.log("The operation took: " + (endTime - startTime) + " milliseconds.");
  })
}
```

<span data-ttu-id="68220-122">前面的代码在文档中使用 200 个实例的 Word on Windows 完成前一Windows。</span><span class="sxs-lookup"><span data-stu-id="68220-122">The preceding code took 1 full second to complete in a document with 200 instances of "the" in Word on Windows.</span></span> <span data-ttu-id="68220-123">但是，当在取消注释循环后将循环中的行注释掉且同一行时，该操作只需 `await context.sync();` 1/10 秒。</span><span class="sxs-lookup"><span data-stu-id="68220-123">But when the `await context.sync();` line inside the loop is commented out and the same line just after the loop is uncommented, the operation took only a 1/10th of a second.</span></span> <span data-ttu-id="68220-124">在Word web 版 (Edge 作为浏览器) 时，循环内同步需要 3 秒钟，在循环后同步只需 6/10 秒，大约快五倍。</span><span class="sxs-lookup"><span data-stu-id="68220-124">In Word on the web (with Edge as the browser), it took 3 full seconds with the synchronization inside the loop and only 6/10ths of a second with the synchronization after the loop, about five times faster.</span></span> <span data-ttu-id="68220-125">在包含 2000 个""实例的文档中，在 (Word web 版) 80 秒（循环内同步）中，在循环后仅同步 4 秒，大约快 20 倍。</span><span class="sxs-lookup"><span data-stu-id="68220-125">In a document with 2000 instances of "the", it took (in Word on the web) 80 seconds with the synchronization inside the loop and only 4 seconds with the synchronization after the loop, about 20 times faster.</span></span>

> [!NOTE]
> <span data-ttu-id="68220-126">值得一提的是，如果同步同时运行（只需从 的前面删除 关键字，就可以完成同步，循环内部同步版本能否更快地 `await` 执行 `context.sync()` ）。</span><span class="sxs-lookup"><span data-stu-id="68220-126">It's worth asking whether the synchronize-inside-the-loop version would execute faster if the synchronizations ran concurrently, which could be done by simply removing the `await` keyword from the front of the `context.sync()`.</span></span> <span data-ttu-id="68220-127">这会使运行时启动同步，然后立即启动循环的下一次迭代，而无需等待同步完成。</span><span class="sxs-lookup"><span data-stu-id="68220-127">This would cause the runtime to initiate the synchronization and then immediately start the next iteration of the loop without waiting for the synchronization to complete.</span></span> <span data-ttu-id="68220-128">但是，这不是一个比完全退出循环好的解决方案，原因 `context.sync` 如下：</span><span class="sxs-lookup"><span data-stu-id="68220-128">However, this is not as good a solution as moving the `context.sync` out of the loop entirely for these reasons:</span></span>
>
> - <span data-ttu-id="68220-129">与同步批处理作业中的命令排入队列一样，批处理作业本身在 Office 中排入队列，Office在队列中支持不超过 50 个批处理作业。</span><span class="sxs-lookup"><span data-stu-id="68220-129">Just as the commands in a synchronization batch job are queued, the batch jobs themselves are queued in Office, but Office supports no more than 50 batch jobs in the queue.</span></span> <span data-ttu-id="68220-130">其他任何操作都会引发错误。</span><span class="sxs-lookup"><span data-stu-id="68220-130">Any more triggers errors.</span></span> <span data-ttu-id="68220-131">因此，如果循环中迭代次数超过 50 次，则有可能超出队列大小。</span><span class="sxs-lookup"><span data-stu-id="68220-131">So, if there are more than 50 iterations in a loop, there is a chance that the queue size is exceeded.</span></span> <span data-ttu-id="68220-132">迭代次数越大，发生迭代的可能性越大。</span><span class="sxs-lookup"><span data-stu-id="68220-132">The greater the number of iterations, the greater the chance of this happening.</span></span> 
> - <span data-ttu-id="68220-133">"并发"并不意味着同时进行。</span><span class="sxs-lookup"><span data-stu-id="68220-133">"Concurrently" does not mean simultaneously.</span></span> <span data-ttu-id="68220-134">执行多个同步操作比执行一个同步操作要长。</span><span class="sxs-lookup"><span data-stu-id="68220-134">It would still take longer to execute multiple synchronization operations than to execute one.</span></span>
> - <span data-ttu-id="68220-135">不保证并发操作按其开始的顺序完成。</span><span class="sxs-lookup"><span data-stu-id="68220-135">Concurrent operations are not guaranteed to complete in the same order in which they started.</span></span> <span data-ttu-id="68220-136">在上一示例中，"the"一词的突出显示顺序无关紧要，但在一些方案中，必须按顺序处理集合中的项目。</span><span class="sxs-lookup"><span data-stu-id="68220-136">In the preceding example, it doesn't matter what order the  word "the" gets highlighted, but there are scenarios where it's important that the items in the collection be processed in order.</span></span>

## <a name="reading-values-from-the-document-with-the-split-loop-pattern"></a><span data-ttu-id="68220-137">使用拆分循环模式从文档读取值</span><span class="sxs-lookup"><span data-stu-id="68220-137">Reading values from the document with the split loop pattern</span></span>

<span data-ttu-id="68220-138">当代码在处理每个集合项时必须读取集合项的属性时，避免在循环内运行 `context.sync` 将更具挑战性。 </span><span class="sxs-lookup"><span data-stu-id="68220-138">Avoiding `context.sync`s inside a loop becomes more challenging when the code must *read* a property of the collection items as it processes each one.</span></span> <span data-ttu-id="68220-139">假设您的代码需要对 Word 文档中的所有内容控件进行重新访问，并记录与每个控件关联的第一段的文本。</span><span class="sxs-lookup"><span data-stu-id="68220-139">Suppose your code needs to iterate all the content controls in a Word document and log the text of the first paragraph associated with each control.</span></span> <span data-ttu-id="68220-140">编程方法可能会引导你循环访问控件、加载每个 (第一个) 段落的属性、调用 以用文档中的文本填充代理段落对象，然后记录它 `text` `context.sync` 。</span><span class="sxs-lookup"><span data-stu-id="68220-140">Your programming instincts might lead you to loop over the controls, load the `text` property of each (first) paragraph, call `context.sync` to populate the proxy paragraph object with the text from the document, and then log it.</span></span> <span data-ttu-id="68220-141">示例如下。</span><span class="sxs-lookup"><span data-stu-id="68220-141">The following is an example.</span></span>

```javascript
Word.run(async (context) => {
    const contentControls = context.document.contentControls.load('items');
    await context.sync();

    for (let i = 0; i < contentControls.items.length; i++) {
      const paragraph = contentControls.items[i].getRange('Whole').paragraphs.getFirst();
      paragraph.load('text');
      await context.sync();
      console.log(paragraph.text);
    }
});
```

<span data-ttu-id="68220-142">在此方案中，为了避免 在循环中出现 ，你应该使用我们调用拆分循环 `context.sync` **模式** 的模式。</span><span class="sxs-lookup"><span data-stu-id="68220-142">In this scenario, to avoid having a `context.sync` in a loop, you should use a pattern we call the **split loop** pattern.</span></span> <span data-ttu-id="68220-143">在获得模式的正式说明之前，让我们看一个具体模式示例。</span><span class="sxs-lookup"><span data-stu-id="68220-143">Let's see a concrete example of the pattern before we get to a formal description of it.</span></span> <span data-ttu-id="68220-144">下面将说明拆分循环模式如何应用于前面的代码段。</span><span class="sxs-lookup"><span data-stu-id="68220-144">Here's how the split loop pattern can be applied to the preceding code snippet.</span></span> <span data-ttu-id="68220-145">对于此代码，请注意以下事项。</span><span class="sxs-lookup"><span data-stu-id="68220-145">Note the following about this code.</span></span>

- <span data-ttu-id="68220-146">现在存在两个循环 `context.sync` ，两个循环之间出现，因此两个循环中 `context.sync` 都不存在。</span><span class="sxs-lookup"><span data-stu-id="68220-146">There are now two loops and the `context.sync` comes between them, so there's no `context.sync` inside either loop.</span></span>
- <span data-ttu-id="68220-147">第一个循环循环访问集合对象中的项目并加载属性，就像原始循环一样，但第一个循环无法记录段落文本，因为它不再包含 用于填充代理对象的属性的 `text` 。 `context.sync` `text` `paragraph`</span><span class="sxs-lookup"><span data-stu-id="68220-147">The first loop iterates through the items in the collection object and loads the `text` property just as the original loop did, but the first loop cannot log the paragraph text because it no longer contains a `context.sync` to populate the `text` property of the `paragraph` proxy object.</span></span> <span data-ttu-id="68220-148">相反，它会 `paragraph` 将对象添加到数组中。</span><span class="sxs-lookup"><span data-stu-id="68220-148">Instead, it adds the `paragraph` object to an array.</span></span>
- <span data-ttu-id="68220-149">第二个循环循环访问由第一个循环创建的数组，并记录 `text` 每个项目的 `paragraph` 。</span><span class="sxs-lookup"><span data-stu-id="68220-149">The second loop iterates through the array that was created by the first loop, and logs the `text` of each `paragraph` item.</span></span> <span data-ttu-id="68220-150">这是可能的，因为 两个循环之间的 填充 `context.sync` 了所有 `text` 属性。</span><span class="sxs-lookup"><span data-stu-id="68220-150">This is possible because the `context.sync` that came between the two loops populated all the `text` properties.</span></span>

```javascript
Word.run(async (context) => {
    const contentControls = context.document.contentControls.load("items");
    await context.sync();

    const firstParagraphsOfCCs = [];
    for (let i = 0; i < contentControls.items.length; i++) {
      const paragraph = contentControls.items[i].getRange('Whole').paragraphs.getFirst();
      paragraph.load('text');
      firstParagraphsOfCCs.push(paragraph);
    }

    await context.sync();

    for (let i = 0; i < firstParagraphsOfCCs.length; i++) {
      console.log(firstParagraphsOfCCs[i].text);
    }
});
```

<span data-ttu-id="68220-151">前面的示例建议以下过程将包含 的 循环 `context.sync` 转换为拆分循环模式。</span><span class="sxs-lookup"><span data-stu-id="68220-151">The preceding example suggests the following procedure for turning a loop that contains a `context.sync` into the split loop pattern.</span></span>

1. <span data-ttu-id="68220-152">将循环替换为两个循环。</span><span class="sxs-lookup"><span data-stu-id="68220-152">Replace the loop with two loops.</span></span>
2. <span data-ttu-id="68220-153">创建第一个循环来循环访问集合，将每个项目添加到数组中，同时加载代码需要读取的项目的任何属性。</span><span class="sxs-lookup"><span data-stu-id="68220-153">Create a first loop to iterate over the collection and add each item to an array while also loading any property of the item that your code needs to read.</span></span>
3. <span data-ttu-id="68220-154">第一个循环之后，调用 `context.sync` 以使用任何加载的属性填充代理对象。</span><span class="sxs-lookup"><span data-stu-id="68220-154">Following the first loop, call `context.sync` to populate the proxy objects with any loaded properties.</span></span>
4. <span data-ttu-id="68220-155">按照 第二个循环操作，循环访问第一个循环中创建的数组并 `context.sync` 读取加载的属性。</span><span class="sxs-lookup"><span data-stu-id="68220-155">Follow the `context.sync` with a second loop to iterate over the array created in the first loop and read the loaded properties.</span></span>

## <a name="processing-objects-in-the-document-with-the-correlated-objects-pattern"></a><span data-ttu-id="68220-156">使用相关对象模式处理文档中的对象</span><span class="sxs-lookup"><span data-stu-id="68220-156">Processing objects in the document with the correlated objects pattern</span></span>

<span data-ttu-id="68220-157">让我们考虑一个更复杂的方案，其中处理集合中的项需要不在项目本身内的数据。</span><span class="sxs-lookup"><span data-stu-id="68220-157">Let's consider a more complex scenario where processing the items in the collection requires data that isn't in the items themselves.</span></span> <span data-ttu-id="68220-158">方案设想一个 Word 外接程序，该外接程序对从具有一些样本文本的模板创建的文档进行操作。</span><span class="sxs-lookup"><span data-stu-id="68220-158">The scenario envisions a Word add-in that operates on documents created from a template with some boilerplate text.</span></span> <span data-ttu-id="68220-159">分散在文本中是以下占位符字符串的一个或多个实例："{Coordinator}"、"{Coordinatory}"和"{Manager}"。</span><span class="sxs-lookup"><span data-stu-id="68220-159">Scattered in the text are one or more instances of the following placeholder strings: "{Coordinator}", "{Deputy}", and "{Manager}".</span></span> <span data-ttu-id="68220-160">外接程序将每个占位符替换为某人的姓名。</span><span class="sxs-lookup"><span data-stu-id="68220-160">The add-in replaces each placeholder with some person's name.</span></span> <span data-ttu-id="68220-161">对于本文，外接程序的 UI 不十分重要。</span><span class="sxs-lookup"><span data-stu-id="68220-161">The UI of the add-in is not important to this article.</span></span> <span data-ttu-id="68220-162">例如，它可以有一个包含三个文本框的任务窗格，每个文本框都标记有一个占位符。</span><span class="sxs-lookup"><span data-stu-id="68220-162">For example, it could have a task pane with three text boxes, each labeled with one of the placeholders.</span></span> <span data-ttu-id="68220-163">用户在每个文本框中输入一个名称，然后 **按"替换** "按钮。</span><span class="sxs-lookup"><span data-stu-id="68220-163">The user enters a name in each text box and then presses a **Replace** button.</span></span> <span data-ttu-id="68220-164">按钮的处理程序创建一个数组，该数组将名称映射到占位符，然后用分配的名称替换每个占位符。</span><span class="sxs-lookup"><span data-stu-id="68220-164">The handler for the button creates an array that maps the names to the placeholders, and then replaces each placeholder with the assigned name.</span></span> 

<span data-ttu-id="68220-165">你无需实际通过此 UI 生成外接程序来试验代码。</span><span class="sxs-lookup"><span data-stu-id="68220-165">You don't need to actually produce an add-in with this UI to experiment with the code.</span></span> <span data-ttu-id="68220-166">可以使用 Script Lab[工具](../overview/explore-with-script-lab.md)构建重要代码的原型。</span><span class="sxs-lookup"><span data-stu-id="68220-166">You can use the [Script Lab tool](../overview/explore-with-script-lab.md) to prototype the important code.</span></span> <span data-ttu-id="68220-167">使用以下赋值语句创建映射数组。</span><span class="sxs-lookup"><span data-stu-id="68220-167">Use the following assignment statement to create the mapping array.</span></span>

```javascript
const jobMapping = [
        { job: "{Coordinator}", person: "Sally" },
        { job: "{Deputy}", person: "Bob" },
        { job: "{Manager}", person: "Kim" }
    ];
```

<span data-ttu-id="68220-168">以下代码显示如何在使用内部循环时，将每个占位符替换为其 `context.sync` 分配的名称。</span><span class="sxs-lookup"><span data-stu-id="68220-168">The following code shows how you might replace each placeholder with its assigned name if you used `context.sync` inside loops.</span></span>

```javascript
Word.run(async (context) => {

    for (let i = 0; i < jobMapping.length; i++) {
      let options = Word.SearchOptions.newObject(context);
      options.matchWildCards = false;
      let searchResults = context.document.body.search(jobMapping[i].job, options);
      searchResults.load('items');

      await context.sync(); 

      for (let j = 0; j < searchResults.items.length; j++) {
        searchResults.items[j].insertText(jobMapping[i].person, Word.InsertLocation.replace);

        await context.sync();
      }
    }
});
```

<span data-ttu-id="68220-169">在上一个代码中，有一个外部和一个内部循环。</span><span class="sxs-lookup"><span data-stu-id="68220-169">In the preceding code, there is an outer and an inner loop.</span></span> <span data-ttu-id="68220-170">其中每个都包含 `context.sync` 一个 。</span><span class="sxs-lookup"><span data-stu-id="68220-170">Each of them contains a `context.sync`.</span></span> <span data-ttu-id="68220-171">根据本文中第一个代码段，你可能会看到，内循环中的 可以仅移到内部 `context.sync` 循环之后。</span><span class="sxs-lookup"><span data-stu-id="68220-171">Based on the very first code snippet in this article, you probably see that the `context.sync` in the inner loop can simply be moved after the inner loop.</span></span> <span data-ttu-id="68220-172">但是，这仍将代码保留为 `context.sync` (，其中两个) 在外部循环中。</span><span class="sxs-lookup"><span data-stu-id="68220-172">But that would still leave the code with a `context.sync` (two of them actually) in the outer loop.</span></span> <span data-ttu-id="68220-173">以下代码演示如何从 `context.sync` 循环中删除。</span><span class="sxs-lookup"><span data-stu-id="68220-173">The following code shows how you can remove `context.sync` from the loops.</span></span> <span data-ttu-id="68220-174">我们将讨论以下代码。</span><span class="sxs-lookup"><span data-stu-id="68220-174">We discuss the code below.</span></span>

```javascript
Word.run(async (context) => {

    const allSearchResults = [];
    for (let i = 0; i < jobMapping.length; i++) {
      let options = Word.SearchOptions.newObject(context);
      options.matchWildCards = false;
      let searchResults = context.document.body.search(jobMapping[i].job, options);
      searchResults.load('items');
      let correlatedSearchResult = {
        rangesMatchingJob: searchResults,
        personAssignedToJob: jobMapping[i].person
      }
      allSearchResults.push(correlatedSearchResult);
    }

    await context.sync()

    for (let i = 0; i < allSearchResults.length; i++) {
      let correlatedObject = allSearchResults[i];

      for (let j = 0; j < correlatedObject.rangesMatchingJob.items.length; j++) {
        let targetRange = correlatedObject.rangesMatchingJob.items[j];
        let name = correlatedObject.personAssignedToJob;
        targetRange.insertText(name, Word.InsertLocation.replace);
      }
    }

    await context.sync();
});
```

<span data-ttu-id="68220-175">请注意，代码使用拆分循环模式：</span><span class="sxs-lookup"><span data-stu-id="68220-175">Note the code uses the split loop pattern:</span></span>

- <span data-ttu-id="68220-176">上例中的外部循环已拆分为两个。</span><span class="sxs-lookup"><span data-stu-id="68220-176">The outer loop from the preceding example has been split into two.</span></span> <span data-ttu-id="68220-177"> (第二个循环有一个内部循环，这是预期的，因为代码将循环遍历一组作业 (或占位符) 并且该循环将在此集合中迭代匹配的范围。) </span><span class="sxs-lookup"><span data-stu-id="68220-177">(The second loop has an inner loop, which is expected because the code is iterating over a set of jobs (or placeholders) and within that set it is iterating over the matching ranges.)</span></span>
- <span data-ttu-id="68220-178">每个主 `context.sync` 循环后都有 一个 ，但在任何 `context.sync` 循环内没有。</span><span class="sxs-lookup"><span data-stu-id="68220-178">There is a `context.sync` after each major loop, but no `context.sync` inside any loop.</span></span>
- <span data-ttu-id="68220-179">第二个主要循环循环访问第一个循环中创建的数组。</span><span class="sxs-lookup"><span data-stu-id="68220-179">The second major loop iterates through an array that is created in the first loop.</span></span>

<span data-ttu-id="68220-180">但是，第一个循环中创建的数组并不只包含一个 Office 对象，正如第一个循环使用拆分循环模式读取文档中[的值一样](#reading-values-from-the-document-with-the-split-loop-pattern)。</span><span class="sxs-lookup"><span data-stu-id="68220-180">But the array created in the first loop does *not* contain only an Office object as the first loop did in the section [Reading values from the document with the split loop pattern](#reading-values-from-the-document-with-the-split-loop-pattern).</span></span> <span data-ttu-id="68220-181">这是因为处理 Word Range 对象所需的某些信息不在 Range 对象本身中，而是来自 `jobMapping` 数组。</span><span class="sxs-lookup"><span data-stu-id="68220-181">This is because some of the information needed to process the Word Range objects is not in the Range objects themselves but instead comes from the `jobMapping` array.</span></span>

<span data-ttu-id="68220-182">因此，第一个循环中创建的数组中的对象是具有两个属性的自定义对象。</span><span class="sxs-lookup"><span data-stu-id="68220-182">So, the objects in the array created in the first loop are custom objects that have two properties.</span></span> <span data-ttu-id="68220-183">第一个数组是匹配特定职务 (（即占位符字符串) ）的 Word 范围数组，第二个字符串提供分配给该工作的人的姓名。</span><span class="sxs-lookup"><span data-stu-id="68220-183">The first is an array of Word Ranges that match a specific job title (that is, a placeholder string) and the second is a string that provides the name of the person assigned to the job.</span></span> <span data-ttu-id="68220-184">这使得最后一个循环易于编写且易于阅读，因为处理给定区域所需的全部信息都包含在包含该范围的同一自定义对象中。</span><span class="sxs-lookup"><span data-stu-id="68220-184">This makes the final loop easy to write and easy to read because all of the information needed to process a given range is contained in the same custom object that contains the range.</span></span> <span data-ttu-id="68220-185">应替换 _**correlatedObject**.rangesMatchingJob.items[j]_ 的名称是同一对象的另一个属性 _**：correlatedObject**.personAssignedToJob_。</span><span class="sxs-lookup"><span data-stu-id="68220-185">The name that should replace _**correlatedObject**.rangesMatchingJob.items[j]_ is the other property of the same object: _**correlatedObject**.personAssignedToJob_.</span></span>

<span data-ttu-id="68220-186">我们将此变体称为拆分循环模式 **的相关对象** 模式。</span><span class="sxs-lookup"><span data-stu-id="68220-186">We call this variation of the split loop pattern the **correlated objects** pattern.</span></span> <span data-ttu-id="68220-187">一般概念是，第一个循环创建一个自定义对象数组。</span><span class="sxs-lookup"><span data-stu-id="68220-187">The general idea is that the first loop creates an array of custom objects.</span></span> <span data-ttu-id="68220-188">每个对象都有一个属性值，该属性是 Office 集合对象 (或此类项目数组中的) 。</span><span class="sxs-lookup"><span data-stu-id="68220-188">Each object has a property whose value is one of the items in an Office collection object (or an array of such items).</span></span> <span data-ttu-id="68220-189">自定义对象具有其他属性，每个属性都提供在最终循环中处理Office对象时所需的信息。</span><span class="sxs-lookup"><span data-stu-id="68220-189">The custom object has other properties, each of which provides information needed to process the Office objects in the final loop.</span></span> <span data-ttu-id="68220-190">有关指向 [自定义关联](#other-examples-of-these-patterns) 对象具有两个以上属性的示例的链接，请参阅这些模式的其他示例一节。</span><span class="sxs-lookup"><span data-stu-id="68220-190">See the section [Other examples of these patterns](#other-examples-of-these-patterns) for a link to an example where the custom correlating object has more than two properties.</span></span>

<span data-ttu-id="68220-191">另一个警告：有时，仅创建自定义关联对象的数组需要多个循环。</span><span class="sxs-lookup"><span data-stu-id="68220-191">One further caveat: sometimes it takes more than one loop just to create the array of custom correlating objects.</span></span> <span data-ttu-id="68220-192">如果需要读取一个集合对象中每个成员的属性，Office收集将用于处理另一个集合对象的信息，则可能会发生这种情况。</span><span class="sxs-lookup"><span data-stu-id="68220-192">This can happen if you need to read a property of each member of one Office collection object just to gather information that will be used to process another collection object.</span></span> <span data-ttu-id="68220-193"> (例如，您的代码需要读取 Excel 表中所有列的标题，因为您的外接程序将基于该列的标题将数字格式应用于某些列的单元格。) 但您可以始终在循环之间保留 ，而不是在循环中保留。 `context.sync`</span><span class="sxs-lookup"><span data-stu-id="68220-193">(For example, your code needs to read the titles of all the columns in an Excel table because your add-in is going to apply a number format to the cells of some columns based on that column's title.) But you can always keep the `context.sync`s between the loops, rather than in a loop.</span></span> <span data-ttu-id="68220-194">有关示例 [，请参阅这些模式的其他](#other-examples-of-these-patterns) 示例部分。</span><span class="sxs-lookup"><span data-stu-id="68220-194">See the section [Other examples of these patterns](#other-examples-of-these-patterns) for an example.</span></span>

## <a name="other-examples-of-these-patterns"></a><span data-ttu-id="68220-195">这些模式的其他示例</span><span class="sxs-lookup"><span data-stu-id="68220-195">Other examples of these patterns</span></span>

- <span data-ttu-id="68220-196">有关使用循环Excel一个非常简单的示例，请参阅此 Stack Overflow 问题的接受答案：在 context.sync 之前，是否可能将多个 `Array.forEach` [context.load](https://stackoverflow.com/questions/44459604/is-it-possible-to-queue-more-than-one-context-load-before-context-sync)排入队列？</span><span class="sxs-lookup"><span data-stu-id="68220-196">For a very simple example for Excel that uses `Array.forEach` loops, see the accepted answer to this Stack Overflow question: [Is it possible to queue more than one context.load before context.sync?](https://stackoverflow.com/questions/44459604/is-it-possible-to-queue-more-than-one-context-load-before-context-sync)</span></span>
- <span data-ttu-id="68220-197">有关使用循环且不使用语法的 Word 的简单示例，请参阅此 Stack Overflow 问题的接受答案：使用 `Array.forEach` `async` / `await` [Office JavaScript API](https://stackoverflow.com/questions/58422113/iterating-over-all-paragraphs-with-content-controls-with-office-javascript-api)遍历包含内容控件的所有段落。</span><span class="sxs-lookup"><span data-stu-id="68220-197">For a simple example for Word that uses `Array.forEach` loops and doesn't use `async`/`await` syntax, see the accepted answer to this Stack Overflow question: [Iterating over all paragraphs with content controls with Office JavaScript API](https://stackoverflow.com/questions/58422113/iterating-over-all-paragraphs-with-content-controls-with-office-javascript-api).</span></span>
- <span data-ttu-id="68220-198">有关使用 TypeScript 编写的 Word 示例，请参阅示例 [Word 外接程序 Angular2](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker)样式检查器，尤其是文件 [word.document.service.ts](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker/blob/master/app/services/word-document/word.document.service.ts)。</span><span class="sxs-lookup"><span data-stu-id="68220-198">For an example for Word that is written in TypeScript, see the sample [Word Add-in Angular2 Style Checker](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker), especially the file [word.document.service.ts](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker/blob/master/app/services/word-document/word.document.service.ts).</span></span> <span data-ttu-id="68220-199">它混合了 `for` 和 `Array.forEach` 循环。</span><span class="sxs-lookup"><span data-stu-id="68220-199">It has a mixture of `for` and `Array.forEach` loops.</span></span>
- <span data-ttu-id="68220-200">对于高级 Word 示例，将[此 gist](https://gist.github.com/9c5a803e52480ec7f00bb3224292e0ab)导入[Script Lab 工具。](../overview/explore-with-script-lab.md)</span><span class="sxs-lookup"><span data-stu-id="68220-200">For an advanced Word sample, import [this gist](https://gist.github.com/9c5a803e52480ec7f00bb3224292e0ab) into the [Script Lab tool](../overview/explore-with-script-lab.md).</span></span> <span data-ttu-id="68220-201">有关使用 gist 的上下文，请参阅 Stack Overflow 问题的接受答案替换文本 [后文档未同步](https://stackoverflow.com/questions/48227941/document-not-in-sync-after-replace-text)。</span><span class="sxs-lookup"><span data-stu-id="68220-201">For context in using the gist, see the accepted answer to the Stack Overflow question [Document not in sync after replace text](https://stackoverflow.com/questions/48227941/document-not-in-sync-after-replace-text).</span></span> <span data-ttu-id="68220-202">此示例创建一个具有三对象类型关联的自定义关联对象。</span><span class="sxs-lookup"><span data-stu-id="68220-202">This sample creates a custom correlating object type that has three properties.</span></span> <span data-ttu-id="68220-203">它总共使用三个循环来构造相关对象的数组，并另外使用两个循环执行最终处理。</span><span class="sxs-lookup"><span data-stu-id="68220-203">It uses a total of three loops to construct the array of correlated objects, and two more loops to do the final processing.</span></span> <span data-ttu-id="68220-204">有 和 `for` `Array.forEach` 循环的混合。</span><span class="sxs-lookup"><span data-stu-id="68220-204">There are a mixture of `for` and `Array.forEach` loops.</span></span>
- <span data-ttu-id="68220-205">尽管不严格是拆分循环或关联对象模式的示例，但还有一个高级 Excel 示例演示如何将一组单元格值转换为仅包含一个 的其他货币 `context.sync` 。</span><span class="sxs-lookup"><span data-stu-id="68220-205">Although not strictly an example of the split loop or correlated objects patterns, there is an advanced Excel sample that shows how to convert a set of cell values to other currencies with just a single `context.sync`.</span></span> <span data-ttu-id="68220-206">若要试用，请打开 [Script Lab 工具](../overview/explore-with-script-lab.md)并导航到 **"货币转换器"** 示例。</span><span class="sxs-lookup"><span data-stu-id="68220-206">To try it, open the [Script Lab tool](../overview/explore-with-script-lab.md) and navigate to the **Currency Converter** sample.</span></span>

## <a name="when-should-you-not-use-the-patterns-in-this-article"></a><span data-ttu-id="68220-207">何时 *不应* 使用本文中的模式？</span><span class="sxs-lookup"><span data-stu-id="68220-207">When should you *not* use the patterns in this article?</span></span>

<span data-ttu-id="68220-208">Excel在给定调用 中读取的数据不能超过 5 `context.sync` MB。</span><span class="sxs-lookup"><span data-stu-id="68220-208">Excel cannot read more than 5 MB of data in a given call of `context.sync`.</span></span> <span data-ttu-id="68220-209">如果超出此限制，将引发错误。</span><span class="sxs-lookup"><span data-stu-id="68220-209">If this limit is exceeded, an error is thrown.</span></span> <span data-ttu-id="68220-210"> (有关详细信息，请参阅 Office 外接程序的资源限制和性能优化的["Excel](resource-limits-and-performance-optimization.md#excel-add-ins)外接程序部分"。) 接近此限制的情况很少见，但如果外接程序可能会发生这种情况，则代码不应在一个循环中加载所有数据，而是使用 执行循环 `context.sync` 。</span><span class="sxs-lookup"><span data-stu-id="68220-210">(See the "Excel add-ins section" of [Resource limits and performance optimization for Office Add-ins](resource-limits-and-performance-optimization.md#excel-add-ins) for more information.) It's very rare that this limit is approached, but if there's a chance that this will happen with your add-in, then your code should *not* load all the data in a single loop and follow the loop with a `context.sync`.</span></span> <span data-ttu-id="68220-211">但是，您仍应避免在集合 `context.sync` 对象的循环的每次迭代中都有 。</span><span class="sxs-lookup"><span data-stu-id="68220-211">But you still should avoid having a `context.sync` in every iteration of a loop over a collection object.</span></span> <span data-ttu-id="68220-212">相反，请定义集合中项的子集，并循环遍历每个子集，在循环之间使用 `context.sync` 。</span><span class="sxs-lookup"><span data-stu-id="68220-212">Instead, define subsets of the items in the collection and loop over each subset in turn, with a `context.sync` between the loops.</span></span> <span data-ttu-id="68220-213">可以使用循环遍历子集的外部循环来构造此结构，并包含每个外部 `context.sync` 迭代中的 。</span><span class="sxs-lookup"><span data-stu-id="68220-213">You could structure this with an outer loop that iterates over the subsets and contains the `context.sync` in each of these outer iterations.</span></span>
