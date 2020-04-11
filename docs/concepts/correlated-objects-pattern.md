---
title: 避免在循环中使用 context。 sync 方法
description: 了解如何使用拆分循环和相关的对象模式以避免调用上下文。循环中的同步。
ms.date: 04/09/2020
localization_priority: Normal
ms.openlocfilehash: 0f967b07b3ccf323321779676021c53c81102f83
ms.sourcegitcommit: 76552b3e5725d9112c772595971b922c295e6b4c
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/10/2020
ms.locfileid: "43225988"
---
# <a name="avoid-using-the-contextsync-method-in-loops"></a><span data-ttu-id="05848-103">避免在循环中使用 context。 sync 方法</span><span class="sxs-lookup"><span data-stu-id="05848-103">Avoid using the context.sync method in loops</span></span>

> [!NOTE]
> <span data-ttu-id="05848-104">本文假定您不在第一阶段使用至少使用用于 Excel、Word、OneNote 和 Visio&mdash;&mdash;的四个特定于 Excel 的 Office JavaScript api 中的一个，这些 api 使用批处理系统与 Office 文档进行交互。</span><span class="sxs-lookup"><span data-stu-id="05848-104">This article assumes that you're beyond the beginning stage of working with at least one of the four host-specific Office JavaScript APIs&mdash;for Excel, Word, OneNote, and Visio&mdash;that use a batch system to interact with the Office document.</span></span> <span data-ttu-id="05848-105">特别是，您应该知道什么是调用`context.sync` ，您应该知道什么是集合对象。</span><span class="sxs-lookup"><span data-stu-id="05848-105">In particular, you should know what a call of `context.sync` does and you should know what a collection object is.</span></span> <span data-ttu-id="05848-106">如果你不在这一阶段，请先了解本文中的 "特定于主机" 下的 " [Office JAVASCRIPT API](../develop/understanding-the-javascript-api-for-office.md) " 和 "文档"。</span><span class="sxs-lookup"><span data-stu-id="05848-106">If you're not at that stage, please start with [Understanding the Office JavaScript API](../develop/understanding-the-javascript-api-for-office.md) and the documentation linked to under "host-specific" in that article.</span></span>

<span data-ttu-id="05848-107">对于使用特定于主机的 API 模型之一（针对 Excel、Word、OneNote 和 Visio）的 Office 外接程序中的某些编程方案，您的代码需要从集合对象的每个成员中读取、写入或处理某些属性。</span><span class="sxs-lookup"><span data-stu-id="05848-107">For some programming scenarios in Office Add-ins that use one of the host-specific API models (for Excel, Word, OneNote, and Visio), your code needs to read, write, or process some property from every member of a collection object.</span></span> <span data-ttu-id="05848-108">例如，需要获取特定表列或 Word 外接程序中每个单元格的值的 Excel 加载项，需要突出显示文档中每个字符串的实例。</span><span class="sxs-lookup"><span data-stu-id="05848-108">For example, an Excel add-in that needs to get the values of every cell in a particular table column or a Word add-in that needs to highlight every instance of a string in the document.</span></span> <span data-ttu-id="05848-109">您需要在集合对象的`items`属性中循环访问这些成员;但是，出于性能方面的考虑，您需要避免`context.sync`在循环的每个迭代中调用。</span><span class="sxs-lookup"><span data-stu-id="05848-109">You need to iterate over the members in the `items` property of the collection object; but, for performance reasons, you need to avoid calling `context.sync` in every iteration of the loop.</span></span> <span data-ttu-id="05848-110">每次调用`context.sync`的是从外接端到 Office 文档的一种往返行程。</span><span class="sxs-lookup"><span data-stu-id="05848-110">Every call of `context.sync` is a round trip from the add-in to the Office document.</span></span> <span data-ttu-id="05848-111">重复往返行程会影响性能，尤其是当加载项在 web 上的 Office 中运行时，由于往返行程跨 internet 进行。</span><span class="sxs-lookup"><span data-stu-id="05848-111">Repeated round trips hurt performance, especially if the add-in is running in Office on the web because the round trips go across the internet.</span></span>

> [!NOTE]
> <span data-ttu-id="05848-112">本文中的所有示例都`for`使用循环，但所述的实践适用于可循环访问数组的任何循环语句，其中包括以下内容：</span><span class="sxs-lookup"><span data-stu-id="05848-112">All examples in this article use `for` loops but the practices described apply to any loop statement that can iterate through an array, including the following:</span></span>
>
> - `for`
> - `for of`
> - `while`
> - `do while`
> 
> <span data-ttu-id="05848-113">它们还适用于将函数传递和应用于数组中的项的任何数组方法，包括以下内容：</span><span class="sxs-lookup"><span data-stu-id="05848-113">They also apply to any array method to which a function is passed and applied to the items in the array, including the following:</span></span>
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

## <a name="writing-to-the-document"></a><span data-ttu-id="05848-114">写入文档</span><span class="sxs-lookup"><span data-stu-id="05848-114">Writing to the document</span></span>

<span data-ttu-id="05848-115">在最简单的情况下，只写入集合对象的成员，而不是读取其属性。</span><span class="sxs-lookup"><span data-stu-id="05848-115">In the simplest case, you are only writing to members of a collection object, not reading their properties.</span></span> <span data-ttu-id="05848-116">例如，下面的代码在 Word 文档中突出显示了每个 "the" 实例的黄色。</span><span class="sxs-lookup"><span data-stu-id="05848-116">For example, the following code highlights in yellow every instance of "the" in a Word document.</span></span> 

> [!NOTE]
> <span data-ttu-id="05848-117">通常，最好`context.sync`先在主机`run`方法的结尾 "}" 字符前加上 final （如`Excel.run` `Word.run`，等）。</span><span class="sxs-lookup"><span data-stu-id="05848-117">It is generally a good practice to put have a final `context.sync` just before the closing "}" character of the host `run` method (such as `Excel.run`, `Word.run`, etc.).</span></span> <span data-ttu-id="05848-118">这是因为此`run`方法会将`context.sync`作为最后一件事情的隐藏调用作为最后一件事情，并且只有在已排队的命令尚未同步的情况下。</span><span class="sxs-lookup"><span data-stu-id="05848-118">This is because the `run` method makes a hidden call of `context.sync` as the last thing it does if, and only if, there are queued commands that have not yet been synchronized.</span></span> <span data-ttu-id="05848-119">此调用是隐藏的这一事实可能会造成混淆，因此我们通常建议您添加显式`context.sync`。</span><span class="sxs-lookup"><span data-stu-id="05848-119">The fact that this call is hidden can be confusing, so we generally recommend that you add the explicit `context.sync`.</span></span> <span data-ttu-id="05848-120">但是，假设本文涉及最小化的调用， `context.sync`则添加完全不必要的最终版本`context.sync`会更加容易混淆。</span><span class="sxs-lookup"><span data-stu-id="05848-120">However, given that this article is about minimizing calls of `context.sync`, it is actually more confusing to add an entirely unnecessary final `context.sync`.</span></span> <span data-ttu-id="05848-121">因此，在本文的结尾处没有未同步的命令时，我们会将其保留下来`run`。</span><span class="sxs-lookup"><span data-stu-id="05848-121">So, in this article, we leave it out when there are no unsynchronized commands at the end of the `run`.</span></span> 

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

<span data-ttu-id="05848-122">上面的代码花了1个完整的秒，在 Windows 中的 Word 中有一个包含200个实例 "the" 的文档。</span><span class="sxs-lookup"><span data-stu-id="05848-122">The preceding code took 1 full second to complete in a document with 200 instances of "the" in Word on Windows.</span></span> <span data-ttu-id="05848-123">但是，当`await context.sync();`循环中的行被注释掉并在循环 uncommented 后的相同行时，该操作只需 1/10 秒。</span><span class="sxs-lookup"><span data-stu-id="05848-123">But when the `await context.sync();` line inside the loop is commented out and the same line just after the loop is uncommented, the operation took only a 1/10th of a second.</span></span> <span data-ttu-id="05848-124">在 web 上的 Word 中（使用边缘作为浏览器），循环内的同步花费了3个完整的秒，并且在循环后的同步速度仅为 6/10ths，而在循环后的速度约为5倍。</span><span class="sxs-lookup"><span data-stu-id="05848-124">In Word on the web (with Edge as the browser), it took 3 full seconds with the synchronization inside the loop and only 6/10ths of a second with the synchronization after the loop, about five times faster.</span></span> <span data-ttu-id="05848-125">在包含2000实例 "the" 的文档中，使用循环中的同步（在 web 中为 "网页"）80秒，并且在循环后的同步只需4秒，速度将加快20倍。</span><span class="sxs-lookup"><span data-stu-id="05848-125">In a document with 2000 instances of "the", it took (in Word on the web) 80 seconds with the synchronization inside the loop and only 4 seconds with the synchronization after the loop, about 20 times faster.</span></span>

> [!NOTE]
> <span data-ttu-id="05848-126">如果同步并发运行，则需要询问是否会更快地执行同步内部循环版本，这只需从的`await` `context.sync()`前面删除关键字即可完成。</span><span class="sxs-lookup"><span data-stu-id="05848-126">It's worth asking whether the synchronize-inside-the-loop version would execute faster if the synchronizations ran concurrently, which could be done by simply removing the `await` keyword from the front of the `context.sync()`.</span></span> <span data-ttu-id="05848-127">这将导致运行时启动同步，然后立即开始循环的下一个迭代，而无需等待同步完成。</span><span class="sxs-lookup"><span data-stu-id="05848-127">This would cause the runtime to initiate the synchronization and then immediately start the next iteration of the loop without waiting for the synchronization to complete.</span></span> <span data-ttu-id="05848-128">但是，这并不像出于这些原因而完全移`context.sync`出循环之外的解决方案：</span><span class="sxs-lookup"><span data-stu-id="05848-128">However, this is not as good a solution as moving the `context.sync` out of the loop entirely for these reasons:</span></span>
>
> - <span data-ttu-id="05848-129">正如同步批处理作业中的命令已排入队列中一样，批处理作业本身在 Office 中排队，但在队列中的批处理作业不支持超过50个。</span><span class="sxs-lookup"><span data-stu-id="05848-129">Just as the commands in a synchronization batch job are queued, the batch jobs themselves are queued in Office, but Office supports no more than 50 batch jobs in the queue.</span></span> <span data-ttu-id="05848-130">任何其他触发器错误。</span><span class="sxs-lookup"><span data-stu-id="05848-130">Any more triggers errors.</span></span> <span data-ttu-id="05848-131">因此，如果循环中的迭代数超过50个，则会有可能超出队列大小。</span><span class="sxs-lookup"><span data-stu-id="05848-131">So, if there are more than 50 iterations in a loop, there is a chance that the queue size is exceeded.</span></span> <span data-ttu-id="05848-132">迭代次数越多，发生此问题的可能性就越大。</span><span class="sxs-lookup"><span data-stu-id="05848-132">The greater the number of iterations, the greater the chance of this happening.</span></span> 
> - <span data-ttu-id="05848-133">"并发" 并不同时表示。</span><span class="sxs-lookup"><span data-stu-id="05848-133">"Concurrently" does not mean simultaneously.</span></span> <span data-ttu-id="05848-134">执行多个同步操作所需的时间要比执行一个同步操作花费更长时间。</span><span class="sxs-lookup"><span data-stu-id="05848-134">It would still take longer to execute multiple synchronization operations than to execute one.</span></span>
> - <span data-ttu-id="05848-135">并发操作不能保证按照其启动顺序完成。</span><span class="sxs-lookup"><span data-stu-id="05848-135">Concurrent operations are not guaranteed to complete in the same order in which they started.</span></span> <span data-ttu-id="05848-136">在上面的示例中，突出显示 "the" 一词的顺序无关紧要，但在某些情况下，将按顺序处理集合中的项目很重要。</span><span class="sxs-lookup"><span data-stu-id="05848-136">In the preceding example, it doesn't matter what order the  word "the" gets highlighted, but there are scenarios where it's important that the items in the collection be processed in order.</span></span>

## <a name="reading-values-from-the-document-with-the-split-loop-pattern"></a><span data-ttu-id="05848-137">使用拆分循环模式从文档中读取值</span><span class="sxs-lookup"><span data-stu-id="05848-137">Reading values from the document with the split loop pattern</span></span>

<span data-ttu-id="05848-138">当`context.sync`代码必须在处理每个集合项的属性时*读取*这些集合项的属性时，避免 s 在循环中变得更具挑战性。</span><span class="sxs-lookup"><span data-stu-id="05848-138">Avoiding `context.sync`s inside a loop becomes more challenging when the code must *read* a property of the collection items as it processes each one.</span></span> <span data-ttu-id="05848-139">假设您的代码需要对 Word 文档中的所有内容控件进行迭代，并记录与每个控件关联的第一个段落的文本。</span><span class="sxs-lookup"><span data-stu-id="05848-139">Suppose your code needs to iterate all the content controls in a Word document and log the text of the first paragraph associated with each control.</span></span> <span data-ttu-id="05848-140">编程 instincts 可能会引导您在控件上循环，加载每个`text` （第一个）段落的属性，调用`context.sync`使用文档中的文本填充代理段落对象，然后将其记录下来。</span><span class="sxs-lookup"><span data-stu-id="05848-140">Your programming instincts might lead you to loop over the controls, load the `text` property of each (first) paragraph, call `context.sync` to populate the proxy paragraph object with the text from the document, and then log it.</span></span> <span data-ttu-id="05848-141">示例如下。</span><span class="sxs-lookup"><span data-stu-id="05848-141">The following is an example.</span></span>

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

<span data-ttu-id="05848-142">在这种情况下，为了避免`context.sync`在循环中使用，应使用一种模式来调用**拆分循环**模式。</span><span class="sxs-lookup"><span data-stu-id="05848-142">In this scenario, to avoid having a `context.sync` in a loop, you should use a pattern we call the **split loop** pattern.</span></span> <span data-ttu-id="05848-143">我们来看看该模式的具体示例，然后再获取该模式的正式说明。</span><span class="sxs-lookup"><span data-stu-id="05848-143">Let's see a concrete example of the pattern before we get to a formal description of it.</span></span> <span data-ttu-id="05848-144">下面介绍了拆分循环模式如何应用于前面的代码段。</span><span class="sxs-lookup"><span data-stu-id="05848-144">Here's how the split loop pattern can be applied to the preceding code snippet.</span></span> <span data-ttu-id="05848-145">关于此代码，请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="05848-145">Note the following about this code:</span></span>

- <span data-ttu-id="05848-146">现在有两个循环`context.sync` ，它们之间存在，因此不`context.sync`会出现在任何循环中。</span><span class="sxs-lookup"><span data-stu-id="05848-146">There are now two loops and the `context.sync` comes between them, so there's no `context.sync` inside either loop.</span></span>
- <span data-ttu-id="05848-147">第一个循环可循环访问 collection 对象中的项目，并像`text`原始循环那样加载该属性，但第一个循环无法记录段落文本，因为它不再包含`context.sync`用于填充`text` `paragraph`代理对象的属性。</span><span class="sxs-lookup"><span data-stu-id="05848-147">The first loop iterates through the items in the collection object and loads the `text` property just as the original loop did, but the first loop cannot log the paragraph text because it no longer contains a `context.sync` to populate the `text` property of the `paragraph` proxy object.</span></span> <span data-ttu-id="05848-148">而是将`paragraph`对象添加到数组中。</span><span class="sxs-lookup"><span data-stu-id="05848-148">Instead, it adds the `paragraph` object to an array.</span></span>
- <span data-ttu-id="05848-149">第二个循环可循环访问第一个循环创建的数组，并记录`text`每个`paragraph`项目的。</span><span class="sxs-lookup"><span data-stu-id="05848-149">The second loop iterates through the array that was created by the first loop, and logs the `text` of each `paragraph` item.</span></span> <span data-ttu-id="05848-150">这是可行的， `context.sync`这是因为两个循环之间的`text`属性都填充了所有属性。</span><span class="sxs-lookup"><span data-stu-id="05848-150">This is possible because the `context.sync` that came between the two loops populated all the `text` properties.</span></span>

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

<span data-ttu-id="05848-151">上面的示例建议了以下过程，用于打开包含`context.sync`拆分循环模式的循环：</span><span class="sxs-lookup"><span data-stu-id="05848-151">The preceding example suggests the following procedure for turning a loop that contains a `context.sync` into the split loop pattern:</span></span> 

1. <span data-ttu-id="05848-152">将循环替换为两个循环。</span><span class="sxs-lookup"><span data-stu-id="05848-152">Replace the loop with two loops.</span></span>
2. <span data-ttu-id="05848-153">创建第一个循环以对集合进行迭代，并将每个项添加到数组中，同时还加载代码需要读取的项的任何属性。</span><span class="sxs-lookup"><span data-stu-id="05848-153">Create a first loop to iterate over the collection and add each item to an array while also loading any property of the item that your code needs to read.</span></span> 
3. <span data-ttu-id="05848-154">在第一个循环之后， `context.sync`调用以使用任何加载的属性填充代理对象。</span><span class="sxs-lookup"><span data-stu-id="05848-154">Following the first loop, call `context.sync` to populate the proxy objects with any loaded properties.</span></span> 
4. <span data-ttu-id="05848-155">`context.sync`执行第二个循环，以循环访问在第一个循环中创建的数组并读取加载的属性。</span><span class="sxs-lookup"><span data-stu-id="05848-155">Follow the `context.sync` with a second loop to iterate over the array created in the first loop and read the loaded properties.</span></span>

## <a name="processing-objects-in-the-document-with-the-correlated-objects-pattern"></a><span data-ttu-id="05848-156">使用关联对象模式处理文档中的对象</span><span class="sxs-lookup"><span data-stu-id="05848-156">Processing objects in the document with the correlated objects pattern</span></span>

<span data-ttu-id="05848-157">让我们考虑更复杂的情况，即处理集合中的项目需要的数据不在项目本身中。</span><span class="sxs-lookup"><span data-stu-id="05848-157">Let's consider a more complex scenario where processing the items in the collection requires data that isn't in the items themselves.</span></span> <span data-ttu-id="05848-158">方案假设一个 Word 加载项，该加载项对使用某些样本文字的模板创建的文档进行操作。</span><span class="sxs-lookup"><span data-stu-id="05848-158">The scenario envisions a Word add-in that operates on documents created from a template with some boilerplate text.</span></span> <span data-ttu-id="05848-159">分散在文本中的是以下占位符字符串的一个或多个实例： "{协调器}"、"{Deputy}" 和 "{Manager}"。</span><span class="sxs-lookup"><span data-stu-id="05848-159">Scattered in the text are one or more instances of the following placeholder strings: "{Coordinator}", "{Deputy}", and "{Manager}".</span></span> <span data-ttu-id="05848-160">加载项会将每个占位符替换为某人的姓名。</span><span class="sxs-lookup"><span data-stu-id="05848-160">The add-in replaces each placeholder with some person's name.</span></span> <span data-ttu-id="05848-161">外接端的 UI 对本文并不重要。</span><span class="sxs-lookup"><span data-stu-id="05848-161">The UI of the add-in is not important to this article.</span></span> <span data-ttu-id="05848-162">例如，它可能有一个具有三个文本框的任务窗格，每个文本框标有一个占位符。</span><span class="sxs-lookup"><span data-stu-id="05848-162">For example, it could have a task pane with three text boxes, each labeled with one of the placeholders.</span></span> <span data-ttu-id="05848-163">用户在每个文本框中输入一个名称，然后按下一个 "**替换**" 按钮。</span><span class="sxs-lookup"><span data-stu-id="05848-163">The user enters a name in each text box and then presses a **Replace** button.</span></span> <span data-ttu-id="05848-164">该按钮的处理程序将创建一个将名称映射到占位符的数组，然后将每个占位符替换为分配的名称。</span><span class="sxs-lookup"><span data-stu-id="05848-164">The handler for the button creates an array that maps the names to the placeholders, and then replaces each placeholder with the assigned name.</span></span> 

<span data-ttu-id="05848-165">您无需实际生成具有此 UI 的外接程序，即可试用代码。</span><span class="sxs-lookup"><span data-stu-id="05848-165">You don't need to actually produce an add-in with this UI to experiment with the code.</span></span> <span data-ttu-id="05848-166">您可以使用[脚本实验室工具](../overview/explore-with-script-lab.md)对重要代码进行原型。</span><span class="sxs-lookup"><span data-stu-id="05848-166">You can use the [Script Lab tool](../overview/explore-with-script-lab.md) to prototype the important code.</span></span> <span data-ttu-id="05848-167">使用以下赋值语句创建映射数组。</span><span class="sxs-lookup"><span data-stu-id="05848-167">Use the following assignment statement to create the mapping array.</span></span>

```javascript
const jobMapping = [
        { job: "{Coordinator}", person: "Sally" },
        { job: "{Deputy}", person: "Bob" },
        { job: "{Manager}", person: "Kim" }
    ];
```

<span data-ttu-id="05848-168">下面的代码演示在使用`context.sync`内部循环时，如何将每个占位符替换为其分配的名称。</span><span class="sxs-lookup"><span data-stu-id="05848-168">The following code shows how you might replace each placeholder with its assigned name if you used `context.sync` inside loops.</span></span>

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

<span data-ttu-id="05848-169">在上面的代码中，有一个外部循环和一个内层循环。</span><span class="sxs-lookup"><span data-stu-id="05848-169">In the preceding code, there is an outer and an inner loop.</span></span> <span data-ttu-id="05848-170">其中每个都包含`context.sync`一个。</span><span class="sxs-lookup"><span data-stu-id="05848-170">Each of them contains a `context.sync`.</span></span> <span data-ttu-id="05848-171">根据本文中的第一个代码段，您可能会发现在内部循环`context.sync`中，可以在 inner 循环之后直接移动到内部循环中。</span><span class="sxs-lookup"><span data-stu-id="05848-171">Based on the very first code snippet in this article, you probably see that the `context.sync` in the inner loop can simply be moved after the inner loop.</span></span> <span data-ttu-id="05848-172">但在外部循环中，此代码仍`context.sync`会保留（其中两个）。</span><span class="sxs-lookup"><span data-stu-id="05848-172">But that would still leave the code with a `context.sync` (two of them actually) in the outer loop.</span></span> <span data-ttu-id="05848-173">下面的代码演示如何从循环中`context.sync`删除。</span><span class="sxs-lookup"><span data-stu-id="05848-173">The following code shows how you can remove `context.sync` from the loops.</span></span> <span data-ttu-id="05848-174">我们将讨论下面的代码。</span><span class="sxs-lookup"><span data-stu-id="05848-174">We discuss the code below.</span></span>

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

<span data-ttu-id="05848-175">注释代码使用拆分循环模式：</span><span class="sxs-lookup"><span data-stu-id="05848-175">Note the code uses the split loop pattern:</span></span>

- <span data-ttu-id="05848-176">前一示例中的外部循环已拆分为两个。</span><span class="sxs-lookup"><span data-stu-id="05848-176">The outer loop from the preceding example has been split into two.</span></span> <span data-ttu-id="05848-177">（第二个循环有一个内部循环，因为代码在一组工作（或多个占位符）上进行迭代，并且在该集合中对匹配区域进行迭代。</span><span class="sxs-lookup"><span data-stu-id="05848-177">(The second loop has an inner loop, which is expected because the code is iterating over a set of jobs (or placeholders) and within that set it is iterating over the matching ranges.)</span></span>
- <span data-ttu-id="05848-178">每个重大`context.sync`循环之后都有一个，但`context.sync`在任何循环中都不存在。</span><span class="sxs-lookup"><span data-stu-id="05848-178">There is a `context.sync` after each major loop, but no `context.sync` inside any loop.</span></span> 
- <span data-ttu-id="05848-179">第二个主要循环可循环访问在第一个循环中创建的数组。</span><span class="sxs-lookup"><span data-stu-id="05848-179">The second major loop iterates through an array that is created in the first loop.</span></span>

<span data-ttu-id="05848-180">但是，在第一个循环中创建的数组*不*包含一个 Office 对象，因为在[使用拆分循环模式的文档中读取值](#reading-values-from-the-document-with-the-split-loop-pattern)的节中的第一个循环。</span><span class="sxs-lookup"><span data-stu-id="05848-180">But the array created in the first loop does *not* contain only an Office object as the first loop did in the section [Reading values from the document with the split loop pattern](#reading-values-from-the-document-with-the-split-loop-pattern).</span></span> <span data-ttu-id="05848-181">这是因为处理 Word Range 对象所需的一些信息不在 Range 对象本身中，而是来自于`jobMapping`数组。</span><span class="sxs-lookup"><span data-stu-id="05848-181">This is because some of the information needed to process the Word Range objects is not in the Range objects themselves but instead comes from the `jobMapping` array.</span></span> 

<span data-ttu-id="05848-182">因此，在第一个循环中创建的数组中的对象是具有两个属性的自定义对象。</span><span class="sxs-lookup"><span data-stu-id="05848-182">So, the objects in the array created in the first loop are custom objects that have two properties.</span></span> <span data-ttu-id="05848-183">第一个是与特定职务（即占位符字符串）匹配的单词范围的数组，第二个是提供分配到该作业的人员姓名的字符串。</span><span class="sxs-lookup"><span data-stu-id="05848-183">The first is an array of Word Ranges that match a specific job title (that is, a placeholder string) and the second is a string that provides the name of the person assigned to the job.</span></span> <span data-ttu-id="05848-184">这使得最终循环易于编写和易于阅读，因为处理给定区域所需的全部信息都包含在包含该范围的同一自定义对象中。</span><span class="sxs-lookup"><span data-stu-id="05848-184">This makes the final loop easy to write and easy to read because all of the information needed to process a given range is contained in the same custom object that contains the range.</span></span> <span data-ttu-id="05848-185">应替换_ **correlatedObject**[j]_ 的名称是同一对象的另一个属性： _ **correlatedObject**_。</span><span class="sxs-lookup"><span data-stu-id="05848-185">The name that should replace _**correlatedObject**.rangesMatchingJob.items[j]_ is the other property of the same object: _**correlatedObject**.personAssignedToJob_.</span></span> 

<span data-ttu-id="05848-186">我们称之为 "**关联对象**" 模式的拆分循环模式的这一变体。</span><span class="sxs-lookup"><span data-stu-id="05848-186">We call this variation of the split loop pattern the **correlated objects** pattern.</span></span> <span data-ttu-id="05848-187">一般来讲，第一条循环创建自定义对象的数组。</span><span class="sxs-lookup"><span data-stu-id="05848-187">The general idea is that the first loop creates an array of custom objects.</span></span> <span data-ttu-id="05848-188">每个对象都有一个属性，其值是 Office collection 对象（或此类项目的数组）中的项目之一。</span><span class="sxs-lookup"><span data-stu-id="05848-188">Each object has a property whose value is one of the items in an Office collection object (or an array of such items).</span></span> <span data-ttu-id="05848-189">自定义对象具有其他属性，每个属性都提供处理最终循环中的 Office 对象所需的信息。</span><span class="sxs-lookup"><span data-stu-id="05848-189">The custom object has other properties, each of which provides information needed to process the Office objects in the final loop.</span></span> <span data-ttu-id="05848-190">请参阅[这些模式的其他示例](#other-examples-of-these-patterns)部分，以获取自定义关联对象具有两个以上属性的示例的链接。</span><span class="sxs-lookup"><span data-stu-id="05848-190">See the section [Other examples of these patterns](#other-examples-of-these-patterns) for a link to an example where the custom correlating object has more than two properties.</span></span>

<span data-ttu-id="05848-191">另一个需要注意的一点是，有时需要多个循环来创建自定义关联对象的数组。</span><span class="sxs-lookup"><span data-stu-id="05848-191">One further caveat: sometimes it takes more than one loop just to create the array of custom correlating objects.</span></span> <span data-ttu-id="05848-192">如果您需要只读取一个 Office 集合对象的每个成员的属性来收集将用于处理另一个集合对象的信息，则会发生这种情况。</span><span class="sxs-lookup"><span data-stu-id="05848-192">This can happen if you need to read a property of each member of one Office collection object just to gather information that will be used to process another collection object.</span></span> <span data-ttu-id="05848-193">（例如，您的代码需要读取 Excel 表中所有列的标题，因为您的外接程序将根据该列的标题对某些列的单元格应用数字格式。）但您始终可以在循环`context.sync`之间，而不是在循环之间保持 s。</span><span class="sxs-lookup"><span data-stu-id="05848-193">(For example, your code needs to read the titles of all the columns in an Excel table because your add-in is going to apply a number format to the cells of some columns based on that column's title.) But you can always keep the `context.sync`s between the loops, rather than in a loop.</span></span> <span data-ttu-id="05848-194">有关示例，请参阅[这些模式的其他示例](#other-examples-of-these-patterns)一节。</span><span class="sxs-lookup"><span data-stu-id="05848-194">See the section [Other examples of these patterns](#other-examples-of-these-patterns) for an example.</span></span>

## <a name="other-examples-of-these-patterns"></a><span data-ttu-id="05848-195">这些模式的其他示例</span><span class="sxs-lookup"><span data-stu-id="05848-195">Other examples of these patterns</span></span>

- <span data-ttu-id="05848-196">有关使用`Array.forEach`循环的 Excel 的非常简单的示例，请参阅此堆栈溢出问题的接受答案：[是否可以对多个上下文进行排队。在 context 之前进行加载？](https://stackoverflow.com/questions/44459604/is-it-possible-to-queue-more-than-one-context-load-before-context-sync)</span><span class="sxs-lookup"><span data-stu-id="05848-196">For a very simple example for Excel that uses `Array.forEach` loops, see the accepted answer to this Stack Overflow question: [Is it possible to queue more than one context.load before context.sync?](https://stackoverflow.com/questions/44459604/is-it-possible-to-queue-more-than-one-context-load-before-context-sync)</span></span>
- <span data-ttu-id="05848-197">有关使用`Array.forEach`循环但不使用`async` / `await`语法的 Word 的简单示例，请参阅 "接受的对此堆栈溢出问题的答案：使用[Office JavaScript API 循环访问包含内容控件的所有段落](https://stackoverflow.com/questions/58422113/iterating-over-all-paragraphs-with-content-controls-with-office-javascript-api)"。</span><span class="sxs-lookup"><span data-stu-id="05848-197">For a simple example for Word that uses `Array.forEach` loops and doesn't use `async`/`await` syntax, see the accepted answer to this Stack Overflow question: [Iterating over all paragraphs with content controls with Office JavaScript API](https://stackoverflow.com/questions/58422113/iterating-over-all-paragraphs-with-content-controls-with-office-javascript-api).</span></span>
- <span data-ttu-id="05848-198">有关使用 TypeScript 编写的 Word 的示例，请参阅示例[Word 外接程序 Angular2 样式检查器](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker)，尤其是文件 " [document](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker/blob/master/app/services/word-document/word.document.service.ts)"。</span><span class="sxs-lookup"><span data-stu-id="05848-198">For an example for Word that is written in TypeScript, see the sample [Word Add-in Angular2 Style Checker](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker), especially the file [word.document.service.ts](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker/blob/master/app/services/word-document/word.document.service.ts).</span></span> <span data-ttu-id="05848-199">它混合了`for`和`Array.forEach`循环。</span><span class="sxs-lookup"><span data-stu-id="05848-199">It has a mixture of `for` and `Array.forEach` loops.</span></span>
- <span data-ttu-id="05848-200">对于高级 Word 示例，请将[此 gist](https://gist.github.com/9c5a803e52480ec7f00bb3224292e0ab)导入[脚本实验室工具](../overview/explore-with-script-lab.md)。</span><span class="sxs-lookup"><span data-stu-id="05848-200">For an advanced Word sample, import [this gist](https://gist.github.com/9c5a803e52480ec7f00bb3224292e0ab) into the [Script Lab tool](../overview/explore-with-script-lab.md).</span></span> <span data-ttu-id="05848-201">有关使用 gist 的上下文，请参阅在[替换文本后，不同步](https://stackoverflow.com/questions/48227941/document-not-in-sync-after-replace-text)"堆栈溢出问题" 文档中的 "已接受的答案"。</span><span class="sxs-lookup"><span data-stu-id="05848-201">For context in using the gist, see the accepted answer to the Stack Overflow question [Document not in sync after replace text](https://stackoverflow.com/questions/48227941/document-not-in-sync-after-replace-text).</span></span> <span data-ttu-id="05848-202">本示例创建一个具有三个属性的自定义关联对象类型。</span><span class="sxs-lookup"><span data-stu-id="05848-202">This sample creates a custom correlating object type that has three properties.</span></span> <span data-ttu-id="05848-203">它总共使用三个循环来构造相关对象的数组，以及执行最后处理的两个更多循环。</span><span class="sxs-lookup"><span data-stu-id="05848-203">It uses a total of three loops to construct the array of correlated objects, and two more loops to do the final processing.</span></span> <span data-ttu-id="05848-204">混合了`for`和`Array.forEach`循环。</span><span class="sxs-lookup"><span data-stu-id="05848-204">There are a mixture of `for` and `Array.forEach` loops.</span></span>
- <span data-ttu-id="05848-205">尽管不是严格的拆分循环或相关对象模式的示例，但还有一个演示如何将一组单元格的值转换为只使用一个的其他货币的高级 Excel `context.sync`示例。</span><span class="sxs-lookup"><span data-stu-id="05848-205">Although not strictly an example of the split loop or correlated objects patterns, there is an advanced Excel sample that shows how to convert a set of cell values to other currencies with just a single `context.sync`.</span></span> <span data-ttu-id="05848-206">若要尝试，请打开[脚本实验室工具](../overview/explore-with-script-lab.md)并导航到**货币转换器**示例。</span><span class="sxs-lookup"><span data-stu-id="05848-206">To try it, open the [Script Lab tool](../overview/explore-with-script-lab.md) and navigate to the **Currency Converter** sample.</span></span> 

## <a name="when-should-you-not-use-the-patterns-in-this-article"></a><span data-ttu-id="05848-207">何时*应使用本文*中的模式？</span><span class="sxs-lookup"><span data-stu-id="05848-207">When should you *not* use the patterns in this article?</span></span>

<span data-ttu-id="05848-208">Excel 在给定的`context.sync`调用中无法读取超过 5 MB 的数据。</span><span class="sxs-lookup"><span data-stu-id="05848-208">Excel cannot read more than 5 MB of data in a given call of `context.sync`.</span></span> <span data-ttu-id="05848-209">如果超过此限制，则会引发错误。</span><span class="sxs-lookup"><span data-stu-id="05848-209">If this limit is exceeded, an error is thrown.</span></span> <span data-ttu-id="05848-210">（有关详细信息，请参阅[Excel data transfer 限制](../develop/common-coding-issues.md#excel-data-transfer-limits)。）很少需要此限制，但如果有机会在外接程序中执行此操作，则代码*不*应在单个循环中加载所有数据，并在循环中使用 a `context.sync`。</span><span class="sxs-lookup"><span data-stu-id="05848-210">(For more information, see [Excel data transfer limits](../develop/common-coding-issues.md#excel-data-transfer-limits).) It is very rare that this limit is approached, but if there's a chance that this will happen with your add-in, then your code should *not* load all the data in a single loop and follow the loop with a `context.sync`.</span></span> <span data-ttu-id="05848-211">但您仍应避免`context.sync`在集合对象上循环的每个迭代。</span><span class="sxs-lookup"><span data-stu-id="05848-211">But you still should avoid having a `context.sync` in every iteration of a loop over a collection object.</span></span> <span data-ttu-id="05848-212">相反，在集合中定义项的子集，并依次对每个子集进行循环，并`context.sync`在循环之间进行循环。</span><span class="sxs-lookup"><span data-stu-id="05848-212">Instead, define subsets of the items in the collection and loop over each subset in turn, with a `context.sync` between the loops.</span></span> <span data-ttu-id="05848-213">您可以使用外部循环对此进行构造，该循环可对子集进行`context.sync`迭代，并在每个外部迭代中包含。</span><span class="sxs-lookup"><span data-stu-id="05848-213">You could structure this with an outer loop that iterates over the subsets and contains the `context.sync` in each of these outer iterations.</span></span>
