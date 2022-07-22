---
title: 避免在循环中使用 context.sync
description: 了解如何使用拆分循环和关联对象模式避免在循环中调用 context.sync。
ms.date: 07/18/2022
ms.localizationpriority: medium
ms.openlocfilehash: 6b0239e05a597949160afbb2604143f3d6626462
ms.sourcegitcommit: b6a3815a1ad17f3522ca35247a3fd5d7105e174e
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/22/2022
ms.locfileid: "66958697"
---
# <a name="avoid-using-the-contextsync-method-in-loops"></a>避免在循环中使用 context.sync

> [!NOTE]
> 本文假定你已超出使用批处理系统与 Office 文档交互的四个特定于应用程序的 Excel、Word、OneNote 和 Visio&mdash;的 Office JavaScript API&mdash;之一的开头阶段。 特别是，你应该知道调用是做什么的 `context.sync` ，你应该知道什么是集合对象。 如果未处于该阶段，请首先了解该文章中“特定于应用程序”的 [Office JavaScript API](../develop/understanding-the-javascript-api-for-office.md) 和链接到的文档。

对于 Office 外接程序中使用 Excel、Word、OneNote 和 Visio)  (的特定于应用程序的 API 模型之一的编程方案，代码需要从集合对象的每个成员读取、写入或处理某些属性。 例如，需要获取特定表列或 Word 加载项中每个单元格的值的 Excel 加载项，该加载项需要突出显示文档中字符串的每个实例。 需要循环访问集合对象属性中 `items` 的成员;但是，出于性能原因，需要避免在循环的每次迭代中调用 `context.sync` 。 每次调用 `context.sync` 都是从加载项到 Office 文档的往返。 重复往返会损害性能，尤其是在加载项在Office web 版中运行时，因为往返会通过 Internet 进行。

> [!NOTE]
> 本文中的所有示例都使用 `for` 循环，但所述的做法适用于可循环访问数组的任何循环语句，包括：
>
> - `for`
> - `for of`
> - `while`
> - `do while`
>
> 它们还适用于向其传递函数并应用于数组中项的任何数组方法，包括：
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

## <a name="writing-to-the-document"></a>写入文档

在最简单的情况下，你只写入集合对象的成员，而不是读取其属性。 例如，以下代码以黄色突出显示 Word 文档中每个“the”实例。

> [!NOTE]
> 通常，在应用程序`run`函数 (（如`Word.run``Excel.run`等）的结束“}”字符) 之前，有一个决赛`context.sync`是一个好的做法。 这是因为函 `run` 数发出隐藏调用 `context.sync` 是它做的最后一件事，前提是，只有当有排队的命令尚未同步时。 此调用隐藏的事实可能会令人困惑，因此我们通常建议添加显式 `context.sync`调用。 然而，鉴于这篇文章是关于尽量减少呼吁 `context.sync`，它实际上是更令人困惑，添加一个完全不必要的最终 `context.sync`。 因此，在本文中，当末尾 `run`没有未同步的命令时，我们将其排除。

```javascript
await Word.run(async function (context) {
  let startTime, endTime;
  const docBody = context.document.body;

  // search() returns an array of Ranges.
  const searchResults = docBody.search('the', { matchWholeWord: true });
  searchResults.load('font');
  await context.sync();

  // Record the system time.
  startTime = performance.now();

  for (let i = 0; i < searchResults.items.length; i++) {
    searchResults.items[i].font.highlightColor = '#FFFF00';

    await context.sync(); // SYNCHRONIZE IN EACH ITERATION
  }
  
  // await context.sync(); // SYNCHRONIZE AFTER THE LOOP

  // Record the system time again then calculate how long the operation took.
  endTime = performance.now();
  console.log("The operation took: " + (endTime - startTime) + " milliseconds.");
})
```

前面的代码花了 1 整秒才在具有 200 个 Windows Word 实例的文档中完成。 但是，当循环内的 `await context.sync();` 行被注释掉，并且循环刚刚取消注释后，该操作只用了第 1/10 秒。 在以 Edge 作为浏览器) 的 Web Word (中，循环内的同步用了整整 3 秒，循环后的同步仅用了 6/10 秒，大约快了 5 倍。 在包含 2000 个“the”实例的文档中，在 Word 网页版中 () 80 秒，循环内同步，循环后同步仅 4 秒，大约快 20 倍。

> [!NOTE]
> 值得一问的是，如果同步并发运行，同步在循环内部版本的执行速度是否会更快，这可以通过从前面`context.sync()`删除`await`关键字来完成。 这将导致运行时启动同步，然后立即启动循环的下一次迭代，而无需等待同步完成。 但是，由于这些原因，这并不像完全从循环中移 `context.sync` 出那样好。
>
> - 正如同步批处理作业中的命令排队一样，批处理作业本身也在 Office 中排队，但 Office 在队列中支持不超过 50 个批处理作业。 任何其他触发器错误。 因此，如果循环中有超过 50 次迭代，则可能会超过队列大小。 迭代次数越多，发生这种情况的几率就越大。
> - “并发”并不意味着同时。 执行多个同步操作所需的时间仍比执行一个同步操作要长。
> - 并发操作不能保证按照启动顺序完成。 在前面的示例中，突出显示“the”一词的顺序并不重要，但在某些情况下，必须按顺序处理集合中的项。

## <a name="read-values-from-the-document-with-the-split-loop-pattern"></a>使用拆分循环模式从文档读取值

当代码在处理每个集合项时必须 *读取* 集合项的属性时，避免`context.sync`循环中的 s 将变得更加具有挑战性。 假设代码需要循环访问 Word 文档中的所有内容控件，并记录与每个控件关联的第一段的文本。 编程本能可能会导致你循环访问控件，先加载 `text` 每个 () 段落的属性，调用 `context.sync` 用文档中的文本填充代理段落对象，然后记录它。 示例如下。

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

在此方案中，为了避免 `context.sync` 出现循环，应使用我们调用 **拆分循环** 模式的模式。 让我们先查看模式的具体示例，然后再对它进行正式说明。 下面介绍拆分循环模式如何应用于前面的代码片段。 对于此代码，请注意以下事项。

- 现在有两个循环， `context.sync` 它们之间有两个循环，所以这两个循环中都没有 `context.sync` 。
- 第一个循环循环访问集合对象中的项，并像原始循环一样加载 `text` 该属性，但第一个 `context.sync` 循环无法记录段落文本，因为它不再包含用于填充 `text` 代理对象的 `paragraph` 属性。 而是将对象添加 `paragraph` 到数组。
- 第二个循环循环遍历由第一个循环创建的数组，并记录 `text` 每个 `paragraph` 项。 这是可能的， `context.sync` 因为这两个循环之间的循环填充了 `text` 所有属性。

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

前面的示例建议将包含 `context.sync` 拆分循环模式的循环转换为以下过程。

1. 将循环替换为两个循环。
2. 创建第一个循环来循环访问集合，并将每个项添加到数组，同时加载代码需要读取的项的任何属性。
3. 在第一个循环之后，调用 `context.sync` 以使用任何加载的属性填充代理对象。
4. `context.sync`使用第二个循环访问在第一个循环中创建的数组并读取已加载的属性。

## <a name="process-objects-in-the-document-with-the-correlated-objects-pattern"></a>使用相关对象模式处理文档中的对象

让我们考虑一个更复杂的方案：处理集合中的项需要项本身不包含的数据。 该方案设想了一个 Word 加载项，该加载项对从具有一些样板文本的模板创建的文档进行操作。 文本中分散的是以下占位符字符串的一个或多个实例：“{Coordinator}”、“{Deputy}”和“{Manager}”。 外接程序将每个占位符替换为某人的姓名。 外接程序的 UI 对本文并不重要。 例如，它可以有一个任务窗格，其中包含三个文本框，每个文本框都带有一个占位符标记。 用户在每个文本框中输入一个名称，然后按“ **替换”** 按钮。 按钮的处理程序创建一个数组，该数组将名称映射到占位符，然后将每个占位符替换为分配的名称。

无需使用此 UI 实际生成加载项即可试验代码。 可以使用[Script Lab工具](../overview/explore-with-script-lab.md)对重要代码进行原型处理。 使用以下赋值语句创建映射数组。

```javascript
const jobMapping = [
        { job: "{Coordinator}", person: "Sally" },
        { job: "{Deputy}", person: "Bob" },
        { job: "{Manager}", person: "Kim" }
    ];
```

以下代码演示了如果在循环中使用 `context.sync` ，如何将每个占位符替换为其分配的名称。

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

在前面的代码中，有一个外部循环和一个内部循环。 其中每个都包含一个 `context.sync`。 根据本文中的第一个代码片段，你可能会看到 `context.sync` 内部循环中的代码片段只需在内部循环之后移动即可。 但是，这仍然会使代码保留一个 (其中两个 `context.sync` 实际上) 在外部循环中。 以下代码演示如何从循环中删除 `context.sync` 。 我们将讨论以下代码。

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

请注意，代码使用拆分循环模式。

- 前面示例中的外部循环已拆分为两个。  (第二个循环有一个内部循环，这是预期的，因为代码正在循环访问一组作业 (或占位符) 并在该设置中循环访问匹配的范围。) 
- 每个主要循环后都有一个 `context.sync` ，但没有任何 `context.sync` 循环。
- 第二个主要循环循环遍历在第一个循环中创建的数组。

但是，在第一个循环中创建的数组 *并不* 只包含 Office 对象，就像第一个循环使用 [拆分循环模式读取文档中的值](#read-values-from-the-document-with-the-split-loop-pattern)一样。 这是因为处理 Word Range 对象所需的一些信息不在 Range 对象本身中，而是来自数 `jobMapping` 组。

因此，在第一个循环中创建的数组中的对象是具有两个属性的自定义对象。 第一个是与特定职务 (匹配的字范围数组，即占位符字符串) ，第二个是提供分配给作业的人员的姓名的字符串。 这使最终循环易于编写和易于阅读，因为处理给定范围所需的所有信息都包含在包含该范围的同一自定义对象中。 应替换 _correlatedObject.rangesMatchingJob.items[j]_ 的名称是同一对象的另一个属性：_**correlatedObject.personAssignedToJob**_。

我们将拆分循环模式的此变体称为 **相关对象** 模式。 一般的想法是，第一个循环创建自定义对象的数组。 每个对象都有一个属性，其值是 Office 集合对象中的项之一 (或此类项的数组) 。 自定义对象具有其他属性，每个属性都提供在最终循环中处理 Office 对象所需的信息。 有关自定义关联对象具有两个以上属性的示例的链接，请参阅 [这些模式的其他示例](#other-examples-of-these-patterns) 部分。

还有一个注意事项：有时只需创建自定义关联对象的数组，就需要多个循环。 如果需要只读取一个 Office 集合对象的每个成员的属性来收集用于处理另一个集合对象的信息，则可能会发生这种情况。  (例如，代码需要读取 Excel 表中所有列的标题，因为加载项将基于该列的标题将数字格式应用于某些列的单元格。) 但始终可以在循环之间保留该值，而不是在循环中保留 `context.sync`。 有关示例，请参阅 [这些模式的其他示例](#other-examples-of-these-patterns) 部分。

## <a name="other-examples-of-these-patterns"></a>这些模式的其他示例

- 有关使用 `Array.forEach` 循环的 Excel 的一个非常简单的示例，请参阅此 Stack Overflow 问题的接受答案： [是否可以在 context.sync 之前对多个 context.load 进行排队？](https://stackoverflow.com/questions/44459604/is-it-possible-to-queue-more-than-one-context-load-before-context-sync)
- 有关使用`Array.forEach`循环且不使用`await``async`/语法的 Word 的简单示例，请参阅此 Stack Overflow 问题的接受答案：[使用 Office JavaScript API 循环访问包含内容控件的所有段落](https://stackoverflow.com/questions/58422113/iterating-over-all-paragraphs-with-content-controls-with-office-javascript-api)。
- 有关以 TypeScript 编写的 Word 的示例，请参阅示例 [Word 加载项 Angular2 样式检查器](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker)，尤其是文件 [word.document.service.ts](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker/blob/master/app/services/word-document/word.document.service.ts)。 它具有混合 `for` 和 `Array.forEach` 循环。
- 对于高级 Word 示例，请将[此 gist](https://gist.github.com/9c5a803e52480ec7f00bb3224292e0ab) 导入[Script Lab工具](../overview/explore-with-script-lab.md)。 有关使用 gist 的上下文，请参阅 [替换文本后未同步](https://stackoverflow.com/questions/48227941/document-not-in-sync-after-replace-text)的 Stack Overflow 问题文档的接受答案。 此示例创建具有三个属性的自定义关联对象类型。 它总共使用三个循环来构造相关对象的数组，另外使用两个循环进行最终处理。 有混合和`for``Array.forEach`循环。
- 虽然不是拆分循环或相关对象模式的严格示例，但有一个高级 Excel 示例，演示如何将一组单元格值转换为单个 `context.sync`单元格值的其他货币。 若要尝试，请打开 [Script Lab工具](../overview/explore-with-script-lab.md)并导航到 **“货币转换器”** 示例。

## <a name="when-should-you-not-use-the-patterns-in-this-article"></a>何时 *不应* 使用本文中的模式？

Excel 在给定调用 `context.sync`中读取的数据不能超过 5 MB。 如果超出此限制，则会引发错误。  (有关详细信息，请参阅 Office 外接程序的 [资源限制和性能优化](resource-limits-and-performance-optimization.md#excel-add-ins) 的“Excel 加载项”部分。) 此限制很少会被接近，但如果加载项可能会发生这种情况，则代码 *不应* 在单个循环中加载所有数据，并按照循环进行操作 `context.sync`。 但是，仍应避免 `context.sync` 在集合对象上进行循环的每次迭代。 相反，定义集合中项的子集，并依次循环访问每个子集，并使用 `context.sync` 循环之间。 可以使用循环访问子集并包含 `context.sync` 每个外部迭代中的外部循环来构建此项。
