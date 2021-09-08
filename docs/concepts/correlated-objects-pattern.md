---
title: 避免在循环中使用 context.sync
description: 了解如何使用拆分循环和相关对象模式避免在循环中调用 context.sync。
ms.date: 07/08/2021
localization_priority: Normal
ms.openlocfilehash: 85230378f40be06c7f3385f5dde88ecaba503cb5
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/08/2021
ms.locfileid: "58938125"
---
# <a name="avoid-using-the-contextsync-method-in-loops"></a>避免在循环中使用 context.sync

> [!NOTE]
> 本文假定你已超出使用批处理系统与 Office 文档交互的 &mdash; Excel、Word、OneNote 和 Visio 的四个特定于应用程序的 Office JavaScript API 中的至少一个的开始阶段。 &mdash; 特别是，你应了解调用功能， `context.sync` 并且应了解集合对象是什么。 如果你未处于该阶段，请从了解[javaScript API](../develop/understanding-the-javascript-api-for-office.md) Office以及该文章中"特定于应用程序"下的链接文档开始。

对于 Office 外接程序中对 Excel、Word、OneNote 和 Visio) 使用应用程序特定的 API 模型 (的一种编程方案，代码需要读取、写入或处理集合对象每个成员中的一些属性。 例如，Excel需要获取特定表格列内每个单元格的值的加载项，或需要突出显示文档中每个字符串实例的 Word 加载项。 您需要循环访问集合对象的 属性中的成员;但出于性能原因，您需要避免在循环的每次迭代中 `items` `context.sync` 调用。 每次调用 `context.sync` 都是从加载项到文档Office行程。 重复的往返会损害性能，尤其是在外接程序运行在 Office web 版，因为往返行程通过 Internet。

> [!NOTE]
> 本文中所有示例都使用循环，但所介绍的做法适用于可以循环访问数组的任何循环语句， `for` 其中包括：
>
> - `for`
> - `for of`
> - `while`
> - `do while`
> 
> 它们还适用于函数传递给并应用于数组中的项目的任何数组方法，包括：
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

在最简单的情况下，您只写入集合对象的成员，而不是读取它们的属性。 例如，以下代码在 Word 文档中以黄色突出显示每个"the"实例。

> [!NOTE]
> 通常，在应用程序方法的结束"}"字符（如 、等）之前 (一个最终 `context.sync` `run` `Excel.run` `Word.run` ) 。 这是因为当（并且仅在存在尚未同步的已排队命令）时，该方法将执行最后一项操作进行隐藏 `run` `context.sync` 调用。 隐藏此调用这一事实可能会令人困惑，因此，我们通常建议添加显式 `context.sync` 。 但是，鉴于本文将调用最小化，添加一个完全不必要的最终 ，实际上会更 `context.sync` 令人困惑 `context.sync` 。 因此，在本文中，当 末尾没有未同步的命令时，我们会不进行介绍 `run` 。

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

前面的代码在文档中使用 200 个实例的 Word on Windows 完成。 但是，当在取消注释循环后将循环中的行注释掉且同一行时，该操作只需 `await context.sync();` 1/10 秒。 在Word web 版 (Edge 作为浏览器) 时，循环内同步需要 3 秒钟，在循环后同步只需 6/10 秒，大约快五倍。 在具有 2000 个实例的"the"的文档中，在 (Word web 版) 80 秒内通过循环内同步，在循环后仅同步 4 秒，大约快 20 倍。

> [!NOTE]
> 值得一提的是，如果同步同时运行（只需从 的前面删除 关键字，同步执行速度是否加快，循环内部同步版本的执行速度是否 `await` 更快 `context.sync()` ）。 这会使运行时启动同步，然后立即启动循环的下一次迭代，而无需等待同步完成。 但是，由于这些原因，这不是一个比完全退出循环 `context.sync` 好的解决方案。
>
> - 就像同步批处理作业中的命令排入队列一样，批处理作业本身在 Office 中排队，Office在队列中支持不超过 50 个批处理作业。 其他任何操作都会引发错误。 因此，如果循环中迭代次数超过 50 次，则有可能超出队列大小。 迭代次数越大，发生迭代的可能性越大。 
> - "并发"并不意味着同时进行。 执行多个同步操作比执行一个同步操作要长。
> - 无法保证并发操作按启动顺序完成。 在上一示例中，"the"一词的突出显示顺序无关紧要，但在某些情况下，必须按顺序处理集合中的项。

## <a name="read-values-from-the-document-with-the-split-loop-pattern"></a>使用拆分循环模式读取文档中的值

当代码在处理每个集合项时必须读取集合项的属性时，避免在循环内运行 `context.sync` 将更具挑战性。  假设您的代码需要对 Word 文档中的所有内容控件进行重新访问，并记录与每个控件关联的第一段的文本。 您的编程尝试可能会导致您循环访问控件、加载每个 (第一个) 段落的属性、调用以使用文档中的文本填充代理段落对象，然后记录它 `text` `context.sync` 。 示例如下。

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

在此方案中，为了避免 在循环中出现 ，你应该使用我们调用拆分循环 `context.sync` **模式** 的模式。 在获得模式的正式说明之前，让我们看一个具体模式示例。 下面将说明拆分循环模式如何应用于前面的代码段。 对于此代码，请注意以下事项。

- 现在存在两个循环 `context.sync` ，两个循环之间出现，因此两个循环中 `context.sync` 都不存在。
- 第一个循环循环访问集合对象中的项目并加载属性，就像原始循环一样，但第一个循环无法记录段落文本，因为它不再包含 用于填充代理对象的属性的 `text` 。 `context.sync` `text` `paragraph` 相反，它会 `paragraph` 将对象添加到数组中。
- 第二个循环循环访问由第一个循环创建的数组，并记录 `text` 每个项目的 `paragraph` 。 这是可能的，因为 两个循环之间的 填充 `context.sync` 了所有 `text` 属性。

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

前面的示例建议以下过程将包含 的 循环 `context.sync` 转换为拆分循环模式。

1. 将循环替换为两个循环。
2. 创建第一个循环来循环访问集合，将每个项目添加到数组中，同时加载代码需要读取的项目的任何属性。
3. 第一个循环之后，调用 `context.sync` 以使用任何加载的属性填充代理对象。
4. 按照 第二个循环操作，循环访问第一个循环中创建的数组并 `context.sync` 读取加载的属性。

## <a name="process-objects-in-the-document-with-the-correlated-objects-pattern"></a>使用相关对象模式处理文档中的对象

让我们考虑一个更复杂的方案，其中处理集合中的项需要不在项目本身内的数据。 方案设想一个 Word 外接程序，该外接程序对从具有一些样本文本的模板创建的文档进行操作。 分散在文本中是以下占位符字符串的一个或多个实例："{Coordinator}"、"{Coordinatory}"和"{Manager}"。 外接程序将每个占位符替换为某人的姓名。 对于本文，外接程序的 UI 不十分重要。 例如，它可以有一个包含三个文本框的任务窗格，每个文本框都标记有一个占位符。 用户在每个文本框中输入一个名称，然后 **按"替换** "按钮。 按钮的处理程序创建一个数组，该数组将名称映射到占位符，然后用分配的名称替换每个占位符。

你无需实际通过此 UI 生成外接程序来试验代码。 可以使用 Script Lab[工具](../overview/explore-with-script-lab.md)建立重要代码的原型。 使用以下赋值语句创建映射数组。

```javascript
const jobMapping = [
        { job: "{Coordinator}", person: "Sally" },
        { job: "{Deputy}", person: "Bob" },
        { job: "{Manager}", person: "Kim" }
    ];
```

以下代码显示如何在使用内部循环时，将每个占位符替换为其 `context.sync` 分配的名称。

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

在上一个代码中，有一个外部和一个内部循环。 其中每个都包含 `context.sync` 一个 。 根据本文中第一个代码段，你可能会看到，内循环中的 可以仅移到内部 `context.sync` 循环之后。 但是，这仍将代码保留为 `context.sync` (，其中两个) 在外部循环中。 以下代码演示如何从 `context.sync` 循环中删除。 我们将讨论以下代码。

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

- 上例中的外部循环已拆分为两个。  (第二个循环有一个内部循环，这是预期的，因为代码将循环一组作业 (或占位符) 并且在此循环内它循环测试匹配的范围。) 
- 每个主 `context.sync` 循环后都有 一个 ，但在任何 `context.sync` 循环内没有。
- 第二个主要循环循环访问第一个循环中创建的数组。

但是，第一个循环中创建的数组并不只包含一个 Office 对象，正如第一个循环在使用拆分循环模式读取文档中的值部分中[所执行。](#read-values-from-the-document-with-the-split-loop-pattern) 这是因为处理 Word Range 对象所需的某些信息不在 Range 对象本身中，而是来自 `jobMapping` 数组。

因此，第一个循环中创建的数组中的对象是具有两个属性的自定义对象。 第一种是匹配特定职务 (（即占位符字符串) ）的 Word 范围数组，第二个数组是一个字符串，该字符串提供分配给该工作的人的姓名。 这使得最后一个循环易于编写且易于阅读，因为处理给定区域所需的全部信息都包含在包含该范围的同一自定义对象中。 应替换 _**correlatedObject**.rangesMatchingJob.items[j]_ 的名称是同一对象的另一个属性 _**：correlatedObject**.personAssignedToJob_。

我们将此变体称为拆分循环模式 **的相关对象** 模式。 一般概念是，第一个循环创建一个自定义对象数组。 每个对象都有一个属性值，该属性是 Office 集合对象 (或此类项目数组中的) 。 自定义对象具有其他属性，每个属性都提供在最终循环中处理Office对象所需的信息。 有关指向 [自定义关联](#other-examples-of-these-patterns) 对象具有两个以上属性的示例的链接，请参阅这些模式的其他示例一节。

另一个警告：有时，仅创建自定义关联对象的数组需要多个循环。 如果需要读取一个集合对象中每个成员的属性，Office收集将用于处理另一个集合对象的信息，则可能会发生这种情况。  (例如，您的代码需要读取 Excel 表中所有列的标题，因为您的外接程序将基于该列的标题将数字格式应用于某些列的单元格。) 但您可以始终在循环之间保留 ，而不是在循环中。 `context.sync` 有关示例 [，请参阅这些模式的其他](#other-examples-of-these-patterns) 示例部分。

## <a name="other-examples-of-these-patterns"></a>这些模式的其他示例

- 有关使用循环的Excel示例，请参阅此 Stack Overflow 问题的接受答案：在 context.sync 之前，是否可能将多个 `Array.forEach` [context.load](https://stackoverflow.com/questions/44459604/is-it-possible-to-queue-more-than-one-context-load-before-context-sync)排入队列？
- 有关使用循环且不使用语法的 Word 的简单示例，请参阅此 Stack Overflow 问题的接受答案：使用 Office JavaScript API 遍历包含内容 `Array.forEach` `async` / `await` [控件的所有](https://stackoverflow.com/questions/58422113/iterating-over-all-paragraphs-with-content-controls-with-office-javascript-api)段落。
- 有关使用 TypeScript 编写的 Word 示例，请参阅示例 [Word 外接程序 Angular2](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker)样式检查器，尤其是文件 [word.document.service.ts](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker/blob/master/app/services/word-document/word.document.service.ts)。 它混合了 `for` 和 `Array.forEach` 循环。
- 对于高级 Word 示例，将[此 gist](https://gist.github.com/9c5a803e52480ec7f00bb3224292e0ab)导入[Script Lab 工具。](../overview/explore-with-script-lab.md) 有关使用 gist 的上下文，请参阅 Stack Overflow 问题的接受答案替换文本 [后文档未同步](https://stackoverflow.com/questions/48227941/document-not-in-sync-after-replace-text)。 此示例创建一个具有三对象类型关联的自定义关联对象。 它总共使用三个循环来构造相关对象的数组，并另外使用两个循环执行最终处理。 有 和 `for` `Array.forEach` 循环的混合。
- 尽管不严格是拆分循环或关联对象模式的示例，但还有一个高级 Excel 示例演示如何将一组单元格值转换为仅包含单个 货币的其他货币 `context.sync` 。 若要试用，请打开 Script Lab [工具](../overview/explore-with-script-lab.md)并导航到 **"货币转换器"** 示例。

## <a name="when-should-you-not-use-the-patterns-in-this-article"></a>何时 *不应* 使用本文中的模式？

Excel在给定调用 中读取的数据不能超过 5 `context.sync` MB。 如果超出此限制，将引发错误。  (有关详细信息，请参阅 Office 外接程序的资源限制和性能优化的["Excel](resource-limits-and-performance-optimization.md#excel-add-ins)外接程序"部分。) 接近此限制的情况很少见，但如果外接程序可能会发生这种情况，则代码不应在单个循环中加载所有数据，而是使用 循环执行 `context.sync` 。 但是，您仍应避免在集合 `context.sync` 对象的循环的每次迭代中都有 。 相反，请定义集合中项的子集，并循环遍历每个子集，在循环之间使用 `context.sync` 。 可以使用循环遍历子集的外部循环来构造此结构，并包含每个外部 `context.sync` 迭代中的 。
