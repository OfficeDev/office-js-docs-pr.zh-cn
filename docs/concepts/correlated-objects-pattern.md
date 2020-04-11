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
# <a name="avoid-using-the-contextsync-method-in-loops"></a>避免在循环中使用 context。 sync 方法

> [!NOTE]
> 本文假定您不在第一阶段使用至少使用用于 Excel、Word、OneNote 和 Visio&mdash;&mdash;的四个特定于 Excel 的 Office JavaScript api 中的一个，这些 api 使用批处理系统与 Office 文档进行交互。 特别是，您应该知道什么是调用`context.sync` ，您应该知道什么是集合对象。 如果你不在这一阶段，请先了解本文中的 "特定于主机" 下的 " [Office JAVASCRIPT API](../develop/understanding-the-javascript-api-for-office.md) " 和 "文档"。

对于使用特定于主机的 API 模型之一（针对 Excel、Word、OneNote 和 Visio）的 Office 外接程序中的某些编程方案，您的代码需要从集合对象的每个成员中读取、写入或处理某些属性。 例如，需要获取特定表列或 Word 外接程序中每个单元格的值的 Excel 加载项，需要突出显示文档中每个字符串的实例。 您需要在集合对象的`items`属性中循环访问这些成员;但是，出于性能方面的考虑，您需要避免`context.sync`在循环的每个迭代中调用。 每次调用`context.sync`的是从外接端到 Office 文档的一种往返行程。 重复往返行程会影响性能，尤其是当加载项在 web 上的 Office 中运行时，由于往返行程跨 internet 进行。

> [!NOTE]
> 本文中的所有示例都`for`使用循环，但所述的实践适用于可循环访问数组的任何循环语句，其中包括以下内容：
>
> - `for`
> - `for of`
> - `while`
> - `do while`
> 
> 它们还适用于将函数传递和应用于数组中的项的任何数组方法，包括以下内容：
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

在最简单的情况下，只写入集合对象的成员，而不是读取其属性。 例如，下面的代码在 Word 文档中突出显示了每个 "the" 实例的黄色。 

> [!NOTE]
> 通常，最好`context.sync`先在主机`run`方法的结尾 "}" 字符前加上 final （如`Excel.run` `Word.run`，等）。 这是因为此`run`方法会将`context.sync`作为最后一件事情的隐藏调用作为最后一件事情，并且只有在已排队的命令尚未同步的情况下。 此调用是隐藏的这一事实可能会造成混淆，因此我们通常建议您添加显式`context.sync`。 但是，假设本文涉及最小化的调用， `context.sync`则添加完全不必要的最终版本`context.sync`会更加容易混淆。 因此，在本文的结尾处没有未同步的命令时，我们会将其保留下来`run`。 

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

上面的代码花了1个完整的秒，在 Windows 中的 Word 中有一个包含200个实例 "the" 的文档。 但是，当`await context.sync();`循环中的行被注释掉并在循环 uncommented 后的相同行时，该操作只需 1/10 秒。 在 web 上的 Word 中（使用边缘作为浏览器），循环内的同步花费了3个完整的秒，并且在循环后的同步速度仅为 6/10ths，而在循环后的速度约为5倍。 在包含2000实例 "the" 的文档中，使用循环中的同步（在 web 中为 "网页"）80秒，并且在循环后的同步只需4秒，速度将加快20倍。

> [!NOTE]
> 如果同步并发运行，则需要询问是否会更快地执行同步内部循环版本，这只需从的`await` `context.sync()`前面删除关键字即可完成。 这将导致运行时启动同步，然后立即开始循环的下一个迭代，而无需等待同步完成。 但是，这并不像出于这些原因而完全移`context.sync`出循环之外的解决方案：
>
> - 正如同步批处理作业中的命令已排入队列中一样，批处理作业本身在 Office 中排队，但在队列中的批处理作业不支持超过50个。 任何其他触发器错误。 因此，如果循环中的迭代数超过50个，则会有可能超出队列大小。 迭代次数越多，发生此问题的可能性就越大。 
> - "并发" 并不同时表示。 执行多个同步操作所需的时间要比执行一个同步操作花费更长时间。
> - 并发操作不能保证按照其启动顺序完成。 在上面的示例中，突出显示 "the" 一词的顺序无关紧要，但在某些情况下，将按顺序处理集合中的项目很重要。

## <a name="reading-values-from-the-document-with-the-split-loop-pattern"></a>使用拆分循环模式从文档中读取值

当`context.sync`代码必须在处理每个集合项的属性时*读取*这些集合项的属性时，避免 s 在循环中变得更具挑战性。 假设您的代码需要对 Word 文档中的所有内容控件进行迭代，并记录与每个控件关联的第一个段落的文本。 编程 instincts 可能会引导您在控件上循环，加载每个`text` （第一个）段落的属性，调用`context.sync`使用文档中的文本填充代理段落对象，然后将其记录下来。 示例如下。

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

在这种情况下，为了避免`context.sync`在循环中使用，应使用一种模式来调用**拆分循环**模式。 我们来看看该模式的具体示例，然后再获取该模式的正式说明。 下面介绍了拆分循环模式如何应用于前面的代码段。 关于此代码，请注意以下几点：

- 现在有两个循环`context.sync` ，它们之间存在，因此不`context.sync`会出现在任何循环中。
- 第一个循环可循环访问 collection 对象中的项目，并像`text`原始循环那样加载该属性，但第一个循环无法记录段落文本，因为它不再包含`context.sync`用于填充`text` `paragraph`代理对象的属性。 而是将`paragraph`对象添加到数组中。
- 第二个循环可循环访问第一个循环创建的数组，并记录`text`每个`paragraph`项目的。 这是可行的， `context.sync`这是因为两个循环之间的`text`属性都填充了所有属性。

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

上面的示例建议了以下过程，用于打开包含`context.sync`拆分循环模式的循环： 

1. 将循环替换为两个循环。
2. 创建第一个循环以对集合进行迭代，并将每个项添加到数组中，同时还加载代码需要读取的项的任何属性。 
3. 在第一个循环之后， `context.sync`调用以使用任何加载的属性填充代理对象。 
4. `context.sync`执行第二个循环，以循环访问在第一个循环中创建的数组并读取加载的属性。

## <a name="processing-objects-in-the-document-with-the-correlated-objects-pattern"></a>使用关联对象模式处理文档中的对象

让我们考虑更复杂的情况，即处理集合中的项目需要的数据不在项目本身中。 方案假设一个 Word 加载项，该加载项对使用某些样本文字的模板创建的文档进行操作。 分散在文本中的是以下占位符字符串的一个或多个实例： "{协调器}"、"{Deputy}" 和 "{Manager}"。 加载项会将每个占位符替换为某人的姓名。 外接端的 UI 对本文并不重要。 例如，它可能有一个具有三个文本框的任务窗格，每个文本框标有一个占位符。 用户在每个文本框中输入一个名称，然后按下一个 "**替换**" 按钮。 该按钮的处理程序将创建一个将名称映射到占位符的数组，然后将每个占位符替换为分配的名称。 

您无需实际生成具有此 UI 的外接程序，即可试用代码。 您可以使用[脚本实验室工具](../overview/explore-with-script-lab.md)对重要代码进行原型。 使用以下赋值语句创建映射数组。

```javascript
const jobMapping = [
        { job: "{Coordinator}", person: "Sally" },
        { job: "{Deputy}", person: "Bob" },
        { job: "{Manager}", person: "Kim" }
    ];
```

下面的代码演示在使用`context.sync`内部循环时，如何将每个占位符替换为其分配的名称。

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

在上面的代码中，有一个外部循环和一个内层循环。 其中每个都包含`context.sync`一个。 根据本文中的第一个代码段，您可能会发现在内部循环`context.sync`中，可以在 inner 循环之后直接移动到内部循环中。 但在外部循环中，此代码仍`context.sync`会保留（其中两个）。 下面的代码演示如何从循环中`context.sync`删除。 我们将讨论下面的代码。

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

注释代码使用拆分循环模式：

- 前一示例中的外部循环已拆分为两个。 （第二个循环有一个内部循环，因为代码在一组工作（或多个占位符）上进行迭代，并且在该集合中对匹配区域进行迭代。
- 每个重大`context.sync`循环之后都有一个，但`context.sync`在任何循环中都不存在。 
- 第二个主要循环可循环访问在第一个循环中创建的数组。

但是，在第一个循环中创建的数组*不*包含一个 Office 对象，因为在[使用拆分循环模式的文档中读取值](#reading-values-from-the-document-with-the-split-loop-pattern)的节中的第一个循环。 这是因为处理 Word Range 对象所需的一些信息不在 Range 对象本身中，而是来自于`jobMapping`数组。 

因此，在第一个循环中创建的数组中的对象是具有两个属性的自定义对象。 第一个是与特定职务（即占位符字符串）匹配的单词范围的数组，第二个是提供分配到该作业的人员姓名的字符串。 这使得最终循环易于编写和易于阅读，因为处理给定区域所需的全部信息都包含在包含该范围的同一自定义对象中。 应替换_ **correlatedObject**[j]_ 的名称是同一对象的另一个属性： _ **correlatedObject**_。 

我们称之为 "**关联对象**" 模式的拆分循环模式的这一变体。 一般来讲，第一条循环创建自定义对象的数组。 每个对象都有一个属性，其值是 Office collection 对象（或此类项目的数组）中的项目之一。 自定义对象具有其他属性，每个属性都提供处理最终循环中的 Office 对象所需的信息。 请参阅[这些模式的其他示例](#other-examples-of-these-patterns)部分，以获取自定义关联对象具有两个以上属性的示例的链接。

另一个需要注意的一点是，有时需要多个循环来创建自定义关联对象的数组。 如果您需要只读取一个 Office 集合对象的每个成员的属性来收集将用于处理另一个集合对象的信息，则会发生这种情况。 （例如，您的代码需要读取 Excel 表中所有列的标题，因为您的外接程序将根据该列的标题对某些列的单元格应用数字格式。）但您始终可以在循环`context.sync`之间，而不是在循环之间保持 s。 有关示例，请参阅[这些模式的其他示例](#other-examples-of-these-patterns)一节。

## <a name="other-examples-of-these-patterns"></a>这些模式的其他示例

- 有关使用`Array.forEach`循环的 Excel 的非常简单的示例，请参阅此堆栈溢出问题的接受答案：[是否可以对多个上下文进行排队。在 context 之前进行加载？](https://stackoverflow.com/questions/44459604/is-it-possible-to-queue-more-than-one-context-load-before-context-sync)
- 有关使用`Array.forEach`循环但不使用`async` / `await`语法的 Word 的简单示例，请参阅 "接受的对此堆栈溢出问题的答案：使用[Office JavaScript API 循环访问包含内容控件的所有段落](https://stackoverflow.com/questions/58422113/iterating-over-all-paragraphs-with-content-controls-with-office-javascript-api)"。
- 有关使用 TypeScript 编写的 Word 的示例，请参阅示例[Word 外接程序 Angular2 样式检查器](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker)，尤其是文件 " [document](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker/blob/master/app/services/word-document/word.document.service.ts)"。 它混合了`for`和`Array.forEach`循环。
- 对于高级 Word 示例，请将[此 gist](https://gist.github.com/9c5a803e52480ec7f00bb3224292e0ab)导入[脚本实验室工具](../overview/explore-with-script-lab.md)。 有关使用 gist 的上下文，请参阅在[替换文本后，不同步](https://stackoverflow.com/questions/48227941/document-not-in-sync-after-replace-text)"堆栈溢出问题" 文档中的 "已接受的答案"。 本示例创建一个具有三个属性的自定义关联对象类型。 它总共使用三个循环来构造相关对象的数组，以及执行最后处理的两个更多循环。 混合了`for`和`Array.forEach`循环。
- 尽管不是严格的拆分循环或相关对象模式的示例，但还有一个演示如何将一组单元格的值转换为只使用一个的其他货币的高级 Excel `context.sync`示例。 若要尝试，请打开[脚本实验室工具](../overview/explore-with-script-lab.md)并导航到**货币转换器**示例。 

## <a name="when-should-you-not-use-the-patterns-in-this-article"></a>何时*应使用本文*中的模式？

Excel 在给定的`context.sync`调用中无法读取超过 5 MB 的数据。 如果超过此限制，则会引发错误。 （有关详细信息，请参阅[Excel data transfer 限制](../develop/common-coding-issues.md#excel-data-transfer-limits)。）很少需要此限制，但如果有机会在外接程序中执行此操作，则代码*不*应在单个循环中加载所有数据，并在循环中使用 a `context.sync`。 但您仍应避免`context.sync`在集合对象上循环的每个迭代。 相反，在集合中定义项的子集，并依次对每个子集进行循环，并`context.sync`在循环之间进行循环。 您可以使用外部循环对此进行构造，该循环可对子集进行`context.sync`迭代，并在每个外部迭代中包含。
