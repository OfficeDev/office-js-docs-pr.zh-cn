---
title: 在 Word 加载项中使用搜索选项查找文本
description: 了解如何在 Word 加载项中使用搜索选项。
ms.date: 02/28/2022
ms.localizationpriority: medium
ms.openlocfilehash: 64ffd3b32329dae98f869abaabcb3218e57a4a34
ms.sourcegitcommit: b6a3815a1ad17f3522ca35247a3fd5d7105e174e
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/22/2022
ms.locfileid: "66958998"
---
# <a name="use-search-options-in-your-word-add-in-to-find-text"></a>在 Word 加载项中使用搜索选项查找文本

加载项经常需要基于文档文本运行。 搜索方法由每个内容控件公开 (这包括 [Body](/javascript/api/word/word.body)、 [Paragraph](/javascript/api/word/word.paragraph)、 [Range](/javascript/api/word/word.range)、 [Table](/javascript/api/word/word.table)、 [TableRow](/javascript/api/word/word.tablerow) 和基本 [ContentControl](/javascript/api/word/word.contentcontrol) 对象) 。 此方法采用字符串 (或通配符表达式) 表示要搜索的文本和 [SearchOptions](/javascript/api/word/word.searchoptions) 对象。 它返回与搜索文本匹配的区域集合。

## <a name="search-options"></a>搜索选项

搜索选项为多个用于定义搜索参数处理方式的布尔值集合。

| 属性       | 说明|
|:---------------|:----|
|ignorePunct|获取或设置一个值，该值指示是否忽略单词之间的标点符号的值。 对应于“ **查找和替换** ”对话框中的“忽略标点字符”复选框。|
|ignoreSpace|获取或设置一个值，该值指示是否忽略单词之间的所有空格。 对应于“ **查找和替换** ”对话框中的“忽略空白字符”复选框。|
|matchCase|获取或设置一个值，该值指示是否执行区分大小写的搜索。 对应于“ **查找和替换** ”对话框中的“匹配案例”复选框。|
|matchPrefix|获取或设置一个值，该值指示是否匹配以搜索字符串开头的单词。 对应于“ **查找和替换** ”对话框中的“匹配前缀”复选框。|
|matchSuffix|获取或设置一个值，该值指示是否匹配以搜索字符串结尾的单词。 对应于“ **查找和替换** ”对话框中的“匹配后缀”复选框。|
|matchWholeWord|获取或设置一个值，该值用于指示是否查找操作仅限整个单词，而非较长单词的一部分的文字。 对应于“ **查找和替换** ”对话框中的“仅查找整个单词”复选框。|
|matchWildcards|获取或设置一个值，该值指示搜索是否使用特殊搜索操作符执行。 对应于“ **查找和替换** ”对话框中的“使用通配符”复选框。|

## <a name="wildcard-guidance"></a>通配符指导

下表提供了与 Word JavaScript API 的搜索通配符相关的指导。

| 若要查找：         | 通配符 |  示例 |
|:-----------------|:--------|:----------|
|任意单个字符| ? |s?t 找到 sat 和 set。 |
|任何字符的字符串| * |s*d 找到 sad 和 started。|
|单词的开头|< |<(inter) 找到 interesting 和 intercept，而不是 splintered。|
|单词结尾 |> |(in)> 找到 in 和 within，而不是 interesting。|
|一个指定的字符|[ ] |w[io]n 找到 win 和 won。|
|此区域中的任何单个字符| [-] |[r-t]ight 找到 right 和 sight。区域必须按升序排列。|
|除括号中区域内的字符以外的任何单个字符|[!x-z] |t[!a-m]ck 找到 tock 和 tuck，而不是 tack 或 tick。|
|上一个字符或表达式的恰好 *n* 个匹配项|{n} |fe{2}d 找到 feed，而不是 fed。|
|上一个字符或表达式至少 *出现 n* 次|{n,} |fe{1,}d 找到 fed 和 feed。|
|上一个字符或表达式的 *从 n* 到 *m* 的匹配项|{n,m} |10{1,3} 找到 10、100 和 1000。|
|前一个字符或表达式出现一次或多次|@ |lo@t 找到 lot 和 loot。|

### <a name="escaping-special-characters"></a>转义特殊字符

通配符搜索实质上与在正则表达式上搜索相同。 正则表达式中有特殊字符，包括“[”、“]”、“ (”、“) ”、“{”、“}\*”、“？”、“<”、“>”、“！”和“@”。 如果其中一个字符是代码正在搜索的文本字符串的一部分，则需要转义它，以便 Word 知道应该从字面上处理它，而不是作为正则表达式逻辑的一部分。 若要在 Word UI 搜索中转义某个字符，请先使用反斜杠字符 (“\\) ，但若要以编程方式将其转义，请将其放在”[]“字符之间。 例如，“[\*]\*”搜索以“”\*开头的任何字符串，后跟任意数目的其他字符。

## <a name="examples"></a>示例

下面示例演示常见情况。

### <a name="ignore-punctuation-search"></a>忽略标点符号搜索

```js
// Run a batch operation against the Word object model.
await Word.run(async (context) => {

    // Queue a command to search the document and ignore punctuation.
    const searchResults = context.document.body.search('video you', {ignorePunct: true});

    // Queue a command to load the font property values.
    searchResults.load('font');

    // Synchronize the document state.
    await context.sync();
    console.log('Found count: ' + searchResults.items.length);

    // Queue a set of commands to change the font for each found item.
    for (let i = 0; i < searchResults.items.length; i++) {
        searchResults.items[i].font.color = 'purple';
        searchResults.items[i].font.highlightColor = '#FFFF00'; //Yellow
        searchResults.items[i].font.bold = true;
    }

    // Synchronize the document state.
    await context.sync();
});
```

### <a name="search-based-on-a-prefix"></a>基于前缀搜索

```js
// Run a batch operation against the Word object model.
await Word.run(async (context) => {

    // Queue a command to search the document based on a prefix.
    const searchResults = context.document.body.search('vid', {matchPrefix: true});

    // Queue a command to load the font property values.
    searchResults.load('font');

    // Synchronize the document state.
    await context.sync();
    console.log('Found count: ' + searchResults.items.length);

    // Queue a set of commands to change the font for each found item.
    for (let i = 0; i < searchResults.items.length; i++) {
        searchResults.items[i].font.color = 'purple';
        searchResults.items[i].font.highlightColor = '#FFFF00'; //Yellow
        searchResults.items[i].font.bold = true;
    }

    // Synchronize the document state.
    await context.sync();
});
```

### <a name="search-based-on-a-suffix"></a>基于后缀搜索

```js
// Run a batch operation against the Word object model.
await Word.run(async (context) => {

    // Queue a command to search the document for any string of characters after 'ly'.
    const searchResults = context.document.body.search('ly', {matchSuffix: true});

    // Queue a command to load the font property values.
    searchResults.load('font');

    // Synchronize the document state.
    await context.sync();
    console.log('Found count: ' + searchResults.items.length);

    // Queue a set of commands to change the font for each found item.
    for (let i = 0; i < searchResults.items.length; i++) {
        searchResults.items[i].font.color = 'orange';
        searchResults.items[i].font.highlightColor = 'black';
        searchResults.items[i].font.bold = true;
    }

    // Synchronize the document state.
    await context.sync();
});
```

### <a name="search-using-a-wildcard"></a>使用通配符搜索

```js
// Run a batch operation against the Word object model.
await Word.run(async (context) => {

    // Queue a command to search the document with a wildcard
    // for any string of characters that starts with 'to' and ends with 'n'.
    const searchResults = context.document.body.search('to*n', {matchWildcards: true});

    // Queue a command to load the font property values.
    searchResults.load('font');

    // Synchronize the document state.
    await context.sync();
    console.log('Found count: ' + searchResults.items.length);

    // Queue a set of commands to change the font for each found item.
    for (let i = 0; i < searchResults.items.length; i++) {
        searchResults.items[i].font.color = 'purple';
        searchResults.items[i].font.highlightColor = 'pink';
        searchResults.items[i].font.bold = true;
    }

    // Synchronize the document state.
    await context.sync();
});
```

更多信息请参阅 [Word JavaScript API 参考](../reference/overview/word-add-ins-reference-overview.md).
