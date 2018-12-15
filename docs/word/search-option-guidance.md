---
title: 使用搜索选项在 Word 加载项中查找文本
description: ''
ms.date: 07/20/2018
ms.openlocfilehash: d2c0fa2d542cd64986c2fd82f8a50a813f14610a
ms.sourcegitcommit: 3d8454055ba4d7aae12f335def97357dea5beb30
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/14/2018
ms.locfileid: "27270619"
---
# <a name="use-search-options-to-find-text-in-your-word-add-in"></a>使用搜索选项在 Word 加载项中查找文本 

加载项经常需要基于文档文本运行。
每种内容控件均有公开的搜索函数（这些内容控件包括 [Body](https://docs.microsoft.com/javascript/api/word/word.body?view=office-js)、[Paragraph](https://docs.microsoft.com/javascript/api/word/word.paragraph?view=office-js)、[Range](https://docs.microsoft.com/javascript/api/word/word.range?view=office-js)、[Table](https://docs.microsoft.com/javascript/api/word/word.table?view=office-js)、[TableRow](https://docs.microsoft.com/javascript/api/word/word.tablerow?view=office-js) 和基本 [ContentControl](https://docs.microsoft.com/javascript/api/word/word.contentcontrol?view=office-js) 对象）。 此函数接受一个代表所搜索文本的字符串（如通配符表达式）和 [SearchOptions](https://docs.microsoft.com/javascript/api/word/word.searchoptions?view=office-js) 对象。 它返回与搜索文本匹配的区域集合。

## <a name="search-options"></a>搜索选项
搜索选项为多个用于定义搜索参数处理方式的布尔值集合。 

| 属性     | 说明|
|:---------------|:----|
|ignorePunct|获取或设置一个值，该值指示是否忽略单词之间的标点符号的值。 对应于“查找和替换”对话框中的“忽略标点符号”复选框。|
|ignoreSpace|获取或设置一个值，该值指示是否忽略单词之间的所有空格。 对应于“查找和替换”对话框中的“忽略空格”复选框。|
|matchCase|获取或设置一个值，该值指示是否执行区分大小写搜索。 对应于“查找和替换”对话框中的“区分大小写”复选框。|
|matchPrefix|获取或设置一个值，该值指示是否匹配以搜索字符串开头的单词。 对应于“查找和替换”对话框中的“匹配前缀”复选框。|
|matchSuffix|获取或设置一个值，该值指示是否匹配以搜索字符串结尾的单词。 对应于“查找和替换”对话框中的“匹配后缀”复选框。|
|matchWholeWord|获取或设置一个值，该值用于指示是否查找操作仅限整个单词，而非较长单词的一部分的文字。 对应于“查找和替换”对话框中的“全字匹配”复选框。|
|matchWildcards|获取或设置一个值，该值指示搜索是否使用特殊搜索操作符执行。 对应于“查找和替换”对话框中的“使用通配符”复选框。|

## <a name="wildcard-guidance"></a>通配符指导
下表提供了与 Word JavaScript API 的搜索通配符相关的指导。

| 若要查找：         | 通配符 |  示例 |
|:-----------------|:--------|:----------|
| 任意单个字符| ? |s?t 找到 sat 和 set。 |
|任何字符的字符串| * |s*d 找到 sad 和 started。|
|单词的开头|< |<(inter) 找到 interesting 和 intercept，而不是 splintered。|
|单词结尾 |> |(in)> 找到 in 和 within，而不是 interesting。|
|一个指定的字符|[ ] |w[io]n 找到 win 和 won。|
|此区域中的任何单个字符| [-] |[r-t]ight 找到 right 和 sight。区域必须按升序排列。|
|除括号中区域内的字符以外的任何单个字符|[!x-z] |t[!a-m]ck 找到 tock 和 tuck，而不是 tack 或 tick。|
|前一个字符或表达式出现 n 次|{n} |fe{2}d 找到 feed，而不是 fed。|
|前一个字符或表达式至少出现 n 次|{n,} |fe{1,}d 找到 fed 和 feed。|
|前一个字符或表达式出现 n 到 m 次|{n,m} |10{1,3} 找到 10、100 和 1000。|
|前一个字符或表达式出现一次或多次|@ |lo@t 找到 lot 和 loot。|

### <a name="escaping-the-special-characters"></a>转义特殊字符

通配符搜索与正则表达式搜索大致相同。正则表达式中有特殊字符，包括“[”、“]”、“(”、“)”、“{”、“}”、“\*”、“?”、“<”、“>”、“!”和“@”。如果其中一个字符属于代码要搜索的文本字符串，则需要转义这个字符，以便让 Word 知道应该以文本形式（而不是作为正则表达式逻辑的一部分）处理这个字符。若要在 Word  UI 搜索中转义字符，请在字符前面添加“\'”字符。不过，若要以编程方式转义，请将字符置于“[]”字符之间。例如，“[\*]\*”搜索以“\*”开头、后跟任意数量的其他字符的所有字符串。 

## <a name="examples"></a>示例
下面示例演示常见情况。

### <a name="ignore-punctuation-search"></a>忽略标点符号搜索

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Queue a command to search the document and ignore punctuation.
    var searchResults = context.document.body.search('video you', {ignorePunct: true});

    // Queue a command to load the search results and get the font property values.
    context.load(searchResults, 'font');
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Found count: ' + searchResults.items.length);

        // Queue a set of commands to change the font for each found item.
        for (var i = 0; i < searchResults.items.length; i++) {
            searchResults.items[i].font.color = 'purple';
            searchResults.items[i].font.highlightColor = '#FFFF00'; //Yellow
            searchResults.items[i].font.bold = true;
        }
        
        // Synchronize the document state by executing the queued commands, 
        // and return a promise to indicate task completion.
        return context.sync();
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### <a name="search-based-on-a-prefix"></a>基于前缀搜索

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Queue a command to search the document based on a prefix.
    var searchResults = context.document.body.search('vid', {matchPrefix: true});

    // Queue a command to load the search results and get the font property values.
    context.load(searchResults, 'font');
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Found count: ' + searchResults.items.length);

        // Queue a set of commands to change the font for each found item.
        for (var i = 0; i < searchResults.items.length; i++) {
            searchResults.items[i].font.color = 'purple';
            searchResults.items[i].font.highlightColor = '#FFFF00'; //Yellow
            searchResults.items[i].font.bold = true;
        }
        
        // Synchronize the document state by executing the queued commands, 
        // and return a promise to indicate task completion.
        return context.sync();
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### <a name="search-based-on-a-suffix"></a>基于后缀搜索

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Queue a command to search the document for any string of characters after 'ly'.
    var searchResults = context.document.body.search('ly', {matchSuffix: true});

    // Queue a command to load the search results and get the font property values.
    context.load(searchResults, 'font');
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Found count: ' + searchResults.items.length);

        // Queue a set of commands to change the font for each found item.
        for (var i = 0; i < searchResults.items.length; i++) {
            searchResults.items[i].font.color = 'orange';
            searchResults.items[i].font.highlightColor = 'black';
            searchResults.items[i].font.bold = true;
        }
        
        // Synchronize the document state by executing the queued commands, 
        // and return a promise to indicate task completion.
        return context.sync();
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### <a name="search-using-a-wildcard"></a>使用通配符搜索

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Queue a command to search the document with a wildcard
    // for any string of characters that starts with 'to' and ends with 'n'.
    var searchResults = context.document.body.search('to*n', {matchWildCards: true});

    // Queue a command to load the search results and get the font property values.
    context.load(searchResults, 'font');
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Found count: ' + searchResults.items.length);

        // Queue a set of commands to change the font for each found item.
        for (var i = 0; i < searchResults.items.length; i++) {
            searchResults.items[i].font.color = 'purple';
            searchResults.items[i].font.highlightColor = 'pink';
            searchResults.items[i].font.bold = true;
        }
        
        // Synchronize the document state by executing the queued commands, 
        // and return a promise to indicate task completion.
        return context.sync();
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

更多信息请参阅 [Word JavaScript API 参考](https://docs.microsoft.com/office/dev/add-ins/reference/overview/word-add-ins-reference-overview?view=office-js).