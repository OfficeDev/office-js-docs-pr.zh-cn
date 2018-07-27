---
title: 使用搜索选项在 Word 加载项中查找文本
description: ''
ms.date: 7/20/2018
ms.openlocfilehash: 9dcd5e42de9cc0816797a4a14b40a0e3e376f158
ms.sourcegitcommit: eea7f2b1679cf9a209d35880b906e311bdf1359c
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/26/2018
ms.locfileid: "21254859"
---
# <a name="use-search-options-to-find-text-in-your-word-add-in"></a>使用搜索选项在 Word 加载项中查找文本 

加载项经常需要根据文档的文本进行操作。
每个内容控件都会暴露搜索函数（包括 [ Body](https://dev.office.com/reference/add-ins/word/body)、[Paragraph](https://dev.office.com/reference/add-ins/word/paragraph)、[Range](https://dev.office.com/reference/add-ins/word/range)、[Table](https://dev.office.com/reference/add-ins/word/table)、[TableRow](https://dev.office.com/reference/add-ins/word/tablerow) 和基本的 [ContentControl](https://dev.office.com/reference/add-ins/word/contentcontrol) 对象）。 此函数接受表示要搜索的文本字符串（或 wldcard 表达式）和一个 [SearchOptions](https://dev.office.com/reference/add-ins/word/searchoptions) 对象。 它返回与搜索文本匹配的范围集合。

## <a name="search-options"></a>搜索选项
搜索选项是一组布尔值，用于定义应如何处理搜索参数。 

| 属性     | 说明|
|:---------------|:----|
|ignorePunct|获取或设置一个值，该值指示是否忽略单词之间的所有标点符号。 对应于“查找和替换”对话框中的“忽略标点字符”复选框。|
|ignoreSpace|获取或设置一个值，该值指示是否忽略单词之间的所有标点符号。 对应于“查找和替换”对话框中的“忽略空格字符”复选框。|
|matchCase|获取或设置一个值，该值指示是否执行区分大小写的搜索。 对应于“查找和替换”对话框中的“区分大小写”复选框。|
|matchPrefix|获取或设置一个值，该值指示是否匹配以搜索字符串开头的单词。 对应于“查找和替换”对话框中的“区分前缀”复选框。|
|matchSuffix|获取或设置一个值，该值指示是否匹配以搜索字符串结束的单词。 对应于“查找和替换”对话框中的“区分后缀”复选框。|
|matchWholeWord|获取或设置一个值，该值指示是仅查找整个单词的操作，而不是作为较大单词的一部分的文本。 对应于“查找和替换”对话框中的“全字匹配”复选框。|
|matchWildcards|获取或设置一个值，该值指示是否使用特殊搜索运算符执行搜索。 对应于“查找和替换”对话框中的“使用通配符”复选框。|

## <a name="wildcard-guidance"></a>通配符指导
下表提供了有关 Word JavaScript API 的搜索通配符的指导。

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
|前一个字符或表达式出现 n 到 m 次|{n,m} |10{1,3} 找到10、100 和 1000。|
|前一个字符或表达式出现一次或多次|@ |lo@t 找到 lot 和 loot。|

### <a name="escaping-the-special-characters"></a>转义特殊字符

通配符搜索与正则表达式搜索大致相同。正则表达式中有特殊字符，包括“[”、“]”、“(”、“)”、“{”、“}”、“\*”、“?”、“<”、“>”、“!”和“@”。如果其中一个字符属于代码要搜索的文本字符串，则需要转义这个字符，以便让 Word 知道应该以文本形式（而不是作为正则表达式逻辑的一部分）处理这个字符。若要在 Word  UI 搜索中转义字符，请在字符前面添加“\'”字符。不过，若要以编程方式转义，请将字符置于“[]”字符之间。例如，“[\*]\*”搜索以“\*”开头、后跟任意数量的其他字符的所有字符串。 

## <a name="examples"></a>示例
以下示例演示了常见方案。

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

更多信息请参见 [Word JavaScript 参考 API](https://dev.office.com/reference/add-ins/word/word-add-ins-reference-overview)。