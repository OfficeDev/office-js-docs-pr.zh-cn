---
title: 使用搜索选项在 Word 加载项中查找文本
description: ''
ms.date: 09/27/2019
localization_priority: Normal
ms.openlocfilehash: 213853af31ae7ae15ad3f6386da70f22698d421d
ms.sourcegitcommit: d15bca2c12732f8599be2ec4b2adc7c254552f52
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/12/2020
ms.locfileid: "41950479"
---
# <a name="use-search-options-to-find-text-in-your-word-add-in"></a><span data-ttu-id="eeadf-102">使用搜索选项在 Word 加载项中查找文本</span><span class="sxs-lookup"><span data-stu-id="eeadf-102">Use search options to find text in your Word add-in</span></span>

<span data-ttu-id="eeadf-103">加载项经常需要基于文档文本运行。</span><span class="sxs-lookup"><span data-stu-id="eeadf-103">Add-ins frequently need to act based on the text of a document.</span></span>
<span data-ttu-id="eeadf-104">每种内容控件均有公开的搜索函数（这些内容控件包括 [Body](/javascript/api/word/word.body)、[Paragraph](/javascript/api/word/word.paragraph)、[Range](/javascript/api/word/word.range)、[Table](/javascript/api/word/word.table)、[TableRow](/javascript/api/word/word.tablerow) 和基本 [ContentControl](/javascript/api/word/word.contentcontrol) 对象）。</span><span class="sxs-lookup"><span data-stu-id="eeadf-104">A search function is exposed by every content control (this includes [Body](/javascript/api/word/word.body), [Paragraph](/javascript/api/word/word.paragraph), [Range](/javascript/api/word/word.range), [Table](/javascript/api/word/word.table), [TableRow](/javascript/api/word/word.tablerow), and the base [ContentControl](/javascript/api/word/word.contentcontrol) object).</span></span> <span data-ttu-id="eeadf-105">此函数接受一个代表所搜索文本的字符串（如通配符表达式）和 [SearchOptions](/javascript/api/word/word.searchoptions) 对象。</span><span class="sxs-lookup"><span data-stu-id="eeadf-105">This function takes in a string (or wildcard expression) representing the text you are searching for and a [SearchOptions](/javascript/api/word/word.searchoptions) object.</span></span> <span data-ttu-id="eeadf-106">它返回与搜索文本匹配的区域集合。</span><span class="sxs-lookup"><span data-stu-id="eeadf-106">It returns a collection of ranges which match the search text.</span></span>

## <a name="search-options"></a><span data-ttu-id="eeadf-107">搜索选项</span><span class="sxs-lookup"><span data-stu-id="eeadf-107">Search options</span></span>

<span data-ttu-id="eeadf-108">搜索选项为多个用于定义搜索参数处理方式的布尔值集合。</span><span class="sxs-lookup"><span data-stu-id="eeadf-108">The search options are a collection of boolean values defining how the search parameter should be treated.</span></span>

| <span data-ttu-id="eeadf-109">属性</span><span class="sxs-lookup"><span data-stu-id="eeadf-109">Property</span></span>     | <span data-ttu-id="eeadf-110">说明</span><span class="sxs-lookup"><span data-stu-id="eeadf-110">Description</span></span>|
|:---------------|:----|
|<span data-ttu-id="eeadf-111">ignorePunct</span><span class="sxs-lookup"><span data-stu-id="eeadf-111">ignorePunct</span></span>|<span data-ttu-id="eeadf-112">获取或设置一个值，该值指示是否忽略单词之间的标点符号的值。</span><span class="sxs-lookup"><span data-stu-id="eeadf-112">Gets or sets a value indicating whether to ignore all punctuation characters between words.</span></span> <span data-ttu-id="eeadf-113">对应于“查找和替换”对话框中的“忽略标点符号”复选框。</span><span class="sxs-lookup"><span data-stu-id="eeadf-113">Corresponds to the "Ignore punctuation characters" check box in the Find and Replace dialog box.</span></span>|
|<span data-ttu-id="eeadf-114">ignoreSpace</span><span class="sxs-lookup"><span data-stu-id="eeadf-114">ignoreSpace</span></span>|<span data-ttu-id="eeadf-115">获取或设置一个值，该值指示是否忽略单词之间的所有空格。</span><span class="sxs-lookup"><span data-stu-id="eeadf-115">Gets or sets a value indicating whether to ignore all whitespace between words.</span></span> <span data-ttu-id="eeadf-116">对应于“查找和替换”对话框中的“忽略空格”复选框。</span><span class="sxs-lookup"><span data-stu-id="eeadf-116">Corresponds to the "Ignore white-space characters" check box in the Find and Replace dialog box.</span></span>|
|<span data-ttu-id="eeadf-117">matchCase</span><span class="sxs-lookup"><span data-stu-id="eeadf-117">matchCase</span></span>|<span data-ttu-id="eeadf-118">获取或设置一个值，该值指示是否执行区分大小写搜索。</span><span class="sxs-lookup"><span data-stu-id="eeadf-118">Gets or sets a value indicating whether to perform a case sensitive search.</span></span> <span data-ttu-id="eeadf-119">对应于“查找和替换”对话框中的“区分大小写”复选框。</span><span class="sxs-lookup"><span data-stu-id="eeadf-119">Corresponds to the "Match case" check box in the Find and Replace dialog box.</span></span>|
|<span data-ttu-id="eeadf-120">matchPrefix</span><span class="sxs-lookup"><span data-stu-id="eeadf-120">matchPrefix</span></span>|<span data-ttu-id="eeadf-121">获取或设置一个值，该值指示是否匹配以搜索字符串开头的单词。</span><span class="sxs-lookup"><span data-stu-id="eeadf-121">Gets or sets a value indicating whether to match words that begin with the search string.</span></span> <span data-ttu-id="eeadf-122">对应于“查找和替换”对话框中的“匹配前缀”复选框。</span><span class="sxs-lookup"><span data-stu-id="eeadf-122">Corresponds to the "Match prefix" check box in the Find and Replace dialog box.</span></span>|
|<span data-ttu-id="eeadf-123">matchSuffix</span><span class="sxs-lookup"><span data-stu-id="eeadf-123">matchSuffix</span></span>|<span data-ttu-id="eeadf-124">获取或设置一个值，该值指示是否匹配以搜索字符串结尾的单词。</span><span class="sxs-lookup"><span data-stu-id="eeadf-124">Gets or sets a value indicating whether to match words that end with the search string.</span></span> <span data-ttu-id="eeadf-125">对应于“查找和替换”对话框中的“匹配后缀”复选框。</span><span class="sxs-lookup"><span data-stu-id="eeadf-125">Corresponds to the "Match suffix" check box in the Find and Replace dialog box.</span></span>|
|<span data-ttu-id="eeadf-126">matchWholeWord</span><span class="sxs-lookup"><span data-stu-id="eeadf-126">matchWholeWord</span></span>|<span data-ttu-id="eeadf-127">获取或设置一个值，该值用于指示是否查找操作仅限整个单词，而非较长单词的一部分的文字。</span><span class="sxs-lookup"><span data-stu-id="eeadf-127">Gets or sets a value indicating whether to find operation only entire words, not text that is part of a larger word.</span></span> <span data-ttu-id="eeadf-128">对应于“查找和替换”对话框中的“全字匹配”复选框。</span><span class="sxs-lookup"><span data-stu-id="eeadf-128">Corresponds to the "Find whole words only" check box in the Find and Replace dialog box.</span></span>|
|<span data-ttu-id="eeadf-129">matchWildcards</span><span class="sxs-lookup"><span data-stu-id="eeadf-129">matchWildcards</span></span>|<span data-ttu-id="eeadf-130">获取或设置一个值，该值指示搜索是否使用特殊搜索操作符执行。</span><span class="sxs-lookup"><span data-stu-id="eeadf-130">Gets or sets a value indicating whether the search will be performed using special search operators.</span></span> <span data-ttu-id="eeadf-131">对应于“查找和替换”对话框中的“使用通配符”复选框。</span><span class="sxs-lookup"><span data-stu-id="eeadf-131">Corresponds to the "Use wildcards" check box in the Find and Replace dialog box.</span></span>|

## <a name="wildcard-guidance"></a><span data-ttu-id="eeadf-132">通配符指导</span><span class="sxs-lookup"><span data-stu-id="eeadf-132">Wildcard guidance</span></span>

<span data-ttu-id="eeadf-133">下表提供了与 Word JavaScript API 的搜索通配符相关的指导。</span><span class="sxs-lookup"><span data-stu-id="eeadf-133">The following table provides guidance around the Word JavaScript API's search wildcards.</span></span>

| <span data-ttu-id="eeadf-134">若要查找：</span><span class="sxs-lookup"><span data-stu-id="eeadf-134">To find:</span></span>         | <span data-ttu-id="eeadf-135">通配符</span><span class="sxs-lookup"><span data-stu-id="eeadf-135">Wildcard</span></span> |  <span data-ttu-id="eeadf-136">示例</span><span class="sxs-lookup"><span data-stu-id="eeadf-136">Sample</span></span> |
|:-----------------|:--------|:----------|
| <span data-ttu-id="eeadf-137">任意单个字符</span><span class="sxs-lookup"><span data-stu-id="eeadf-137">Any single character</span></span>| <span data-ttu-id="eeadf-138">?</span><span class="sxs-lookup"><span data-stu-id="eeadf-138">?</span></span> |<span data-ttu-id="eeadf-139">s?t 找到 sat 和 set。</span><span class="sxs-lookup"><span data-stu-id="eeadf-139">s?t finds sat and set.</span></span> |
|<span data-ttu-id="eeadf-140">任何字符的字符串</span><span class="sxs-lookup"><span data-stu-id="eeadf-140">Any string of characters</span></span>| * |<span data-ttu-id="eeadf-141">s\*d 找到 sad 和 started。</span><span class="sxs-lookup"><span data-stu-id="eeadf-141">s\*d finds sad and started.</span></span>|
|<span data-ttu-id="eeadf-142">单词的开头</span><span class="sxs-lookup"><span data-stu-id="eeadf-142">The beginning of a word</span></span>|< |<span data-ttu-id="eeadf-143"><(inter) 找到 interesting 和 intercept，而不是 splintered。</span><span class="sxs-lookup"><span data-stu-id="eeadf-143"><(inter) finds interesting and intercept, but not splintered.</span></span>|
|<span data-ttu-id="eeadf-144">单词结尾</span><span class="sxs-lookup"><span data-stu-id="eeadf-144">The end of a word</span></span> |> |<span data-ttu-id="eeadf-145">(in)> 找到 in 和 within，而不是 interesting。</span><span class="sxs-lookup"><span data-stu-id="eeadf-145">(in)> finds in and within, but not interesting.</span></span>|
|<span data-ttu-id="eeadf-146">一个指定的字符</span><span class="sxs-lookup"><span data-stu-id="eeadf-146">One of the specified characters</span></span>|<span data-ttu-id="eeadf-147">[ ]</span><span class="sxs-lookup"><span data-stu-id="eeadf-147">[ ]</span></span> |<span data-ttu-id="eeadf-148">w[io]n 找到 win 和 won。</span><span class="sxs-lookup"><span data-stu-id="eeadf-148">w[io]n finds win and won.</span></span>|
|<span data-ttu-id="eeadf-149">此区域中的任何单个字符</span><span class="sxs-lookup"><span data-stu-id="eeadf-149">Any single character in this range</span></span>| <span data-ttu-id="eeadf-150">[-]</span><span class="sxs-lookup"><span data-stu-id="eeadf-150">[-]</span></span> |<span data-ttu-id="eeadf-p109">[r-t]ight 找到 right 和 sight。区域必须按升序排列。</span><span class="sxs-lookup"><span data-stu-id="eeadf-p109">[r-t]ight finds right and sight. Ranges must be in ascending order.</span></span>|
|<span data-ttu-id="eeadf-153">除括号中区域内的字符以外的任何单个字符</span><span class="sxs-lookup"><span data-stu-id="eeadf-153">Any single character except the characters in the range inside the brackets</span></span>|[!x-z] |<span data-ttu-id="eeadf-155">t[!a-m]ck 找到 tock 和 tuck，而不是 tack 或 tick。</span><span class="sxs-lookup"><span data-stu-id="eeadf-155">t[!a-m]ck finds tock and tuck, but not tack or tick.</span></span>|
|<span data-ttu-id="eeadf-156">前一个字符或表达式出现 n 次</span><span class="sxs-lookup"><span data-stu-id="eeadf-156">Exactly n occurrences of the previous character or expression</span></span>|<span data-ttu-id="eeadf-157">{n}</span><span class="sxs-lookup"><span data-stu-id="eeadf-157">{n}</span></span> |<span data-ttu-id="eeadf-158">fe{2}d 找到 feed，而不是 fed。</span><span class="sxs-lookup"><span data-stu-id="eeadf-158">fe{2}d finds feed but not fed.</span></span>|
|<span data-ttu-id="eeadf-159">前一个字符或表达式至少出现 n 次</span><span class="sxs-lookup"><span data-stu-id="eeadf-159">At least n occurrences of the previous character or expression</span></span>|<span data-ttu-id="eeadf-160">{n,}</span><span class="sxs-lookup"><span data-stu-id="eeadf-160">{n,}</span></span> |<span data-ttu-id="eeadf-161">fe{1,}d 找到 fed 和 feed。</span><span class="sxs-lookup"><span data-stu-id="eeadf-161">fe{1,}d finds fed and feed.</span></span>|
|<span data-ttu-id="eeadf-162">前一个字符或表达式出现 n 到 m 次</span><span class="sxs-lookup"><span data-stu-id="eeadf-162">From n to m occurrences of the previous character or expression</span></span>|<span data-ttu-id="eeadf-163">{n,m}</span><span class="sxs-lookup"><span data-stu-id="eeadf-163">{n,m}</span></span> |<span data-ttu-id="eeadf-164">10{1,3} 找到 10、100 和 1000。</span><span class="sxs-lookup"><span data-stu-id="eeadf-164">10{1,3} finds 10, 100, and 1000.</span></span>|
|<span data-ttu-id="eeadf-165">前一个字符或表达式出现一次或多次</span><span class="sxs-lookup"><span data-stu-id="eeadf-165">One or more occurrences of the previous character or expression</span></span>|@ |<span data-ttu-id="eeadf-166">lo@t 找到 lot 和 loot。</span><span class="sxs-lookup"><span data-stu-id="eeadf-166">lo@t finds lot and loot.</span></span>|

### <a name="escaping-the-special-characters"></a><span data-ttu-id="eeadf-167">转义特殊字符</span><span class="sxs-lookup"><span data-stu-id="eeadf-167">Escaping the special characters</span></span>

<span data-ttu-id="eeadf-p110">通配符搜索与正则表达式搜索大致相同。正则表达式中有特殊字符，包括“[”、“]”、“(”、“)”、“{”、“}”、“\*”、“?”、“<”、“>”、“!”和“@”。如果其中一个字符属于代码要搜索的文本字符串，则需要转义这个字符，以便让 Word 知道应该以文本形式（而不是作为正则表达式逻辑的一部分）处理这个字符。若要在 Word  UI 搜索中转义字符，请在字符前面添加“\'”字符。不过，若要以编程方式转义，请将字符置于“[]”字符之间。例如，“[\*]\*”搜索以“\*”开头、后跟任意数量的其他字符的所有字符串。</span><span class="sxs-lookup"><span data-stu-id="eeadf-p110">Wildcard search is essentially the same as searching on a regular expression. There are special characters in regular expressions, including '[', ']', '(', ')', '{', '}', '\*', '?', '<', '>', '!', and '@'. If one of these characters is part of the literal string the code is searching for, then it needs to be escaped, so that Word knows it should be treated literally and not as part of the logic of the regular expression. To escape a character in the Word UI search, you would precede it with a '\' character, but to escape it programmatically, put it between '[]' characters. For example, '[\*]\*' searches for any string that begins with a '\*' followed by any number of other characters.</span></span> 

## <a name="examples"></a><span data-ttu-id="eeadf-173">示例</span><span class="sxs-lookup"><span data-stu-id="eeadf-173">Examples</span></span>

<span data-ttu-id="eeadf-174">下面示例演示常见情况。</span><span class="sxs-lookup"><span data-stu-id="eeadf-174">The following examples demonstrate common scenarios.</span></span>

### <a name="ignore-punctuation-search"></a><span data-ttu-id="eeadf-175">忽略标点符号搜索</span><span class="sxs-lookup"><span data-stu-id="eeadf-175">Ignore punctuation search</span></span>

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

### <a name="search-based-on-a-prefix"></a><span data-ttu-id="eeadf-176">基于前缀搜索</span><span class="sxs-lookup"><span data-stu-id="eeadf-176">Search based on a prefix</span></span>

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

### <a name="search-based-on-a-suffix"></a><span data-ttu-id="eeadf-177">基于后缀搜索</span><span class="sxs-lookup"><span data-stu-id="eeadf-177">Search based on a suffix</span></span>

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

### <a name="search-using-a-wildcard"></a><span data-ttu-id="eeadf-178">使用通配符搜索</span><span class="sxs-lookup"><span data-stu-id="eeadf-178">Search using a wildcard</span></span>

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

<span data-ttu-id="eeadf-179">更多信息请参阅 [Word JavaScript API 参考](/office/dev/add-ins/reference/overview/word-add-ins-reference-overview).</span><span class="sxs-lookup"><span data-stu-id="eeadf-179">More information can be found in the [Word JavaScript Reference API](/office/dev/add-ins/reference/overview/word-add-ins-reference-overview).</span></span>
