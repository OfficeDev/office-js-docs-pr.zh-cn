---
title: 使用 Excel JavaScript API 处理注释
description: 有关使用 Api 添加、删除和编辑注释和注释线程的信息。
ms.date: 03/17/2020
localization_priority: Normal
ms.openlocfilehash: 275828915730d3438101315ee28bf76aa8b8bf3f
ms.sourcegitcommit: 6c381634c77d316f34747131860db0a0bced2529
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/21/2020
ms.locfileid: "42890568"
---
# <a name="work-with-comments-using-the-excel-javascript-api"></a>使用 Excel JavaScript API 处理注释

本文介绍如何使用 Excel JavaScript API 在工作簿中添加、读取、修改和删除注释。 您可以从 Excel 文章的 "[插入注释和注释](https://support.office.com/article/insert-comments-and-notes-in-excel-bdcc9f5d-38e2-45b4-9a92-0b2b5c7bf6f8)" 中了解有关注释功能的详细信息。

在 Excel JavaScript API 中，注释包括单个初始注释和已连接的线程讨论。 它与单个单元格相关联。 任何人查看具有足够权限的工作簿都可以答复注释。 [Comment](/javascript/api/excel/excel.comment)对象将那些答复存储为[CommentReply](/javascript/api/excel/excel.commentreply)对象。 应将注释视为线程，并且线程必须具有特殊条目作为起始点。

![带有两个答复的标签为 "Comment" 的 Excel 注释，标记为 "Comment. 答复 [0]" 和 "Comment. 答复 [1]"。](../images/excel-comments.png)

工作簿中的注释由`Workbook.comments`属性跟踪。 这包括由用户创建的批注以及由加载项创建的批注。 `Workbook.comments` 属性是一个包含一系列 [Comment](/javascript/api/excel/excel.comment) 对象的 [CommentCollection](/javascript/api/excel/excel.commentcollection) 对象。 此外，还可以在[工作表](/javascript/api/excel/excel.worksheet)级别访问注释。 本文中的示例处理工作簿级别的注释，但可以轻松地将其修改为使用`Worksheet.comments`属性。

## <a name="add-comments"></a>添加备注

使用`CommentCollection.add`方法将注释添加到工作簿中。 此方法最长可使用三个参数：

- `cellAddress`：添加了注释的单元格。 它可以是一个字符串或[Range](/javascript/api/excel/excel.range)对象。 区域必须是单个单元格。
- `content`：注释的内容。 将字符串用于纯文本注释。 将[CommentRichContent](/javascript/api/excel/excel.commentrichcontent)对象用于包含[提及](#mentions-online-only)的注释。 
- `contentType`：用于指定内容类型的[ContentType](/javascript/api/excel/excel.contenttype)枚举。 默认值为 `ContentType.plain`。

下面的代码示例将向单元格 **A2** 添加批注。

```js
Excel.run(function (context) {
    // Add a comment to A2 on the "MyWorksheet" worksheet.
    var comments = context.workbook.comments;

    // Note that an InvalidArgument error will be thrown if multiple cells passed to `Comment.add`.
    comments.add("MyWorksheet!A2", "TODO: add data.");
    return context.sync();
});
```

> [!NOTE]
> 外接程序添加的注释被应用于该外接程序的当前用户。

### <a name="add-comment-replies"></a>添加批注答复

`Comment`对象是包含零个或多个答复的注释线程。 `Comment` 对象具有 `replies` 属性，后者是一个包含 [CommentReply](/javascript/api/excel/excel.commentreply) 对象的 [CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection) 对象。 若要向批注添加回复，请使用 `CommentReplyCollection.add` 方法，传入回复的文本。 回复将按照添加的顺序显示。 此外接加载项的当前用户也具有这些属性。

下面的代码示例向工作簿中的第一个批注添加回复。

```js
Excel.run(function (context) {
    // Get the first comment added to the workbook.
    var comment = context.workbook.comments.getItemAt(0);
    comment.replies.add("Thanks for the reminder!");
    return context.sync();
});
```

## <a name="edit-comments"></a>编辑批注

若要编辑批注或批注回复，请设置其 `Comment.content` 属性或 `CommentReply.content` 属性。

```js
Excel.run(function (context) {
    // Edit the first comment in the workbook.
    var comment = context.workbook.comments.getItemAt(0);
    comment.content = "PLEASE add headers here.";
    return context.sync();
});
```

### <a name="edit-comment-replies"></a>编辑批注答复

若要编辑批注答复，请设置`CommentReply.content`其属性。

```js
Excel.run(function (context) {
    // Edit the first comment reply on the first comment in the workbook.
    var comment = context.workbook.comments.getItemAt(0);
    var reply = comment.replies.getItemAt(0);
    reply.content = "Never mind";
    return context.sync();
});
```

## <a name="delete-comments"></a>删除注释

若要删除注释，请`Comment.delete`使用方法。 删除注释的同时也会删除与该注释相关的答复。

```js
Excel.run(function (context) {
    // Delete the comment thread at A2 on the "MyWorksheet" worksheet.
    context.workbook.comments.getItemByCell("MyWorksheet!A2").delete();
    return context.sync();
});
```

### <a name="delete-comment-replies"></a>删除批注答复

若要删除批注答复，请使用`CommentReply.delete`方法。

```js
Excel.run(function (context) {
    // Delete the first comment reply from this worksheet's first comment.
    var comment = context.workbook.comments.getItemAt(0);
    comment.replies.getItemAt(0).delete();
    return context.sync();
});
```

## <a name="resolve-comment-threads-preview"></a>解析注释线索（[预览](../reference/requirement-sets/excel-preview-apis.md)） 

注释线程具有可配置的布尔值， `resolved`以指示是否已解决。 值`true`表示注释线程已解析。 值`false`表示注释线程是新的，也可能是重新打开的。

```js
Excel.run(function (context) {
    // Resolve the first comment thread in the workbook.
    context.workbook.comments.getItemAt(0).resolved = true;
    return context.sync();
});
```

批注答复有一个 readonly `resolved`属性。 它的值始终等于线程的其余部分的值。

## <a name="comment-metadata"></a>注释元数据

每个批注都包含有关其创建情况的元数据，如作者和创建日期。 由加载项创建的批注将被视为是由当前用户创作的。

下面的示例演示如何显示 **A2** 中批注的作者电子邮件、作者姓名和创建日期。

```js
Excel.run(function (context) {
    // Get the comment at cell A2 in the "MyWorksheet" worksheet.
    var comment = context.workbook.comments.getItemByCell("MyWorksheet!A2");

    // Load and print the following values.
    comment.load(["authorEmail", "authorName", "creationDate"]);
    return context.sync().then(function () {
        console.log(`${comment.creationDate.toDateString()}: ${comment.authorName} (${comment.authorEmail})`);
    });
});
```

### <a name="comment-reply-metadata"></a>批注答复元数据

批注答复存储与初始注释相同类型的元数据。

下面的示例展示了如何在**A2**中显示作者的电子邮件、作者的姓名以及最新注释答复的创建日期。

```js
Excel.run(function (context) {
    // Get the comment at cell A2 in the "MyWorksheet" worksheet.
    var comment = context.workbook.comments.getItemByCell("MyWorksheet!A2");
    var replyCount = comment.replies.getCount();
    // Sync to get the current number of comment replies.
    return context.sync().then(function () {
        // Get the last comment reply in the comment thread.
        var reply = comment.replies.getItemAt(replyCount.value - 1);
        reply.load(["authorEmail", "authorName", "creationDate"]);
        // Sync to load the reply metadata to print.
        return context.sync().then(function () {
            console.log(`Latest reply: ${reply.creationDate.toDateString()}: ${reply.authorName} ${reply.authorEmail})`);
            return context.sync();
        });
    });
});
```

## <a name="mentions-online-only"></a>提及（[仅联机](../reference/requirement-sets/excel-api-online-requirement-set.md)） 

> [!NOTE]
> 注释提到的 Api 当前仅适用于公共预览版。 [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]

> [!IMPORTANT]
> 目前仅支持对 web 上的 Excel 进行注释提及。

[提及](https://support.office.com/article/use-mention-in-comments-to-tag-someone-for-feedback-644bf689-31a0-4977-a4fb-afe01820c1fd)用于在注释中标记同事。 这将向他们发送你的评论内容通知。 你的外接程序可以代表你创建这些提及。

包含提及的注释需要使用[CommentRichContent](/javascript/api/excel/excel.commentrichcontent)对象创建。 使用`CommentCollection.add`包含一个`CommentRichContent`或多个提及的调用， `ContentType.mention`并将`contentType`其指定为参数。 此外`content` ，还需要设置字符串格式，以在文本中插入所提及的内容。 提及的格式为： `<at id="{replyIndex}">{mentionName}</at>`。

> 便笺目前，只有提及的确切名称可用作提及链接的文本。 稍后将添加对名称的缩写版本的支持。

下面的示例展示了一个注释，其中包含一个注明。

```js
Excel.run(function (context) {
    // Add an "@mention" for "Kate Kristensen" to cell A1 in the "MyWorksheet" worksheet.
    var mention = {
        email: "kakri@contoso.com",
        id: 0,
        name: "Kate Kristensen"
    };

    // This will tag the mention's name using the '@' syntax.
    // They will be notified via email.
    var commentBody = {
        mentions: [mention],
        richContent: '<at id="0">' + mention.name + "</at> -  Can you take a look?"
    };

    // Note that an InvalidArgument error will be thrown if multiple cells passed to `comment.add`.
    context.workbook.comments.add("MyWorksheet!A1", commentBody, Excel.ContentType.mention);
    return context.sync();
});
```

## <a name="see-also"></a>另请参阅

- [Excel JavaScript API 基本编程概念](excel-add-ins-core-concepts.md)
- [使用 Excel JavaScript API 处理工作簿](excel-add-ins-workbooks.md)
- [在 Excel 中插入批注和备注](https://support.office.com/article/insert-comments-and-notes-in-excel-bdcc9f5d-38e2-45b4-9a92-0b2b5c7bf6f8)
