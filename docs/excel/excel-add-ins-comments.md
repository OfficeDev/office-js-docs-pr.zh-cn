---
title: 使用 JavaScript API Excel注释
description: 有关使用 API 添加、删除和编辑注释和注释线程的信息。
ms.date: 10/09/2020
localization_priority: Normal
ms.openlocfilehash: 5e292dab77b080906d77b1517a8de715bc0d2122f29e3de73b04f5b9d9276c85
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/07/2021
ms.locfileid: "57084317"
---
# <a name="work-with-comments-using-the-excel-javascript-api"></a>使用 JavaScript API Excel注释

本文介绍如何使用 JavaScript API 在工作簿中添加、读取、修改Excel注释。 可以从在文档中插入注释和注释一文了解有关[Excel功能。](https://support.office.com/article/insert-comments-and-notes-in-excel-bdcc9f5d-38e2-45b4-9a92-0b2b5c7bf6f8)

在 Excel JavaScript API 中，注释包括单个初始注释和已连接的线程讨论。 它绑定到单个单元格。 查看具有足够权限的工作簿的任何人都可以回复注释。 Comment [对象](/javascript/api/excel/excel.comment) 将回复存储为 [CommentReply](/javascript/api/excel/excel.commentreply) 对象。 你应该将注释视为一个线程，并且线程必须具有一个特殊条目作为起点。

![一Excel批注，标记为"Comment"，带有两个回复，标记为"Comment.replies[0]"和"Comment.replies[1]。](../images/excel-comments.png)

工作簿中的注释由 属性 `Workbook.comments` 进行跟踪。 这包括由用户创建的批注以及由加载项创建的批注。 `Workbook.comments` 属性是一个包含一系列 [Comment](/javascript/api/excel/excel.comment) 对象的 [CommentCollection](/javascript/api/excel/excel.commentcollection) 对象。 注释也可在 [工作表级别访问](/javascript/api/excel/excel.worksheet) 。 本文中的示例处理工作簿级别的注释，但可以轻松地修改它们以使用 `Worksheet.comments` 属性。

## <a name="add-comments"></a>添加备注

使用 `CommentCollection.add` 方法向工作簿添加注释。 此方法最多需要三个参数：

- `cellAddress`：添加注释的单元格。 它可以是字符串或 [Range](/javascript/api/excel/excel.range) 对象。 区域必须是单个单元格。
- `content`：注释的内容。 将字符串用于纯文本注释。 将 [CommentRichContent](/javascript/api/excel/excel.commentrichcontent) 对象用于包含提及 [的评论](#mentions)。
- `contentType`：指定内容类型的 [ContentType](/javascript/api/excel/excel.contenttype) 枚举。 默认值为 `ContentType.plain`。

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
> 由加载项添加的注释将归结到加载项的当前用户。

### <a name="add-comment-replies"></a>添加批注回复

对象 `Comment` 是包含零个或多个回复的注释线程。 `Comment` 对象具有 `replies` 属性，后者是一个包含 [CommentReply](/javascript/api/excel/excel.commentreply) 对象的 [CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection) 对象。 若要向批注添加回复，请使用 `CommentReplyCollection.add` 方法，传入回复的文本。 回复将按照添加的顺序显示。 它们还会归结到加载项的当前用户。

下面的代码示例向工作簿中的第一个批注添加回复。

```js
Excel.run(function (context) {
    // Get the first comment added to the workbook.
    var comment = context.workbook.comments.getItemAt(0);
    comment.replies.add("Thanks for the reminder!");
    return context.sync();
});
```

## <a name="edit-comments"></a>编辑注释

若要编辑批注或批注回复，请设置其 `Comment.content` 属性或 `CommentReply.content` 属性。

```js
Excel.run(function (context) {
    // Edit the first comment in the workbook.
    var comment = context.workbook.comments.getItemAt(0);
    comment.content = "PLEASE add headers here.";
    return context.sync();
});
```

### <a name="edit-comment-replies"></a>编辑批注回复

若要编辑批注回复，请设置其 `CommentReply.content` 属性。

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

若要删除批注，请使用 `Comment.delete` 方法。 删除注释还会删除与该注释关联的回复。

```js
Excel.run(function (context) {
    // Delete the comment thread at A2 on the "MyWorksheet" worksheet.
    context.workbook.comments.getItemByCell("MyWorksheet!A2").delete();
    return context.sync();
});
```

### <a name="delete-comment-replies"></a>删除批注回复

若要删除批注回复，请使用 `CommentReply.delete` 方法。

```js
Excel.run(function (context) {
    // Delete the first comment reply from this worksheet's first comment.
    var comment = context.workbook.comments.getItemAt(0);
    comment.replies.getItemAt(0).delete();
    return context.sync();
});
```

## <a name="resolve-comment-threads"></a>解析注释线程

注释线程具有可配置的布尔值 `resolved` ，以指示是否解析。 的值 `true` 表示注释线程已解析。 的值 `false` 表示注释线程是新的或重新打开的。

```js
Excel.run(function (context) {
    // Resolve the first comment thread in the workbook.
    context.workbook.comments.getItemAt(0).resolved = true;
    return context.sync();
});
```

批注回复具有 readonly `resolved` 属性。 它的值始终等于线程其余部分的值。

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

### <a name="comment-reply-metadata"></a>注释回复元数据

注释回复存储与初始注释相同的元数据类型。

以下示例演示如何在 **A2** 上显示作者的电子邮件、作者姓名和最新批注回复的创建日期。

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

## <a name="mentions"></a>提及

[提及](https://support.office.com/article/use-mention-in-comments-to-tag-someone-for-feedback-644bf689-31a0-4977-a4fb-afe01820c1fd) 用于在注释中标记同事。 这会向用户发送包含注释内容的通知。 加载项可以代表你创建这些提及内容。

需要用 [CommentRichContent](/javascript/api/excel/excel.commentrichcontent) 对象创建包含提及内容的评论。 使用 `CommentCollection.add` 包含 `CommentRichContent` 一个或多个提及项的 调用，并 `ContentType.mention` 指定 作为 `contentType` 参数。 `content`还需要设置字符串的格式，以在文本中插入提及内容。 提及的格式为 `<at id="{replyIndex}">{mentionName}</at>` ：。

> [!NOTE]
> 目前，仅提及的确切名称可以用作提及链接的文本。 稍后将添加对名称的缩短版本的支持。

以下示例显示一个提及评论。

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

## <a name="comment-events"></a>注释事件

加载项可以侦听注释添加、更改和删除。 [注释事件](/javascript/api/excel/excel.commentcollection#event-details) 在对象 `CommentCollection` 上发生。 若要侦听注释事件，请注册 `onAdded` 、 `onChanged` 或 `onDeleted` comment 事件处理程序。 检测到注释事件时，使用此事件处理程序检索有关已添加、已更改或删除的注释的数据。 `onChanged`该事件还会处理批注回复的添加、更改和删除。 

当同时执行多个添加、更改或删除操作时，每个注释事件仅触发一次。 所有[CommentAddedEventArgs、CommentChangedEventArgs](/javascript/api/excel/excel.commentaddedeventargs)和[CommentDeletedEventArgs](/javascript/api/excel/excel.commentdeletedeventargs)对象都包含注释 ID 数组，用于将事件操作映射回注释集合。 [](/javascript/api/excel/excel.commentchangedeventargs)

有关注册事件处理程序、处理事件和删除事件处理程序的其他信息，请参阅使用[Excel JavaScript API](excel-add-ins-events.md)处理事件一文。 

### <a name="comment-addition-events"></a>注释添加事件 
向 `onAdded` 注释集合中添加一个或多个新批注时，将触发该事件。 将回复 *添加到* 注释线程时，不会触发此事件 (请参阅注释更改事件以了解注释回复事件) 。 [](#comment-change-events)

以下示例演示如何注册事件 `onAdded` 处理程序，然后使用 `CommentAddedEventArgs` 对象检索添加的注释 `commentDetails` 的数组。

> [!NOTE]
> 本示例仅在添加单个批注时有效。 

```js
Excel.run(function (context) {
    var comments = context.workbook.worksheets.getActiveWorksheet().comments;

    // Register the onAdded comment event handler.
    comments.onAdded.add(commentAdded);

    return context.sync();
});

function commentAdded() {
    Excel.run(function (context) {
        // Retrieve the added comment using the comment ID.
        // Note: This method assumes only a single comment is added at a time. 
        var addedComment = context.workbook.comments.getItem(event.commentDetails[0].commentId);

        // Load the added comment's data.
        addedComment.load(["content", "authorName"]);

        return context.sync().then(function () {
            // Print out the added comment's data.
            console.log(`A comment was added. ID: ${event.commentDetails[0].commentId}. Comment content:${addedComment.content}. Comment author:${addedComment.authorName}`);
            return context.sync();
        });            
    });
}
```

### <a name="comment-change-events"></a>注释更改事件 
注释 `onChanged` 事件在下列方案中触发。

- 注释的内容已更新。
- 注释线程已解析。
- 将重新打开注释线程。
- 回复将添加到注释线程。
- 回复在注释线程中更新。
- 在注释线程中删除回复。

以下示例演示如何注册事件 `onChanged` 处理程序，然后使用 `CommentChangedEventArgs` 对象检索已更改注释 `commentDetails` 的数组。

> [!NOTE]
> 本示例仅在更改单个批注时有效。 

```js
Excel.run(function (context) {
    var comments = context.workbook.worksheets.getActiveWorksheet().comments;

    // Register the onChanged comment event handler.
    comments.onChanged.add(commentChanged);

    return context.sync();
});    

function commentChanged() {
    Excel.run(function (context) {
        // Retrieve the changed comment using the comment ID.
        // Note: This method assumes only a single comment is changed at a time. 
        var changedComment = context.workbook.comments.getItem(event.commentDetails[0].commentId);

        // Load the changed comment's data.
        changedComment.load(["content", "authorName"]);

        return context.sync().then(function () {
            // Print out the changed comment's data.
            console.log(`A comment was changed. ID: ${event.commentDetails[0].commentId}`. Updated comment content: ${changedComment.content}`. Comment author: ${changedComment.authorName}`);
            return context.sync();
        });
    });
}
```

### <a name="comment-deletion-events"></a>注释删除事件
从 `onDeleted` 注释集合中删除批注时，将触发该事件。 删除注释后，其元数据将不再可用。 [CommentDeletedEventArgs](/javascript/api/excel/excel.commentdeletedeventargs)对象提供注释 ID，以防加载项管理单个注释。

以下示例演示如何注册事件 `onDeleted` 处理程序，然后使用 `CommentDeletedEventArgs` 对象检索已删除 `commentDetails` 注释的数组。

> [!NOTE]
> 本示例仅在删除单个批注时有效。 

```js
Excel.run(function (context) {
    var comments = context.workbook.worksheets.getActiveWorksheet().comments;

    // Register the onDeleted comment event handler.
    comments.onDeleted.add(commentDeleted);

    return context.sync();
});

function commentDeleted() {
    Excel.run(function (context) {
        // Print out the deleted comment's ID.
        // Note: This method assumes only a single comment is deleted at a time. 
        console.log(`A comment was deleted. ID: ${event.commentDetails[0].commentId}`);
    });
}
```

## <a name="see-also"></a>另请参阅

- [Excel 加载项中的 Word JavaScript 对象模型](excel-add-ins-core-concepts.md)
- [使用 Excel JavaScript API 处理工作簿](excel-add-ins-workbooks.md)
- [使用 Excel JavaScript API 处理事件](excel-add-ins-events.md)
- [在文档中插入注释Excel](https://support.office.com/article/insert-comments-and-notes-in-excel-bdcc9f5d-38e2-45b4-9a92-0b2b5c7bf6f8)
