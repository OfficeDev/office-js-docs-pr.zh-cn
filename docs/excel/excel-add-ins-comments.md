---
title: 使用 Excel JavaScript API 处理注释
description: 有关使用 Api 添加、删除和编辑注释和注释线程的信息。
ms.date: 10/09/2020
localization_priority: Normal
ms.openlocfilehash: 85312cbd92aa6c9d0f82fd167e8a372c2eff8c85
ms.sourcegitcommit: b50eebd303adcc22eb86e65756ce7e9a82f41a57
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/14/2020
ms.locfileid: "48456550"
---
# <a name="work-with-comments-using-the-excel-javascript-api"></a>使用 Excel JavaScript API 处理注释

本文介绍如何使用 Excel JavaScript API 在工作簿中添加、读取、修改和删除注释。 您可以从 Excel 文章的 " [插入注释和注释](https://support.office.com/article/insert-comments-and-notes-in-excel-bdcc9f5d-38e2-45b4-9a92-0b2b5c7bf6f8) " 中了解有关注释功能的详细信息。

在 Excel JavaScript API 中，注释包括单个初始注释和已连接的线程讨论。 它与单个单元格相关联。 任何人查看具有足够权限的工作簿都可以答复注释。 [Comment](/javascript/api/excel/excel.comment)对象将那些答复存储为[CommentReply](/javascript/api/excel/excel.commentreply)对象。 应将注释视为线程，并且线程必须具有特殊条目作为起始点。

![带有两个答复的标签为 "Comment" 的 Excel 注释，标记为 "Comment. 答复 [0]" 和 "Comment. 答复 [1]"。](../images/excel-comments.png)

工作簿中的注释由属性跟踪 `Workbook.comments` 。 这包括由用户创建的批注以及由加载项创建的批注。 `Workbook.comments` 属性是一个包含一系列 [Comment](/javascript/api/excel/excel.comment) 对象的 [CommentCollection](/javascript/api/excel/excel.commentcollection) 对象。 此外，还可以在 [工作表](/javascript/api/excel/excel.worksheet) 级别访问注释。 本文中的示例处理工作簿级别的注释，但可以轻松地将其修改为使用 `Worksheet.comments` 属性。

## <a name="add-comments"></a>添加备注

使用 `CommentCollection.add` 方法将注释添加到工作簿中。 此方法最长可使用三个参数：

- `cellAddress`：添加了注释的单元格。 它可以是一个字符串或 [Range](/javascript/api/excel/excel.range) 对象。 区域必须是单个单元格。
- `content`：注释的内容。 将字符串用于纯文本注释。 将 [CommentRichContent](/javascript/api/excel/excel.commentrichcontent) 对象用于包含 [提及](#mentions)的注释。
- `contentType`：用于指定内容类型的 [ContentType](/javascript/api/excel/excel.contenttype) 枚举。 默认值为 `ContentType.plain`。

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

若要编辑批注答复，请设置其 `CommentReply.content` 属性。

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

若要删除注释，请使用 `Comment.delete` 方法。 删除注释的同时也会删除与该注释相关的答复。

```js
Excel.run(function (context) {
    // Delete the comment thread at A2 on the "MyWorksheet" worksheet.
    context.workbook.comments.getItemByCell("MyWorksheet!A2").delete();
    return context.sync();
});
```

### <a name="delete-comment-replies"></a>删除批注答复

若要删除批注答复，请使用 `CommentReply.delete` 方法。

```js
Excel.run(function (context) {
    // Delete the first comment reply from this worksheet's first comment.
    var comment = context.workbook.comments.getItemAt(0);
    comment.replies.getItemAt(0).delete();
    return context.sync();
});
```

## <a name="resolve-comment-threads"></a>解析注释线程

注释线程具有可配置的布尔值， `resolved` 以指示是否已解决。 值 `true` 表示注释线程已解析。 值 `false` 表示注释线程是新的，也可能是重新打开的。

```js
Excel.run(function (context) {
    // Resolve the first comment thread in the workbook.
    context.workbook.comments.getItemAt(0).resolved = true;
    return context.sync();
});
```

批注答复有一个 readonly `resolved` 属性。 它的值始终等于线程的其余部分的值。

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

下面的示例展示了如何在 **A2**中显示作者的电子邮件、作者的姓名以及最新注释答复的创建日期。

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

[提及](https://support.office.com/article/use-mention-in-comments-to-tag-someone-for-feedback-644bf689-31a0-4977-a4fb-afe01820c1fd) 用于在注释中标记同事。 这将向他们发送你的评论内容通知。 你的外接程序可以代表你创建这些提及。

包含提及的注释需要使用 [CommentRichContent](/javascript/api/excel/excel.commentrichcontent) 对象创建。 `CommentCollection.add`使用 `CommentRichContent` 包含一个或多个提及的调用，并将其指定 `ContentType.mention` 为 `contentType` 参数。 `content`此外，还需要设置字符串格式，以在文本中插入所提及的内容。 提及的格式为： `<at id="{replyIndex}">{mentionName}</at>` 。

> [!NOTE]
> 目前，只有提及的确切名称可用作提及链接的文本。 稍后将添加对名称的缩写版本的支持。

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

## <a name="comment-events"></a>注释事件

您的外接程序可以侦听注释的添加、更改和删除。 [批注事件](/javascript/api/excel/excel.commentcollection#event-details) 发生在 `CommentCollection` 对象上。 若要侦听注释事件，请注册 `onAdded` 、 `onChanged` 或 `onDeleted` 注释事件处理程序。 当检测到注释事件时，请使用此事件处理程序检索有关添加的、已更改或已删除的注释的数据。 该 `onChanged` 事件还处理注释添加、更改和删除。 

每个注释事件仅在同时执行多个添加、更改或删除时触发一次。 所有 [CommentAddedEventArgs](/javascript/api/excel/excel.commentaddedeventargs)、 [CommentChangedEventArgs](/javascript/api/excel/excel.commentchangedeventarg)和 [CommentDeletedEventArgs](/javascript/api/excel/excel.commentdeletedeventargs) 对象都包含注释 id 的数组，用于将事件操作映射回注释集合。

若要详细了解如何注册事件处理程序、处理事件和删除事件处理程序，请参阅使用 [Excel JAVASCRIPT API 文章处理事件](excel-add-ins-events.md) 。 

### <a name="comment-addition-events"></a>注释添加事件 
向 `onAdded` 注释集合中添加一个或多个新注释时，将触发该事件。 将答复添加到注释线程中时， *不* 会触发此事件 (请参阅 [注释更改事件](#comment-change-events) 以了解有关注释答复事件) 。

下面的示例展示了如何注册 `onAdded` 事件处理程序，然后使用该 `CommentAddedEventArgs` 对象来检索 `commentDetails` 添加的注释的数组。

> [!NOTE]
> 此示例仅在添加单个批注时才起作用。 

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

### <a name="comment-change-events"></a>批注更改事件 
`onChanged`在下列情况下，会触发注释事件。

- 更新注释的内容。
- 解析注释线程。
- 重新打开注释线程。
- 将答复添加到注释线程中。
- 在注释线程中更新答复。
- 在注释线程中删除答复。

下面的示例展示了如何注册 `onChanged` 事件处理程序，然后使用该 `CommentChangedEventArgs` 对象来检索 `commentDetails` 已更改注释的数组。

> [!NOTE]
> 此示例仅在更改单个批注时才起作用。 

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
`onDeleted`从注释集合中删除注释时将触发该事件。 删除注释后，其元数据将不再可用。 如果外接程序管理各个注释，则 [CommentDeletedEventArgs](/javascript/api/excel/excel.commentdeletedeventargs) 对象提供注释 id。

下面的示例展示了如何注册 `onDeleted` 事件处理程序，然后使用该 `CommentDeletedEventArgs` 对象来检索 `commentDetails` 已删除注释的数组。

> [!NOTE]
> 此示例仅在删除单个批注时才起作用。 

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

- [Office 外接程序中的 Excel JavaScript 对象模型](excel-add-ins-core-concepts.md)
- [使用 Excel JavaScript API 处理工作簿](excel-add-ins-workbooks.md)
- [使用 Excel JavaScript API 处理事件](excel-add-ins-events.md)
- [在 Excel 中插入批注和备注](https://support.office.com/article/insert-comments-and-notes-in-excel-bdcc9f5d-38e2-45b4-9a92-0b2b5c7bf6f8)
