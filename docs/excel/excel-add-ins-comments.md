---
title: 使用 Excel JavaScript API 处理注释
description: 有关使用 API 添加、删除和编辑批注和注释线程的信息。
ms.date: 02/15/2022
ms.localizationpriority: medium
ms.openlocfilehash: 5996c1bb55c3d4a358786b15f7c3e46aae6f42aa
ms.sourcegitcommit: eef2064d7966db91f8401372dd255a32d76168c2
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/31/2022
ms.locfileid: "67464795"
---
# <a name="work-with-comments-using-the-excel-javascript-api"></a>使用 Excel JavaScript API 处理注释

本文介绍如何使用 Excel JavaScript API 在工作簿中添加、读取、修改和删除注释。 可以从 [Excel 文章中的“插入”注释和备注中](https://support.microsoft.com/office/bdcc9f5d-38e2-45b4-9a92-0b2b5c7bf6f8) 了解有关注释功能的详细信息。

在 Excel JavaScript API 中，注释包括单个初始注释和连接的线程讨论。 它绑定到单个单元格。 查看具有足够权限的工作簿的任何人都可以回复批注。 [Comment](/javascript/api/excel/excel.comment) 对象将这些答复存储为 [CommentReply](/javascript/api/excel/excel.commentreply) 对象。 应将注释视为线程，并且线程必须具有特殊条目作为起点。

![一个 Excel 注释，标有“Comment”和两个答复，标签为“Comment.replies[0]”和“Comment.replies[1]”。](../images/excel-comments.png)

该属性跟踪 `Workbook.comments` 工作簿中的注释。 这包括由用户创建的批注以及由加载项创建的批注。 `Workbook.comments` 属性是一个包含一系列 [Comment](/javascript/api/excel/excel.comment) 对象的 [CommentCollection](/javascript/api/excel/excel.commentcollection) 对象。 也可以在 [工作表](/javascript/api/excel/excel.worksheet) 级别访问注释。 本文中的示例使用工作簿级别的注释，但可以轻松修改这些示例以使用该 `Worksheet.comments` 属性。

## <a name="add-comments"></a>添加备注

使用该 `CommentCollection.add` 方法将注释添加到工作簿。 此方法最多采用三个参数：

- `cellAddress`：添加批注的单元格。 这可以是字符串或 [Range](/javascript/api/excel/excel.range) 对象。 该区域必须是单个单元格。
- `content`：注释的内容。 对纯文本注释使用字符串。 使用 [CommentRichContent](/javascript/api/excel/excel.commentrichcontent) 对象获取 [带提及的注释](#mentions)。
- `contentType`：指定内容类型的 [ContentType](/javascript/api/excel/excel.contenttype) 枚举。 默认值为 `ContentType.plain`。

下面的代码示例将向单元格 **A2** 添加批注。

```js
await Excel.run(async (context) => {
    // Add a comment to A2 on the "MyWorksheet" worksheet.
    let comments = context.workbook.comments;

    // Note that an InvalidArgument error will be thrown if multiple cells passed to `Comment.add`.
    comments.add("MyWorksheet!A2", "TODO: add data.");
    await context.sync();
});
```

> [!NOTE]
> 加载项添加的注释将归因于该加载项的当前用户。

### <a name="add-comment-replies"></a>添加批注回复

`Comment`对象是包含零个或多个答复的注释线程。 `Comment` 对象具有 `replies` 属性，后者是一个包含 [CommentReply](/javascript/api/excel/excel.commentreply) 对象的 [CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection) 对象。 若要向批注添加回复，请使用 `CommentReplyCollection.add` 方法，传入回复的文本。 回复将按照添加的顺序显示。 它们还归于加载项的当前用户。

下面的代码示例向工作簿中的第一个批注添加回复。

```js
await Excel.run(async (context) => {
    // Get the first comment added to the workbook.
    let comment = context.workbook.comments.getItemAt(0);
    comment.replies.add("Thanks for the reminder!");
    await context.sync();
});
```

## <a name="edit-comments"></a>编辑注释

若要编辑批注或批注回复，请设置其 `Comment.content` 属性或 `CommentReply.content` 属性。

```js
await Excel.run(async (context) => {
    // Edit the first comment in the workbook.
    let comment = context.workbook.comments.getItemAt(0);
    comment.content = "PLEASE add headers here.";
    await context.sync();
});
```

### <a name="edit-comment-replies"></a>编辑批注回复

若要编辑批注回复，请设置其 `CommentReply.content` 属性。

```js
await Excel.run(async (context) => {
    // Edit the first comment reply on the first comment in the workbook.
    let comment = context.workbook.comments.getItemAt(0);
    let reply = comment.replies.getItemAt(0);
    reply.content = "Never mind";
    await context.sync();
});
```

## <a name="delete-comments"></a>删除注释

若要删除注释，请使用该 `Comment.delete` 方法。 删除注释还会删除与该注释关联的答复。

```js
await Excel.run(async (context) => {
    // Delete the comment thread at A2 on the "MyWorksheet" worksheet.
    context.workbook.comments.getItemByCell("MyWorksheet!A2").delete();
    await context.sync();
});
```

### <a name="delete-comment-replies"></a>删除批注回复

若要删除批注回复，请使用该 `CommentReply.delete` 方法。

```js
await Excel.run(async (context) => {
    // Delete the first comment reply from this worksheet's first comment.
    let comment = context.workbook.comments.getItemAt(0);
    comment.replies.getItemAt(0).delete();
    await context.sync();
});
```

## <a name="resolve-comment-threads"></a>解析注释线程

注释线程具有可配置的布尔值， `resolved`用于指示是否已解析。 表示解析注释线程的 `true` 值。 表示注释线程为新线程或重新打开的值 `false` 。

```js
await Excel.run(async (context) => {
    // Resolve the first comment thread in the workbook.
    context.workbook.comments.getItemAt(0).resolved = true;
    await context.sync();
});
```

批注答复具有只 `resolved` 读属性。 其值始终等于线程其余部分的值。

## <a name="comment-metadata"></a>注释元数据

每个批注都包含有关其创建情况的元数据，如作者和创建日期。 由加载项创建的批注将被视为是由当前用户创作的。

下面的示例演示如何显示 **A2** 中批注的作者电子邮件、作者姓名和创建日期。

```js
await Excel.run(async (context) => {
    // Get the comment at cell A2 in the "MyWorksheet" worksheet.
    let comment = context.workbook.comments.getItemByCell("MyWorksheet!A2");

    // Load and print the following values.
    comment.load(["authorEmail", "authorName", "creationDate"]);
    await context.sync();
    
    console.log(`${comment.creationDate.toDateString()}: ${comment.authorName} (${comment.authorEmail})`);
});
```

### <a name="comment-reply-metadata"></a>注释回复元数据

注释回复存储与初始注释相同的元数据类型。

以下示例演示如何在 **A2** 上显示作者的电子邮件、作者的姓名和最新批注回复的创建日期。

```js
await Excel.run(async (context) => {
    // Get the comment at cell A2 in the "MyWorksheet" worksheet.
    let comment = context.workbook.comments.getItemByCell("MyWorksheet!A2");
    let replyCount = comment.replies.getCount();
    // Sync to get the current number of comment replies.
    await context.sync();

    // Get the last comment reply in the comment thread.
    let reply = comment.replies.getItemAt(replyCount.value - 1);
    reply.load(["authorEmail", "authorName", "creationDate"]);

    // Sync to load the reply metadata to print.
    await context.sync();

    console.log(`Latest reply: ${reply.creationDate.toDateString()}: ${reply.authorName} ${reply.authorEmail})`);
    await context.sync();
});
```

## <a name="mentions"></a>提及

[提及](https://support.microsoft.com/office/644bf689-31a0-4977-a4fb-afe01820c1fd) 用于标记批注中的同事。 这会向他们发送包含批注内容的通知。 加载项可以代表你创建这些提及。

需要使用 [CommentRichContent](/javascript/api/excel/excel.commentrichcontent) 对象创建带提及的注释。 使用包含一个`CommentRichContent`或多个提及的调用`CommentCollection.add`并指定`ContentType.mention`为`contentType`参数。 还需要 `content` 设置字符串格式才能将提及内容插入到文本中。 提及的格式为： `<at id="{replyIndex}">{mentionName}</at>`.

> [!NOTE]
> 目前，只能将提及的确切名称用作提及链接的文本。 稍后将添加对缩短的名称版本的支持。

以下示例显示一个注释，其中一次提及。

```js
await Excel.run(async (context) => {
    // Add an "@mention" for "Kate Kristensen" to cell A1 in the "MyWorksheet" worksheet.
    let mention = {
        email: "kakri@contoso.com",
        id: 0,
        name: "Kate Kristensen"
    };

    // This will tag the mention's name using the '@' syntax.
    // They will be notified via email.
    let commentBody = {
        mentions: [mention],
        richContent: '<at id="0">' + mention.name + "</at> -  Can you take a look?"
    };

    // Note that an InvalidArgument error will be thrown if multiple cells passed to `comment.add`.
    context.workbook.comments.add("MyWorksheet!A1", commentBody, Excel.ContentType.mention);
    await context.sync();
});
```

## <a name="comment-events"></a>注释事件

外接程序可以侦听注释添加、更改和删除。 在对象上`CommentCollection`发生[注释事件](/javascript/api/excel/excel.commentcollection#event-details)。 若要侦听批注事件，请注册`onAdded``onChanged`或`onDeleted`注释事件处理程序。 检测到注释事件时，请使用此事件处理程序检索有关添加、更改或删除的注释的数据。 该 `onChanged` 事件还处理批注回复的添加、更改和删除。

当同时执行多个添加、更改或删除时，每个注释事件仅触发一次。 所有 [CommentAddedEventArgs](/javascript/api/excel/excel.commentaddedeventargs)、 [CommentChangedEventArgs](/javascript/api/excel/excel.commentchangedeventargs) 和 [CommentDeletedEventArgs](/javascript/api/excel/excel.commentdeletedeventargs) 对象都包含批注 ID 数组，用于将事件操作映射回注释集合。

有关注册事件处理程序、处理事件和删除事件处理程序的其他信息，请参阅 [使用 Excel JavaScript API](excel-add-ins-events.md) 文章处理事件。

### <a name="comment-addition-events"></a>注释添加事件

将 `onAdded` 一个或多个新注释添加到注释集合时，将触发该事件。 将回复添加到批注线程 (查看 [注释更改事件](#comment-change-events)以了解批注回复事件) 时，*不会* 触发此事件。

以下示例演示如何注册 `onAdded` 事件处理程序，然后使用该 `CommentAddedEventArgs` 对象检索 `commentDetails` 添加的注释的数组。

> [!NOTE]
> 仅当添加单个注释时，此示例才有效。

```js
await Excel.run(async (context) => {
    let comments = context.workbook.worksheets.getActiveWorksheet().comments;

    // Register the onAdded comment event handler.
    comments.onAdded.add(commentAdded);

    await context.sync();
});

async function commentAdded() {
    await Excel.run(async (context) => {
        // Retrieve the added comment using the comment ID.
        // Note: This method assumes only a single comment is added at a time. 
        let addedComment = context.workbook.comments.getItem(event.commentDetails[0].commentId);

        // Load the added comment's data.
        addedComment.load(["content", "authorName"]);

        await context.sync();

        // Print out the added comment's data.
        console.log(`A comment was added. ID: ${event.commentDetails[0].commentId}. Comment content:${addedComment.content}. Comment author:${addedComment.authorName}`);
        await context.sync();
    });
}
```

### <a name="comment-change-events"></a>注释更改事件

批 `onChanged` 注事件在以下情况下触发。

- 注释的内容已更新。
- 注释线程已解析。
- 注释线程将重新打开。
- 回复将添加到批注线程。
- 回复会在注释线程中更新。
- 在注释线程中删除答复。

以下示例演示如何注册 `onChanged` 事件处理程序，然后使用该 `CommentChangedEventArgs` 对象检索 `commentDetails` 已更改批注的数组。

> [!NOTE]
> 此示例仅在更改单个注释时才有效。

```js
await Excel.run(async (context) => {
    let comments = context.workbook.worksheets.getActiveWorksheet().comments;

    // Register the onChanged comment event handler.
    comments.onChanged.add(commentChanged);

    await context.sync();
});

async function commentChanged() {
    await Excel.run(async (context) => {
        // Retrieve the changed comment using the comment ID.
        // Note: This method assumes only a single comment is changed at a time. 
        let changedComment = context.workbook.comments.getItem(event.commentDetails[0].commentId);

        // Load the changed comment's data.
        changedComment.load(["content", "authorName"]);

        await context.sync();

        // Print out the changed comment's data.
        console.log(`A comment was changed. ID: ${event.commentDetails[0].commentId}. Updated comment content: ${changedComment.content}. Comment author: ${changedComment.authorName}`);
        await context.sync();
    });
}
```

### <a name="comment-deletion-events"></a>注释删除事件

`onDeleted`从注释集合中删除注释时，将触发该事件。 删除批注后，其元数据将不再可用。 [CommentDeletedEventArgs](/javascript/api/excel/excel.commentdeletedeventargs) 对象提供注释 ID，以防加载项管理单个注释。

以下示例演示如何注册 `onDeleted` 事件处理程序，然后使用该 `CommentDeletedEventArgs` 对象检索 `commentDetails` 已删除注释的数组。

> [!NOTE]
> 此示例仅在删除单个注释时有效。

```js
await Excel.run(async (context) => {
    let comments = context.workbook.worksheets.getActiveWorksheet().comments;

    // Register the onDeleted comment event handler.
    comments.onDeleted.add(commentDeleted);

    await context.sync();
});

async function commentDeleted() {
    await Excel.run(async (context) => {
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
- [在 Excel 中插入注释和备注](https://support.microsoft.com/office/bdcc9f5d-38e2-45b4-9a92-0b2b5c7bf6f8)
