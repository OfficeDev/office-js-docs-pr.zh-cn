---
title: 使用 Excel JavaScript API 处理注释
description: 有关使用 Api 添加、删除和编辑注释和注释线程的信息。
ms.date: 10/09/2020
localization_priority: Normal
ms.openlocfilehash: 00f7dd22fb2148902152197521098482071e5284
ms.sourcegitcommit: 4e7c74ad67ea8bf6b47d65b2fde54a967090f65b
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/20/2020
ms.locfileid: "48626419"
---
# <a name="work-with-comments-using-the-excel-javascript-api"></a><span data-ttu-id="65a95-103">使用 Excel JavaScript API 处理注释</span><span class="sxs-lookup"><span data-stu-id="65a95-103">Work with comments using the Excel JavaScript API</span></span>

<span data-ttu-id="65a95-104">本文介绍如何使用 Excel JavaScript API 在工作簿中添加、读取、修改和删除注释。</span><span class="sxs-lookup"><span data-stu-id="65a95-104">This article describes how to add, read, modify, and remove comments in a workbook with the Excel JavaScript API.</span></span> <span data-ttu-id="65a95-105">您可以从 Excel 文章的 " [插入注释和注释](https://support.office.com/article/insert-comments-and-notes-in-excel-bdcc9f5d-38e2-45b4-9a92-0b2b5c7bf6f8) " 中了解有关注释功能的详细信息。</span><span class="sxs-lookup"><span data-stu-id="65a95-105">You can learn more about the comment feature from the [Insert comments and notes in Excel](https://support.office.com/article/insert-comments-and-notes-in-excel-bdcc9f5d-38e2-45b4-9a92-0b2b5c7bf6f8) article.</span></span>

<span data-ttu-id="65a95-106">在 Excel JavaScript API 中，注释包括单个初始注释和已连接的线程讨论。</span><span class="sxs-lookup"><span data-stu-id="65a95-106">In the Excel JavaScript API, a comment includes both the single initial comment and the connected threaded discussion.</span></span> <span data-ttu-id="65a95-107">它与单个单元格相关联。</span><span class="sxs-lookup"><span data-stu-id="65a95-107">It is tied to an individual cell.</span></span> <span data-ttu-id="65a95-108">任何人查看具有足够权限的工作簿都可以答复注释。</span><span class="sxs-lookup"><span data-stu-id="65a95-108">Anyone viewing the workbook with sufficient permissions can reply to a comment.</span></span> <span data-ttu-id="65a95-109">[Comment](/javascript/api/excel/excel.comment)对象将那些答复存储为[CommentReply](/javascript/api/excel/excel.commentreply)对象。</span><span class="sxs-lookup"><span data-stu-id="65a95-109">A [Comment](/javascript/api/excel/excel.comment) object stores those replies as [CommentReply](/javascript/api/excel/excel.commentreply) objects.</span></span> <span data-ttu-id="65a95-110">应将注释视为线程，并且线程必须具有特殊条目作为起始点。</span><span class="sxs-lookup"><span data-stu-id="65a95-110">You should consider a comment to be a thread and that a thread must have a special entry as the starting point.</span></span>

![带有两个答复的标签为 "Comment" 的 Excel 注释，标记为 "Comment. 答复 [0]" 和 "Comment. 答复 [1]"。](../images/excel-comments.png)

<span data-ttu-id="65a95-112">工作簿中的注释由属性跟踪 `Workbook.comments` 。</span><span class="sxs-lookup"><span data-stu-id="65a95-112">Comments within a workbook are tracked by the `Workbook.comments` property.</span></span> <span data-ttu-id="65a95-113">这包括由用户创建的批注以及由加载项创建的批注。</span><span class="sxs-lookup"><span data-stu-id="65a95-113">This includes comments created by users and also comments created by your add-in.</span></span> <span data-ttu-id="65a95-114">`Workbook.comments` 属性是一个包含一系列 [Comment](/javascript/api/excel/excel.comment) 对象的 [CommentCollection](/javascript/api/excel/excel.commentcollection) 对象。</span><span class="sxs-lookup"><span data-stu-id="65a95-114">The `Workbook.comments` property is a [CommentCollection](/javascript/api/excel/excel.commentcollection) object that contains a collection of [Comment](/javascript/api/excel/excel.comment) objects.</span></span> <span data-ttu-id="65a95-115">此外，还可以在 [工作表](/javascript/api/excel/excel.worksheet) 级别访问注释。</span><span class="sxs-lookup"><span data-stu-id="65a95-115">Comments are also accessible at the [Worksheet](/javascript/api/excel/excel.worksheet) level.</span></span> <span data-ttu-id="65a95-116">本文中的示例处理工作簿级别的注释，但可以轻松地将其修改为使用 `Worksheet.comments` 属性。</span><span class="sxs-lookup"><span data-stu-id="65a95-116">The samples in this article work with comments at the workbook level, but they can be easily modified to use the `Worksheet.comments` property.</span></span>

## <a name="add-comments"></a><span data-ttu-id="65a95-117">添加备注</span><span class="sxs-lookup"><span data-stu-id="65a95-117">Add comments</span></span>

<span data-ttu-id="65a95-118">使用 `CommentCollection.add` 方法将注释添加到工作簿中。</span><span class="sxs-lookup"><span data-stu-id="65a95-118">Use the `CommentCollection.add` method to add comments to a workbook.</span></span> <span data-ttu-id="65a95-119">此方法最长可使用三个参数：</span><span class="sxs-lookup"><span data-stu-id="65a95-119">This method takes up to three parameters:</span></span>

- <span data-ttu-id="65a95-120">`cellAddress`：添加了注释的单元格。</span><span class="sxs-lookup"><span data-stu-id="65a95-120">`cellAddress`: The cell where the comment is added.</span></span> <span data-ttu-id="65a95-121">它可以是一个字符串或 [Range](/javascript/api/excel/excel.range) 对象。</span><span class="sxs-lookup"><span data-stu-id="65a95-121">This can either be a string or [Range](/javascript/api/excel/excel.range) object.</span></span> <span data-ttu-id="65a95-122">区域必须是单个单元格。</span><span class="sxs-lookup"><span data-stu-id="65a95-122">The range must be a single cell.</span></span>
- <span data-ttu-id="65a95-123">`content`：注释的内容。</span><span class="sxs-lookup"><span data-stu-id="65a95-123">`content`: The comment's content.</span></span> <span data-ttu-id="65a95-124">将字符串用于纯文本注释。</span><span class="sxs-lookup"><span data-stu-id="65a95-124">Use a string for plain text comments.</span></span> <span data-ttu-id="65a95-125">将 [CommentRichContent](/javascript/api/excel/excel.commentrichcontent) 对象用于包含 [提及](#mentions)的注释。</span><span class="sxs-lookup"><span data-stu-id="65a95-125">Use a [CommentRichContent](/javascript/api/excel/excel.commentrichcontent) object for comments with [mentions](#mentions).</span></span>
- <span data-ttu-id="65a95-126">`contentType`：用于指定内容类型的 [ContentType](/javascript/api/excel/excel.contenttype) 枚举。</span><span class="sxs-lookup"><span data-stu-id="65a95-126">`contentType`: A [ContentType](/javascript/api/excel/excel.contenttype) enum specifying type of content.</span></span> <span data-ttu-id="65a95-127">默认值为 `ContentType.plain`。</span><span class="sxs-lookup"><span data-stu-id="65a95-127">The default value is `ContentType.plain`.</span></span>

<span data-ttu-id="65a95-128">下面的代码示例将向单元格 **A2** 添加批注。</span><span class="sxs-lookup"><span data-stu-id="65a95-128">The following code sample adds a comment to cell **A2**.</span></span>

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
> <span data-ttu-id="65a95-129">外接程序添加的注释被应用于该外接程序的当前用户。</span><span class="sxs-lookup"><span data-stu-id="65a95-129">Comments added by an add-in are attributed to the current user of that add-in.</span></span>

### <a name="add-comment-replies"></a><span data-ttu-id="65a95-130">添加批注答复</span><span class="sxs-lookup"><span data-stu-id="65a95-130">Add comment replies</span></span>

<span data-ttu-id="65a95-131">`Comment`对象是包含零个或多个答复的注释线程。</span><span class="sxs-lookup"><span data-stu-id="65a95-131">A `Comment` object is a comment thread that contains zero or more replies.</span></span> <span data-ttu-id="65a95-132">`Comment` 对象具有 `replies` 属性，后者是一个包含 [CommentReply](/javascript/api/excel/excel.commentreply) 对象的 [CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection) 对象。</span><span class="sxs-lookup"><span data-stu-id="65a95-132">`Comment` objects have a `replies` property, which is a [CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection) that contains [CommentReply](/javascript/api/excel/excel.commentreply) objects.</span></span> <span data-ttu-id="65a95-133">若要向批注添加回复，请使用 `CommentReplyCollection.add` 方法，传入回复的文本。</span><span class="sxs-lookup"><span data-stu-id="65a95-133">To add a reply to a comment, use the `CommentReplyCollection.add` method, passing in the text of the reply.</span></span> <span data-ttu-id="65a95-134">回复将按照添加的顺序显示。</span><span class="sxs-lookup"><span data-stu-id="65a95-134">Replies are displayed in the order they are added.</span></span> <span data-ttu-id="65a95-135">此外接加载项的当前用户也具有这些属性。</span><span class="sxs-lookup"><span data-stu-id="65a95-135">They are also attributed to the current user of the add-in.</span></span>

<span data-ttu-id="65a95-136">下面的代码示例向工作簿中的第一个批注添加回复。</span><span class="sxs-lookup"><span data-stu-id="65a95-136">The following code sample adds a reply to the first comment in the workbook.</span></span>

```js
Excel.run(function (context) {
    // Get the first comment added to the workbook.
    var comment = context.workbook.comments.getItemAt(0);
    comment.replies.add("Thanks for the reminder!");
    return context.sync();
});
```

## <a name="edit-comments"></a><span data-ttu-id="65a95-137">编辑批注</span><span class="sxs-lookup"><span data-stu-id="65a95-137">Edit comments</span></span>

<span data-ttu-id="65a95-138">若要编辑批注或批注回复，请设置其 `Comment.content` 属性或 `CommentReply.content` 属性。</span><span class="sxs-lookup"><span data-stu-id="65a95-138">To edit a comment or comment reply, set its `Comment.content` property or `CommentReply.content` property.</span></span>

```js
Excel.run(function (context) {
    // Edit the first comment in the workbook.
    var comment = context.workbook.comments.getItemAt(0);
    comment.content = "PLEASE add headers here.";
    return context.sync();
});
```

### <a name="edit-comment-replies"></a><span data-ttu-id="65a95-139">编辑批注答复</span><span class="sxs-lookup"><span data-stu-id="65a95-139">Edit comment replies</span></span>

<span data-ttu-id="65a95-140">若要编辑批注答复，请设置其 `CommentReply.content` 属性。</span><span class="sxs-lookup"><span data-stu-id="65a95-140">To edit a comment reply, set its `CommentReply.content` property.</span></span>

```js
Excel.run(function (context) {
    // Edit the first comment reply on the first comment in the workbook.
    var comment = context.workbook.comments.getItemAt(0);
    var reply = comment.replies.getItemAt(0);
    reply.content = "Never mind";
    return context.sync();
});
```

## <a name="delete-comments"></a><span data-ttu-id="65a95-141">删除注释</span><span class="sxs-lookup"><span data-stu-id="65a95-141">Delete comments</span></span>

<span data-ttu-id="65a95-142">若要删除注释，请使用 `Comment.delete` 方法。</span><span class="sxs-lookup"><span data-stu-id="65a95-142">To delete a comment use the `Comment.delete` method.</span></span> <span data-ttu-id="65a95-143">删除注释的同时也会删除与该注释相关的答复。</span><span class="sxs-lookup"><span data-stu-id="65a95-143">Deleting a comment also deletes the replies associated with that comment.</span></span>

```js
Excel.run(function (context) {
    // Delete the comment thread at A2 on the "MyWorksheet" worksheet.
    context.workbook.comments.getItemByCell("MyWorksheet!A2").delete();
    return context.sync();
});
```

### <a name="delete-comment-replies"></a><span data-ttu-id="65a95-144">删除批注答复</span><span class="sxs-lookup"><span data-stu-id="65a95-144">Delete comment replies</span></span>

<span data-ttu-id="65a95-145">若要删除批注答复，请使用 `CommentReply.delete` 方法。</span><span class="sxs-lookup"><span data-stu-id="65a95-145">To delete a comment reply, use the `CommentReply.delete` method.</span></span>

```js
Excel.run(function (context) {
    // Delete the first comment reply from this worksheet's first comment.
    var comment = context.workbook.comments.getItemAt(0);
    comment.replies.getItemAt(0).delete();
    return context.sync();
});
```

## <a name="resolve-comment-threads"></a><span data-ttu-id="65a95-146">解析注释线程</span><span class="sxs-lookup"><span data-stu-id="65a95-146">Resolve comment threads</span></span>

<span data-ttu-id="65a95-147">注释线程具有可配置的布尔值， `resolved` 以指示是否已解决。</span><span class="sxs-lookup"><span data-stu-id="65a95-147">A comment thread has a configurable boolean value, `resolved`, to indicate if it is resolved.</span></span> <span data-ttu-id="65a95-148">值 `true` 表示注释线程已解析。</span><span class="sxs-lookup"><span data-stu-id="65a95-148">A value of `true` means the comment thread is resolved.</span></span> <span data-ttu-id="65a95-149">值 `false` 表示注释线程是新的，也可能是重新打开的。</span><span class="sxs-lookup"><span data-stu-id="65a95-149">A value of `false` means the comment thread is either new or reopened.</span></span>

```js
Excel.run(function (context) {
    // Resolve the first comment thread in the workbook.
    context.workbook.comments.getItemAt(0).resolved = true;
    return context.sync();
});
```

<span data-ttu-id="65a95-150">批注答复有一个 readonly `resolved` 属性。</span><span class="sxs-lookup"><span data-stu-id="65a95-150">Comment replies have a readonly `resolved` property.</span></span> <span data-ttu-id="65a95-151">它的值始终等于线程的其余部分的值。</span><span class="sxs-lookup"><span data-stu-id="65a95-151">Its value is always equal to that of the rest of the thread.</span></span>

## <a name="comment-metadata"></a><span data-ttu-id="65a95-152">注释元数据</span><span class="sxs-lookup"><span data-stu-id="65a95-152">Comment metadata</span></span>

<span data-ttu-id="65a95-153">每个批注都包含有关其创建情况的元数据，如作者和创建日期。</span><span class="sxs-lookup"><span data-stu-id="65a95-153">Each comment contains metadata about its creation, such as the author and creation date.</span></span> <span data-ttu-id="65a95-154">由加载项创建的批注将被视为是由当前用户创作的。</span><span class="sxs-lookup"><span data-stu-id="65a95-154">Comments created by your add-in are considered to be authored by the current user.</span></span>

<span data-ttu-id="65a95-155">下面的示例演示如何显示 **A2** 中批注的作者电子邮件、作者姓名和创建日期。</span><span class="sxs-lookup"><span data-stu-id="65a95-155">The following sample shows how to display the author's email, author's name, and creation date of a comment at **A2**.</span></span>

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

### <a name="comment-reply-metadata"></a><span data-ttu-id="65a95-156">批注答复元数据</span><span class="sxs-lookup"><span data-stu-id="65a95-156">Comment reply metadata</span></span>

<span data-ttu-id="65a95-157">批注答复存储与初始注释相同类型的元数据。</span><span class="sxs-lookup"><span data-stu-id="65a95-157">Comment replies store the same types of metadata as the initial comment.</span></span>

<span data-ttu-id="65a95-158">下面的示例展示了如何在 **A2**中显示作者的电子邮件、作者的姓名以及最新注释答复的创建日期。</span><span class="sxs-lookup"><span data-stu-id="65a95-158">The following sample shows how to display the author's email, author's name, and creation date of the latest comment reply at **A2**.</span></span>

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

## <a name="mentions"></a><span data-ttu-id="65a95-159">提及</span><span class="sxs-lookup"><span data-stu-id="65a95-159">Mentions</span></span>

<span data-ttu-id="65a95-160">[提及](https://support.office.com/article/use-mention-in-comments-to-tag-someone-for-feedback-644bf689-31a0-4977-a4fb-afe01820c1fd) 用于在注释中标记同事。</span><span class="sxs-lookup"><span data-stu-id="65a95-160">[Mentions](https://support.office.com/article/use-mention-in-comments-to-tag-someone-for-feedback-644bf689-31a0-4977-a4fb-afe01820c1fd) are used to tag colleagues in a comment.</span></span> <span data-ttu-id="65a95-161">这将向他们发送你的评论内容通知。</span><span class="sxs-lookup"><span data-stu-id="65a95-161">This sends them notifications with your comment's content.</span></span> <span data-ttu-id="65a95-162">你的外接程序可以代表你创建这些提及。</span><span class="sxs-lookup"><span data-stu-id="65a95-162">Your add-in can create these mentions on your behalf.</span></span>

<span data-ttu-id="65a95-163">包含提及的注释需要使用 [CommentRichContent](/javascript/api/excel/excel.commentrichcontent) 对象创建。</span><span class="sxs-lookup"><span data-stu-id="65a95-163">Comments with mentions need to be created with [CommentRichContent](/javascript/api/excel/excel.commentrichcontent) objects.</span></span> <span data-ttu-id="65a95-164">`CommentCollection.add`使用 `CommentRichContent` 包含一个或多个提及的调用，并将其指定 `ContentType.mention` 为 `contentType` 参数。</span><span class="sxs-lookup"><span data-stu-id="65a95-164">Call `CommentCollection.add` with a `CommentRichContent` containing one or more mentions and specify `ContentType.mention` as the `contentType` parameter.</span></span> <span data-ttu-id="65a95-165">`content`此外，还需要设置字符串格式，以在文本中插入所提及的内容。</span><span class="sxs-lookup"><span data-stu-id="65a95-165">The `content` string also needs to be formatted to insert the mention into the text.</span></span> <span data-ttu-id="65a95-166">提及的格式为： `<at id="{replyIndex}">{mentionName}</at>` 。</span><span class="sxs-lookup"><span data-stu-id="65a95-166">The format for a mention is: `<at id="{replyIndex}">{mentionName}</at>`.</span></span>

> [!NOTE]
> <span data-ttu-id="65a95-167">目前，只有提及的确切名称可用作提及链接的文本。</span><span class="sxs-lookup"><span data-stu-id="65a95-167">Currently, only the mention's exact name can be used as the text of the mention link.</span></span> <span data-ttu-id="65a95-168">稍后将添加对名称的缩写版本的支持。</span><span class="sxs-lookup"><span data-stu-id="65a95-168">Support for shortened versions of a name will be added later.</span></span>

<span data-ttu-id="65a95-169">下面的示例展示了一个注释，其中包含一个注明。</span><span class="sxs-lookup"><span data-stu-id="65a95-169">The following example shows a comment with a single mention.</span></span>

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

## <a name="comment-events"></a><span data-ttu-id="65a95-170">注释事件</span><span class="sxs-lookup"><span data-stu-id="65a95-170">Comment events</span></span>

<span data-ttu-id="65a95-171">您的外接程序可以侦听注释的添加、更改和删除。</span><span class="sxs-lookup"><span data-stu-id="65a95-171">Your add-in can listen for comment additions, changes, and deletions.</span></span> <span data-ttu-id="65a95-172">[批注事件](/javascript/api/excel/excel.commentcollection#event-details) 发生在 `CommentCollection` 对象上。</span><span class="sxs-lookup"><span data-stu-id="65a95-172">[Comment events](/javascript/api/excel/excel.commentcollection#event-details) occur on the `CommentCollection` object.</span></span> <span data-ttu-id="65a95-173">若要侦听注释事件，请注册 `onAdded` 、 `onChanged` 或 `onDeleted` 注释事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="65a95-173">To listen for comment events, register the `onAdded`, `onChanged`, or `onDeleted` comment event handler.</span></span> <span data-ttu-id="65a95-174">当检测到注释事件时，请使用此事件处理程序检索有关添加的、已更改或已删除的注释的数据。</span><span class="sxs-lookup"><span data-stu-id="65a95-174">When a comment event is detected, use this event handler to retrieve data about the added, changed, or deleted comment.</span></span> <span data-ttu-id="65a95-175">该 `onChanged` 事件还处理注释添加、更改和删除。</span><span class="sxs-lookup"><span data-stu-id="65a95-175">The `onChanged` event also handles comment reply additions, changes, and deletions.</span></span> 

<span data-ttu-id="65a95-176">每个注释事件仅在同时执行多个添加、更改或删除时触发一次。</span><span class="sxs-lookup"><span data-stu-id="65a95-176">Each comment event only triggers once when multiple additions, changes, or deletions are performed at the same time.</span></span> <span data-ttu-id="65a95-177">所有 [CommentAddedEventArgs](/javascript/api/excel/excel.commentaddedeventargs)、 [CommentChangedEventArgs](/javascript/api/excel/excel.commentchangedeventargs)和 [CommentDeletedEventArgs](/javascript/api/excel/excel.commentdeletedeventargs) 对象都包含注释 id 的数组，用于将事件操作映射回注释集合。</span><span class="sxs-lookup"><span data-stu-id="65a95-177">All the [CommentAddedEventArgs](/javascript/api/excel/excel.commentaddedeventargs), [CommentChangedEventArgs](/javascript/api/excel/excel.commentchangedeventargs), and [CommentDeletedEventArgs](/javascript/api/excel/excel.commentdeletedeventargs) objects contain arrays of comment IDs to map the event actions back to the comment collections.</span></span>

<span data-ttu-id="65a95-178">若要详细了解如何注册事件处理程序、处理事件和删除事件处理程序，请参阅使用 [Excel JAVASCRIPT API 文章处理事件](excel-add-ins-events.md) 。</span><span class="sxs-lookup"><span data-stu-id="65a95-178">See the [Work with Events using the Excel JavaScript API](excel-add-ins-events.md) article for additional information about registering event handlers, handling events, and removing event handlers.</span></span> 

### <a name="comment-addition-events"></a><span data-ttu-id="65a95-179">注释添加事件</span><span class="sxs-lookup"><span data-stu-id="65a95-179">Comment addition events</span></span> 
<span data-ttu-id="65a95-180">向 `onAdded` 注释集合中添加一个或多个新注释时，将触发该事件。</span><span class="sxs-lookup"><span data-stu-id="65a95-180">The `onAdded` event is triggered when one or more new comments are added to the comment collection.</span></span> <span data-ttu-id="65a95-181">将答复添加到注释线程中时， *不* 会触发此事件 (请参阅 [注释更改事件](#comment-change-events) 以了解有关注释答复事件) 。</span><span class="sxs-lookup"><span data-stu-id="65a95-181">This event is *not* triggered when replies are added to a comment thread (see [Comment change events](#comment-change-events) to learn about comment reply events).</span></span>

<span data-ttu-id="65a95-182">下面的示例展示了如何注册 `onAdded` 事件处理程序，然后使用该 `CommentAddedEventArgs` 对象来检索 `commentDetails` 添加的注释的数组。</span><span class="sxs-lookup"><span data-stu-id="65a95-182">The following sample shows how to register the `onAdded` event handler and then use the `CommentAddedEventArgs` object to retrieve the `commentDetails` array of the added comment.</span></span>

> [!NOTE]
> <span data-ttu-id="65a95-183">此示例仅在添加单个批注时才起作用。</span><span class="sxs-lookup"><span data-stu-id="65a95-183">This sample only works when a single comment is added.</span></span> 

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

### <a name="comment-change-events"></a><span data-ttu-id="65a95-184">批注更改事件</span><span class="sxs-lookup"><span data-stu-id="65a95-184">Comment change events</span></span> 
<span data-ttu-id="65a95-185">`onChanged`在下列情况下，会触发注释事件。</span><span class="sxs-lookup"><span data-stu-id="65a95-185">The `onChanged` comment event is triggered in the following scenarios.</span></span>

- <span data-ttu-id="65a95-186">更新注释的内容。</span><span class="sxs-lookup"><span data-stu-id="65a95-186">A comment's content is updated.</span></span>
- <span data-ttu-id="65a95-187">解析注释线程。</span><span class="sxs-lookup"><span data-stu-id="65a95-187">A comment thread is resolved.</span></span>
- <span data-ttu-id="65a95-188">重新打开注释线程。</span><span class="sxs-lookup"><span data-stu-id="65a95-188">A comment thread is reopened.</span></span>
- <span data-ttu-id="65a95-189">将答复添加到注释线程中。</span><span class="sxs-lookup"><span data-stu-id="65a95-189">A reply is added to a comment thread.</span></span>
- <span data-ttu-id="65a95-190">在注释线程中更新答复。</span><span class="sxs-lookup"><span data-stu-id="65a95-190">A reply is updated in a comment thread.</span></span>
- <span data-ttu-id="65a95-191">在注释线程中删除答复。</span><span class="sxs-lookup"><span data-stu-id="65a95-191">A reply is deleted in a comment thread.</span></span>

<span data-ttu-id="65a95-192">下面的示例展示了如何注册 `onChanged` 事件处理程序，然后使用该 `CommentChangedEventArgs` 对象来检索 `commentDetails` 已更改注释的数组。</span><span class="sxs-lookup"><span data-stu-id="65a95-192">The following sample shows how to register the `onChanged` event handler and then use the `CommentChangedEventArgs` object to retrieve the `commentDetails` array of the changed comment.</span></span>

> [!NOTE]
> <span data-ttu-id="65a95-193">此示例仅在更改单个批注时才起作用。</span><span class="sxs-lookup"><span data-stu-id="65a95-193">This sample only works when a single comment is changed.</span></span> 

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

### <a name="comment-deletion-events"></a><span data-ttu-id="65a95-194">注释删除事件</span><span class="sxs-lookup"><span data-stu-id="65a95-194">Comment deletion events</span></span>
<span data-ttu-id="65a95-195">`onDeleted`从注释集合中删除注释时将触发该事件。</span><span class="sxs-lookup"><span data-stu-id="65a95-195">The `onDeleted` event is triggered when a comment is deleted from the comment collection.</span></span> <span data-ttu-id="65a95-196">删除注释后，其元数据将不再可用。</span><span class="sxs-lookup"><span data-stu-id="65a95-196">Once a comment has been deleted, its metadata is no longer available.</span></span> <span data-ttu-id="65a95-197">如果外接程序管理各个注释，则 [CommentDeletedEventArgs](/javascript/api/excel/excel.commentdeletedeventargs) 对象提供注释 id。</span><span class="sxs-lookup"><span data-stu-id="65a95-197">The [CommentDeletedEventArgs](/javascript/api/excel/excel.commentdeletedeventargs) object provides comment IDs, in case your add-in is managing individual comments.</span></span>

<span data-ttu-id="65a95-198">下面的示例展示了如何注册 `onDeleted` 事件处理程序，然后使用该 `CommentDeletedEventArgs` 对象来检索 `commentDetails` 已删除注释的数组。</span><span class="sxs-lookup"><span data-stu-id="65a95-198">The following sample shows how to register the `onDeleted` event handler and then use the `CommentDeletedEventArgs` object to retrieve the `commentDetails` array of the deleted comment.</span></span>

> [!NOTE]
> <span data-ttu-id="65a95-199">此示例仅在删除单个批注时才起作用。</span><span class="sxs-lookup"><span data-stu-id="65a95-199">This sample only works when a single comment is deleted.</span></span> 

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

## <a name="see-also"></a><span data-ttu-id="65a95-200">另请参阅</span><span class="sxs-lookup"><span data-stu-id="65a95-200">See also</span></span>

- [<span data-ttu-id="65a95-201">Office 外接程序中的 Excel JavaScript 对象模型</span><span class="sxs-lookup"><span data-stu-id="65a95-201">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="65a95-202">使用 Excel JavaScript API 处理工作簿</span><span class="sxs-lookup"><span data-stu-id="65a95-202">Work with workbooks using the Excel JavaScript API</span></span>](excel-add-ins-workbooks.md)
- [<span data-ttu-id="65a95-203">使用 Excel JavaScript API 处理事件</span><span class="sxs-lookup"><span data-stu-id="65a95-203">Work with Events using the Excel JavaScript API</span></span>](excel-add-ins-events.md)
- [<span data-ttu-id="65a95-204">在 Excel 中插入批注和备注</span><span class="sxs-lookup"><span data-stu-id="65a95-204">Insert comments and notes in Excel</span></span>](https://support.office.com/article/insert-comments-and-notes-in-excel-bdcc9f5d-38e2-45b4-9a92-0b2b5c7bf6f8)
