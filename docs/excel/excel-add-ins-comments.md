---
title: 使用 Excel JavaScript API 处理注释
description: 有关使用 Api 添加、删除和编辑注释和注释线程的信息。
ms.date: 03/17/2020
localization_priority: Normal
ms.openlocfilehash: 59751086b645ed56cb2a00bc70678967cae085c5
ms.sourcegitcommit: 472b81642e9eb5fb2a55cd98a7b0826d37eb7f73
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/17/2020
ms.locfileid: "45159428"
---
# <a name="work-with-comments-using-the-excel-javascript-api"></a><span data-ttu-id="bfe84-103">使用 Excel JavaScript API 处理注释</span><span class="sxs-lookup"><span data-stu-id="bfe84-103">Work with comments using the Excel JavaScript API</span></span>

<span data-ttu-id="bfe84-104">本文介绍如何使用 Excel JavaScript API 在工作簿中添加、读取、修改和删除注释。</span><span class="sxs-lookup"><span data-stu-id="bfe84-104">This article describes how to add, read, modify, and remove comments in a workbook with the Excel JavaScript API.</span></span> <span data-ttu-id="bfe84-105">您可以从 Excel 文章的 "[插入注释和注释](https://support.office.com/article/insert-comments-and-notes-in-excel-bdcc9f5d-38e2-45b4-9a92-0b2b5c7bf6f8)" 中了解有关注释功能的详细信息。</span><span class="sxs-lookup"><span data-stu-id="bfe84-105">You can learn more about the comment feature from the [Insert comments and notes in Excel](https://support.office.com/article/insert-comments-and-notes-in-excel-bdcc9f5d-38e2-45b4-9a92-0b2b5c7bf6f8) article.</span></span>

<span data-ttu-id="bfe84-106">在 Excel JavaScript API 中，注释包括单个初始注释和已连接的线程讨论。</span><span class="sxs-lookup"><span data-stu-id="bfe84-106">In the Excel JavaScript API, a comment includes both the single initial comment and the connected threaded discussion.</span></span> <span data-ttu-id="bfe84-107">它与单个单元格相关联。</span><span class="sxs-lookup"><span data-stu-id="bfe84-107">It is tied to an individual cell.</span></span> <span data-ttu-id="bfe84-108">任何人查看具有足够权限的工作簿都可以答复注释。</span><span class="sxs-lookup"><span data-stu-id="bfe84-108">Anyone viewing the workbook with sufficient permissions can reply to a comment.</span></span> <span data-ttu-id="bfe84-109">[Comment](/javascript/api/excel/excel.comment)对象将那些答复存储为[CommentReply](/javascript/api/excel/excel.commentreply)对象。</span><span class="sxs-lookup"><span data-stu-id="bfe84-109">A [Comment](/javascript/api/excel/excel.comment) object stores those replies as [CommentReply](/javascript/api/excel/excel.commentreply) objects.</span></span> <span data-ttu-id="bfe84-110">应将注释视为线程，并且线程必须具有特殊条目作为起始点。</span><span class="sxs-lookup"><span data-stu-id="bfe84-110">You should consider a comment to be a thread and that a thread must have a special entry as the starting point.</span></span>

![带有两个答复的标签为 "Comment" 的 Excel 注释，标记为 "Comment. 答复 [0]" 和 "Comment. 答复 [1]"。](../images/excel-comments.png)

<span data-ttu-id="bfe84-112">工作簿中的注释由属性跟踪 `Workbook.comments` 。</span><span class="sxs-lookup"><span data-stu-id="bfe84-112">Comments within a workbook are tracked by the `Workbook.comments` property.</span></span> <span data-ttu-id="bfe84-113">这包括由用户创建的批注以及由加载项创建的批注。</span><span class="sxs-lookup"><span data-stu-id="bfe84-113">This includes comments created by users and also comments created by your add-in.</span></span> <span data-ttu-id="bfe84-114">`Workbook.comments` 属性是一个包含一系列 [Comment](/javascript/api/excel/excel.comment) 对象的 [CommentCollection](/javascript/api/excel/excel.commentcollection) 对象。</span><span class="sxs-lookup"><span data-stu-id="bfe84-114">The `Workbook.comments` property is a [CommentCollection](/javascript/api/excel/excel.commentcollection) object that contains a collection of [Comment](/javascript/api/excel/excel.comment) objects.</span></span> <span data-ttu-id="bfe84-115">此外，还可以在[工作表](/javascript/api/excel/excel.worksheet)级别访问注释。</span><span class="sxs-lookup"><span data-stu-id="bfe84-115">Comments are also accessible at the [Worksheet](/javascript/api/excel/excel.worksheet) level.</span></span> <span data-ttu-id="bfe84-116">本文中的示例处理工作簿级别的注释，但可以轻松地将其修改为使用 `Worksheet.comments` 属性。</span><span class="sxs-lookup"><span data-stu-id="bfe84-116">The samples in this article work with comments at the workbook level, but they can be easily modified to use the `Worksheet.comments` property.</span></span>

## <a name="add-comments"></a><span data-ttu-id="bfe84-117">添加备注</span><span class="sxs-lookup"><span data-stu-id="bfe84-117">Add comments</span></span>

<span data-ttu-id="bfe84-118">使用 `CommentCollection.add` 方法将注释添加到工作簿中。</span><span class="sxs-lookup"><span data-stu-id="bfe84-118">Use the `CommentCollection.add` method to add comments to a workbook.</span></span> <span data-ttu-id="bfe84-119">此方法最长可使用三个参数：</span><span class="sxs-lookup"><span data-stu-id="bfe84-119">This method takes up to three parameters:</span></span>

- <span data-ttu-id="bfe84-120">`cellAddress`：添加了注释的单元格。</span><span class="sxs-lookup"><span data-stu-id="bfe84-120">`cellAddress`: The cell where the comment is added.</span></span> <span data-ttu-id="bfe84-121">它可以是一个字符串或[Range](/javascript/api/excel/excel.range)对象。</span><span class="sxs-lookup"><span data-stu-id="bfe84-121">This can either be a string or [Range](/javascript/api/excel/excel.range) object.</span></span> <span data-ttu-id="bfe84-122">区域必须是单个单元格。</span><span class="sxs-lookup"><span data-stu-id="bfe84-122">The range must be a single cell.</span></span>
- <span data-ttu-id="bfe84-123">`content`：注释的内容。</span><span class="sxs-lookup"><span data-stu-id="bfe84-123">`content`: The comment's content.</span></span> <span data-ttu-id="bfe84-124">将字符串用于纯文本注释。</span><span class="sxs-lookup"><span data-stu-id="bfe84-124">Use a string for plain text comments.</span></span> <span data-ttu-id="bfe84-125">将[CommentRichContent](/javascript/api/excel/excel.commentrichcontent)对象用于包含[提及](#mentions)的注释。</span><span class="sxs-lookup"><span data-stu-id="bfe84-125">Use a [CommentRichContent](/javascript/api/excel/excel.commentrichcontent) object for comments with [mentions](#mentions).</span></span>
- <span data-ttu-id="bfe84-126">`contentType`：用于指定内容类型的[ContentType](/javascript/api/excel/excel.contenttype)枚举。</span><span class="sxs-lookup"><span data-stu-id="bfe84-126">`contentType`: A [ContentType](/javascript/api/excel/excel.contenttype) enum specifying type of content.</span></span> <span data-ttu-id="bfe84-127">默认值为 `ContentType.plain`。</span><span class="sxs-lookup"><span data-stu-id="bfe84-127">The default value is `ContentType.plain`.</span></span>

<span data-ttu-id="bfe84-128">下面的代码示例将向单元格 **A2** 添加批注。</span><span class="sxs-lookup"><span data-stu-id="bfe84-128">The following code sample adds a comment to cell **A2**.</span></span>

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
> <span data-ttu-id="bfe84-129">外接程序添加的注释被应用于该外接程序的当前用户。</span><span class="sxs-lookup"><span data-stu-id="bfe84-129">Comments added by an add-in are attributed to the current user of that add-in.</span></span>

### <a name="add-comment-replies"></a><span data-ttu-id="bfe84-130">添加批注答复</span><span class="sxs-lookup"><span data-stu-id="bfe84-130">Add comment replies</span></span>

<span data-ttu-id="bfe84-131">`Comment`对象是包含零个或多个答复的注释线程。</span><span class="sxs-lookup"><span data-stu-id="bfe84-131">A `Comment` object is a comment thread that contains zero or more replies.</span></span> <span data-ttu-id="bfe84-132">`Comment` 对象具有 `replies` 属性，后者是一个包含 [CommentReply](/javascript/api/excel/excel.commentreply) 对象的 [CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection) 对象。</span><span class="sxs-lookup"><span data-stu-id="bfe84-132">`Comment` objects have a `replies` property, which is a [CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection) that contains [CommentReply](/javascript/api/excel/excel.commentreply) objects.</span></span> <span data-ttu-id="bfe84-133">若要向批注添加回复，请使用 `CommentReplyCollection.add` 方法，传入回复的文本。</span><span class="sxs-lookup"><span data-stu-id="bfe84-133">To add a reply to a comment, use the `CommentReplyCollection.add` method, passing in the text of the reply.</span></span> <span data-ttu-id="bfe84-134">回复将按照添加的顺序显示。</span><span class="sxs-lookup"><span data-stu-id="bfe84-134">Replies are displayed in the order they are added.</span></span> <span data-ttu-id="bfe84-135">此外接加载项的当前用户也具有这些属性。</span><span class="sxs-lookup"><span data-stu-id="bfe84-135">They are also attributed to the current user of the add-in.</span></span>

<span data-ttu-id="bfe84-136">下面的代码示例向工作簿中的第一个批注添加回复。</span><span class="sxs-lookup"><span data-stu-id="bfe84-136">The following code sample adds a reply to the first comment in the workbook.</span></span>

```js
Excel.run(function (context) {
    // Get the first comment added to the workbook.
    var comment = context.workbook.comments.getItemAt(0);
    comment.replies.add("Thanks for the reminder!");
    return context.sync();
});
```

## <a name="edit-comments"></a><span data-ttu-id="bfe84-137">编辑批注</span><span class="sxs-lookup"><span data-stu-id="bfe84-137">Edit comments</span></span>

<span data-ttu-id="bfe84-138">若要编辑批注或批注回复，请设置其 `Comment.content` 属性或 `CommentReply.content` 属性。</span><span class="sxs-lookup"><span data-stu-id="bfe84-138">To edit a comment or comment reply, set its `Comment.content` property or `CommentReply.content` property.</span></span>

```js
Excel.run(function (context) {
    // Edit the first comment in the workbook.
    var comment = context.workbook.comments.getItemAt(0);
    comment.content = "PLEASE add headers here.";
    return context.sync();
});
```

### <a name="edit-comment-replies"></a><span data-ttu-id="bfe84-139">编辑批注答复</span><span class="sxs-lookup"><span data-stu-id="bfe84-139">Edit comment replies</span></span>

<span data-ttu-id="bfe84-140">若要编辑批注答复，请设置其 `CommentReply.content` 属性。</span><span class="sxs-lookup"><span data-stu-id="bfe84-140">To edit a comment reply, set its `CommentReply.content` property.</span></span>

```js
Excel.run(function (context) {
    // Edit the first comment reply on the first comment in the workbook.
    var comment = context.workbook.comments.getItemAt(0);
    var reply = comment.replies.getItemAt(0);
    reply.content = "Never mind";
    return context.sync();
});
```

## <a name="delete-comments"></a><span data-ttu-id="bfe84-141">删除注释</span><span class="sxs-lookup"><span data-stu-id="bfe84-141">Delete comments</span></span>

<span data-ttu-id="bfe84-142">若要删除注释，请使用 `Comment.delete` 方法。</span><span class="sxs-lookup"><span data-stu-id="bfe84-142">To delete a comment use the `Comment.delete` method.</span></span> <span data-ttu-id="bfe84-143">删除注释的同时也会删除与该注释相关的答复。</span><span class="sxs-lookup"><span data-stu-id="bfe84-143">Deleting a comment also deletes the replies associated with that comment.</span></span>

```js
Excel.run(function (context) {
    // Delete the comment thread at A2 on the "MyWorksheet" worksheet.
    context.workbook.comments.getItemByCell("MyWorksheet!A2").delete();
    return context.sync();
});
```

### <a name="delete-comment-replies"></a><span data-ttu-id="bfe84-144">删除批注答复</span><span class="sxs-lookup"><span data-stu-id="bfe84-144">Delete comment replies</span></span>

<span data-ttu-id="bfe84-145">若要删除批注答复，请使用 `CommentReply.delete` 方法。</span><span class="sxs-lookup"><span data-stu-id="bfe84-145">To delete a comment reply, use the `CommentReply.delete` method.</span></span>

```js
Excel.run(function (context) {
    // Delete the first comment reply from this worksheet's first comment.
    var comment = context.workbook.comments.getItemAt(0);
    comment.replies.getItemAt(0).delete();
    return context.sync();
});
```

## <a name="resolve-comment-threads"></a><span data-ttu-id="bfe84-146">解析注释线程</span><span class="sxs-lookup"><span data-stu-id="bfe84-146">Resolve comment threads</span></span>

<span data-ttu-id="bfe84-147">注释线程具有可配置的布尔值， `resolved` 以指示是否已解决。</span><span class="sxs-lookup"><span data-stu-id="bfe84-147">A comment thread has a configurable boolean value, `resolved`, to indicate if it is resolved.</span></span> <span data-ttu-id="bfe84-148">值 `true` 表示注释线程已解析。</span><span class="sxs-lookup"><span data-stu-id="bfe84-148">A value of `true` means the comment thread is resolved.</span></span> <span data-ttu-id="bfe84-149">值 `false` 表示注释线程是新的，也可能是重新打开的。</span><span class="sxs-lookup"><span data-stu-id="bfe84-149">A value of `false` means the comment thread is either new or reopened.</span></span>

```js
Excel.run(function (context) {
    // Resolve the first comment thread in the workbook.
    context.workbook.comments.getItemAt(0).resolved = true;
    return context.sync();
});
```

<span data-ttu-id="bfe84-150">批注答复有一个 readonly `resolved` 属性。</span><span class="sxs-lookup"><span data-stu-id="bfe84-150">Comment replies have a readonly `resolved` property.</span></span> <span data-ttu-id="bfe84-151">它的值始终等于线程的其余部分的值。</span><span class="sxs-lookup"><span data-stu-id="bfe84-151">Its value is always equal to that of the rest of the thread.</span></span>

## <a name="comment-metadata"></a><span data-ttu-id="bfe84-152">注释元数据</span><span class="sxs-lookup"><span data-stu-id="bfe84-152">Comment metadata</span></span>

<span data-ttu-id="bfe84-153">每个批注都包含有关其创建情况的元数据，如作者和创建日期。</span><span class="sxs-lookup"><span data-stu-id="bfe84-153">Each comment contains metadata about its creation, such as the author and creation date.</span></span> <span data-ttu-id="bfe84-154">由加载项创建的批注将被视为是由当前用户创作的。</span><span class="sxs-lookup"><span data-stu-id="bfe84-154">Comments created by your add-in are considered to be authored by the current user.</span></span>

<span data-ttu-id="bfe84-155">下面的示例演示如何显示 **A2** 中批注的作者电子邮件、作者姓名和创建日期。</span><span class="sxs-lookup"><span data-stu-id="bfe84-155">The following sample shows how to display the author's email, author's name, and creation date of a comment at **A2**.</span></span>

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

### <a name="comment-reply-metadata"></a><span data-ttu-id="bfe84-156">批注答复元数据</span><span class="sxs-lookup"><span data-stu-id="bfe84-156">Comment reply metadata</span></span>

<span data-ttu-id="bfe84-157">批注答复存储与初始注释相同类型的元数据。</span><span class="sxs-lookup"><span data-stu-id="bfe84-157">Comment replies store the same types of metadata as the initial comment.</span></span>

<span data-ttu-id="bfe84-158">下面的示例展示了如何在**A2**中显示作者的电子邮件、作者的姓名以及最新注释答复的创建日期。</span><span class="sxs-lookup"><span data-stu-id="bfe84-158">The following sample shows how to display the author's email, author's name, and creation date of the latest comment reply at **A2**.</span></span>

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

## <a name="mentions"></a><span data-ttu-id="bfe84-159">提及</span><span class="sxs-lookup"><span data-stu-id="bfe84-159">Mentions</span></span>

<span data-ttu-id="bfe84-160">[提及](https://support.office.com/article/use-mention-in-comments-to-tag-someone-for-feedback-644bf689-31a0-4977-a4fb-afe01820c1fd)用于在注释中标记同事。</span><span class="sxs-lookup"><span data-stu-id="bfe84-160">[Mentions](https://support.office.com/article/use-mention-in-comments-to-tag-someone-for-feedback-644bf689-31a0-4977-a4fb-afe01820c1fd) are used to tag colleagues in a comment.</span></span> <span data-ttu-id="bfe84-161">这将向他们发送你的评论内容通知。</span><span class="sxs-lookup"><span data-stu-id="bfe84-161">This sends them notifications with your comment's content.</span></span> <span data-ttu-id="bfe84-162">你的外接程序可以代表你创建这些提及。</span><span class="sxs-lookup"><span data-stu-id="bfe84-162">Your add-in can create these mentions on your behalf.</span></span>

<span data-ttu-id="bfe84-163">包含提及的注释需要使用[CommentRichContent](/javascript/api/excel/excel.commentrichcontent)对象创建。</span><span class="sxs-lookup"><span data-stu-id="bfe84-163">Comments with mentions need to be created with [CommentRichContent](/javascript/api/excel/excel.commentrichcontent) objects.</span></span> <span data-ttu-id="bfe84-164">`CommentCollection.add`使用 `CommentRichContent` 包含一个或多个提及的调用，并将其指定 `ContentType.mention` 为 `contentType` 参数。</span><span class="sxs-lookup"><span data-stu-id="bfe84-164">Call `CommentCollection.add` with a `CommentRichContent` containing one or more mentions and specify `ContentType.mention` as the `contentType` parameter.</span></span> <span data-ttu-id="bfe84-165">`content`此外，还需要设置字符串格式，以在文本中插入所提及的内容。</span><span class="sxs-lookup"><span data-stu-id="bfe84-165">The `content` string also needs to be formatted to insert the mention into the text.</span></span> <span data-ttu-id="bfe84-166">提及的格式为： `<at id="{replyIndex}">{mentionName}</at>` 。</span><span class="sxs-lookup"><span data-stu-id="bfe84-166">The format for a mention is: `<at id="{replyIndex}">{mentionName}</at>`.</span></span>

> [!NOTE]
> <span data-ttu-id="bfe84-167">目前，只有提及的确切名称可用作提及链接的文本。</span><span class="sxs-lookup"><span data-stu-id="bfe84-167">Currently, only the mention's exact name can be used as the text of the mention link.</span></span> <span data-ttu-id="bfe84-168">稍后将添加对名称的缩写版本的支持。</span><span class="sxs-lookup"><span data-stu-id="bfe84-168">Support for shortened versions of a name will be added later.</span></span>

<span data-ttu-id="bfe84-169">下面的示例展示了一个注释，其中包含一个注明。</span><span class="sxs-lookup"><span data-stu-id="bfe84-169">The following example shows a comment with a single mention.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="bfe84-170">另请参阅</span><span class="sxs-lookup"><span data-stu-id="bfe84-170">See also</span></span>

- [<span data-ttu-id="bfe84-171">Excel JavaScript API 基本编程概念</span><span class="sxs-lookup"><span data-stu-id="bfe84-171">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="bfe84-172">使用 Excel JavaScript API 处理工作簿</span><span class="sxs-lookup"><span data-stu-id="bfe84-172">Work with workbooks using the Excel JavaScript API</span></span>](excel-add-ins-workbooks.md)
- [<span data-ttu-id="bfe84-173">在 Excel 中插入批注和备注</span><span class="sxs-lookup"><span data-stu-id="bfe84-173">Insert comments and notes in Excel</span></span>](https://support.office.com/article/insert-comments-and-notes-in-excel-bdcc9f5d-38e2-45b4-9a92-0b2b5c7bf6f8)
