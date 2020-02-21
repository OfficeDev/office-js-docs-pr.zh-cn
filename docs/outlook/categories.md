---
title: 获取和设置类别
description: 如何管理邮箱和项目上的类别
ms.date: 01/14/2020
localization_priority: Normal
ms.openlocfilehash: 50b98191661674b50c5636733075e4a882183d82
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/20/2020
ms.locfileid: "42166025"
---
# <a name="get-and-set-categories"></a><span data-ttu-id="baee5-103">获取和设置类别</span><span class="sxs-lookup"><span data-stu-id="baee5-103">Get and set categories</span></span>

<span data-ttu-id="baee5-104">在 Outlook 中，用户可以将类别应用于邮件和约会，作为组织其邮箱数据的手段。</span><span class="sxs-lookup"><span data-stu-id="baee5-104">In Outlook, a user can apply categories to messages and appointments as a means of organizing their mailbox data.</span></span> <span data-ttu-id="baee5-105">用户定义其邮箱的颜色编码类别的主列表，然后可以将这些类别中的一个或多个类别应用于任何邮件或约会项目。</span><span class="sxs-lookup"><span data-stu-id="baee5-105">The user defines the master list of color-coded categories for their mailbox, and can then apply one or more of those categories to any message or appointment item.</span></span> <span data-ttu-id="baee5-106">主列表中的每个[类别](/javascript/api/outlook/office.categorydetails)都由用户指定的名称和[颜色](/javascript/api/outlook/office.mailboxenums.categorycolor)表示。</span><span class="sxs-lookup"><span data-stu-id="baee5-106">Each [category](/javascript/api/outlook/office.categorydetails) in the master list is represented by the name and [color](/javascript/api/outlook/office.mailboxenums.categorycolor) that the user specifies.</span></span> <span data-ttu-id="baee5-107">您可以使用 Office JavaScript API 管理邮箱上的类别主机列表和应用于项目的类别。</span><span class="sxs-lookup"><span data-stu-id="baee5-107">You can use the Office JavaScript API to manage the categories master list on the mailbox and the categories applied to an item.</span></span>

> [!NOTE]
> <span data-ttu-id="baee5-108">对此功能的支持是在要求集1.8 中引入的。</span><span class="sxs-lookup"><span data-stu-id="baee5-108">Support for this feature was introduced in requirement set 1.8.</span></span> <span data-ttu-id="baee5-109">请查看支持此要求集的[客户端和平台](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients)。</span><span class="sxs-lookup"><span data-stu-id="baee5-109">See [clients and platforms](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) that support this requirement set.</span></span>

## <a name="manage-categories-in-the-master-list"></a><span data-ttu-id="baee5-110">管理主列表中的类别</span><span class="sxs-lookup"><span data-stu-id="baee5-110">Manage categories in the master list</span></span>

<span data-ttu-id="baee5-111">只有邮箱上的主列表中的类别可供您应用到邮件或约会。</span><span class="sxs-lookup"><span data-stu-id="baee5-111">Only categories in the master list on your mailbox are available for you to apply to a message or appointment.</span></span> <span data-ttu-id="baee5-112">您可以使用 API 添加、获取和删除主类别。</span><span class="sxs-lookup"><span data-stu-id="baee5-112">You can use the API to add, get, and remove master categories.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="baee5-113">若要将外接程序管理类别主机列表，您必须将清单中`Permissions`的节点设置为`ReadWriteMailbox`。</span><span class="sxs-lookup"><span data-stu-id="baee5-113">For the add-in to manage the categories master list, you must set the `Permissions` node in the manifest to `ReadWriteMailbox`.</span></span>

### <a name="add-master-categories"></a><span data-ttu-id="baee5-114">添加母版类别</span><span class="sxs-lookup"><span data-stu-id="baee5-114">Add master categories</span></span>

<span data-ttu-id="baee5-115">下面的示例展示了如何添加名为 "Urgent！" 的类别。</span><span class="sxs-lookup"><span data-stu-id="baee5-115">The following example shows how to add a category named "Urgent!"</span></span> <span data-ttu-id="baee5-116">通过在[masterCategories](/javascript/api/outlook/office.mailbox#mastercategories)上调用[addAsync](/javascript/api/outlook/office.mastercategories#addasync-categories--options--callback-)来指向主列表。</span><span class="sxs-lookup"><span data-stu-id="baee5-116">to the master list by calling [addAsync](/javascript/api/outlook/office.mastercategories#addasync-categories--options--callback-) on [mailbox.masterCategories](/javascript/api/outlook/office.mailbox#mastercategories).</span></span>

```js
var masterCategoriesToAdd = [
    {
        "displayName": "Urgent!",
        "color": Office.MailboxEnums.CategoryColor.Preset0
    }
];

Office.context.mailbox.masterCategories.addAsync(masterCategoriesToAdd, function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        console.log("Successfully added categories to master list");
    } else {
        console.log("masterCategories.addAsync call failed with error: " + asyncResult.error.message);
    }
});
```

### <a name="get-master-categories"></a><span data-ttu-id="baee5-117">获取主类别</span><span class="sxs-lookup"><span data-stu-id="baee5-117">Get master categories</span></span>

<span data-ttu-id="baee5-118">下面的示例演示如何通过在[masterCategories](/javascript/api/outlook/office.mailbox#mastercategories)上调用[getAsync](/javascript/api/outlook/office.mastercategories#getasync-options--callback-)来获取类别的列表。</span><span class="sxs-lookup"><span data-stu-id="baee5-118">The following example shows how to get the list of categories by calling [getAsync](/javascript/api/outlook/office.mastercategories#getasync-options--callback-) on [mailbox.masterCategories](/javascript/api/outlook/office.mailbox#mastercategories).</span></span>

```js
Office.context.mailbox.masterCategories.getAsync(function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.log("Action failed with error: " + asyncResult.error.message);
    } else {
        var masterCategories = asyncResult.value;
        console.log("Master categories:");
        masterCategories.forEach(function (item) {
            console.log("-- " + JSON.stringify(item));
        });
    }
});
```

### <a name="remove-master-categories"></a><span data-ttu-id="baee5-119">删除母版类别</span><span class="sxs-lookup"><span data-stu-id="baee5-119">Remove master categories</span></span>

<span data-ttu-id="baee5-120">下面的示例展示了如何删除名为 "Urgent！" 的类别。</span><span class="sxs-lookup"><span data-stu-id="baee5-120">The following example shows how to remove the category named "Urgent!"</span></span> <span data-ttu-id="baee5-121">通过在[masterCategories](/javascript/api/outlook/office.mailbox#mastercategories)上调用[removeAsync](/javascript/api/outlook/office.mastercategories#removeasync-categories--options--callback-)的主列表。</span><span class="sxs-lookup"><span data-stu-id="baee5-121">from the master list by calling [removeAsync](/javascript/api/outlook/office.mastercategories#removeasync-categories--options--callback-) on [mailbox.masterCategories](/javascript/api/outlook/office.mailbox#mastercategories).</span></span>

```js
var masterCategoriesToRemove = ["Urgent!"];

Office.context.mailbox.masterCategories.removeAsync(masterCategoriesToRemove, function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        console.log("Successfully removed categories from master list");
    } else {
        console.log("masterCategories.removeAsync call failed with error: " + asyncResult.error.message);
    }
});
```

## <a name="manage-categories-on-a-message-or-appointment"></a><span data-ttu-id="baee5-122">管理邮件或约会上的类别</span><span class="sxs-lookup"><span data-stu-id="baee5-122">Manage categories on a message or appointment</span></span>

<span data-ttu-id="baee5-123">您可以使用 API 添加、获取和删除邮件或约会项目的类别。</span><span class="sxs-lookup"><span data-stu-id="baee5-123">You can use the API to add, get, and remove categories for a message or appointment item.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="baee5-124">只有邮箱上的主列表中的类别可供您应用到邮件或约会。</span><span class="sxs-lookup"><span data-stu-id="baee5-124">Only categories in the master list on your mailbox are available for you to apply to a message or appointment.</span></span> <span data-ttu-id="baee5-125">有关详细信息，请参阅上文[中的管理主列表中的类别](#manage-categories-in-the-master-list)一节。</span><span class="sxs-lookup"><span data-stu-id="baee5-125">See the earlier section [Manage categories in the master list](#manage-categories-in-the-master-list) for more information.</span></span>
>
> <span data-ttu-id="baee5-126">在 web 上的 Outlook 中，不能使用 API 在阅读模式下管理邮件的类别。</span><span class="sxs-lookup"><span data-stu-id="baee5-126">In Outlook on the web, you can't use the API to manage categories on a message in Read mode.</span></span>

### <a name="add-categories-to-an-item"></a><span data-ttu-id="baee5-127">将类别添加到项目</span><span class="sxs-lookup"><span data-stu-id="baee5-127">Add categories to an item</span></span>

<span data-ttu-id="baee5-128">下面的示例展示了如何应用名为 "Urgent！" 的类别。</span><span class="sxs-lookup"><span data-stu-id="baee5-128">The following example shows how to apply the category named "Urgent!"</span></span> <span data-ttu-id="baee5-129">通过调用[addAsync](/javascript/api/outlook/office.categories#addasync-categories--options--callback-)的当前项`item.categories`。</span><span class="sxs-lookup"><span data-stu-id="baee5-129">to the current item by calling [addAsync](/javascript/api/outlook/office.categories#addasync-categories--options--callback-) on `item.categories`.</span></span>

```js
var categoriesToAdd = ["Urgent!"];

Office.context.mailbox.item.categories.addAsync(categoriesToAdd, function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        console.log("Successfully added categories");
    } else {
        console.log("categories.addAsync call failed with error: " + asyncResult.error.message);
    }
});
```

### <a name="get-an-items-categories"></a><span data-ttu-id="baee5-130">获取项目的类别</span><span class="sxs-lookup"><span data-stu-id="baee5-130">Get an item's categories</span></span>

<span data-ttu-id="baee5-131">下面的示例演示如何通过在上`item.categories`调用[getAsync](/javascript/api/outlook/office.categories#getasync-options--callback-)来获取应用于当前项的类别。</span><span class="sxs-lookup"><span data-stu-id="baee5-131">The following example shows how to get the categories applied to the current item by calling [getAsync](/javascript/api/outlook/office.categories#getasync-options--callback-) on `item.categories`.</span></span>

```js
Office.context.mailbox.item.categories.getAsync(function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.log("Action failed with error: " + asyncResult.error.message);
    } else {
        var categories = asyncResult.value;
        console.log("Categories:");
        categories.forEach(function (item) {
            console.log("-- " + JSON.stringify(item));
        });
    }
});
```

### <a name="remove-categories-from-an-item"></a><span data-ttu-id="baee5-132">从项目中删除类别</span><span class="sxs-lookup"><span data-stu-id="baee5-132">Remove categories from an item</span></span>

<span data-ttu-id="baee5-133">下面的示例展示了如何删除名为 "Urgent！" 的类别。</span><span class="sxs-lookup"><span data-stu-id="baee5-133">The following example shows how to remove the category named "Urgent!"</span></span> <span data-ttu-id="baee5-134">通过调用[removeAsync](/javascript/api/outlook/office.categories#removeasync-categories--options--callback-)的当前项目`item.categories`。</span><span class="sxs-lookup"><span data-stu-id="baee5-134">from the current item by calling [removeAsync](/javascript/api/outlook/office.categories#removeasync-categories--options--callback-) on `item.categories`.</span></span>

```js
var categoriesToRemove = ["Urgent!"];

Office.context.mailbox.item.categories.removeAsync(categoriesToRemove, function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        console.log("Successfully removed categories");
    } else {
        console.log("categories.removeAsync call failed with error: " + asyncResult.error.message);
    }
});
```

## <a name="see-also"></a><span data-ttu-id="baee5-135">另请参阅</span><span class="sxs-lookup"><span data-stu-id="baee5-135">See also</span></span>

- [<span data-ttu-id="baee5-136">Outlook 权限</span><span class="sxs-lookup"><span data-stu-id="baee5-136">Outlook permissions</span></span>](understanding-outlook-add-in-permissions.md)
- [<span data-ttu-id="baee5-137">清单中的权限元素</span><span class="sxs-lookup"><span data-stu-id="baee5-137">Permissions element in the manifest</span></span>](../reference/manifest/permissions.md)
