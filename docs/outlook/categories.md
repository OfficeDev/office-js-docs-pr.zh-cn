---
title: 获取和设置类别
description: 如何管理邮箱和项目上的类别
ms.date: 01/14/2020
localization_priority: Normal
ms.openlocfilehash: d0bb2e9f51675c263d0a3a130c64e02e7d55b764
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/17/2020
ms.locfileid: "42721021"
---
# <a name="get-and-set-categories"></a>获取和设置类别

在 Outlook 中，用户可以将类别应用于邮件和约会，作为组织其邮箱数据的手段。 用户定义其邮箱的颜色编码类别的主列表，然后可以将这些类别中的一个或多个类别应用于任何邮件或约会项目。 主列表中的每个[类别](/javascript/api/outlook/office.categorydetails)都由用户指定的名称和[颜色](/javascript/api/outlook/office.mailboxenums.categorycolor)表示。 您可以使用 Office JavaScript API 管理邮箱上的类别主机列表和应用于项目的类别。

> [!NOTE]
> 对此功能的支持是在要求集1.8 中引入的。 请查看支持此要求集的[客户端和平台](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients)。

## <a name="manage-categories-in-the-master-list"></a>管理主列表中的类别

只有邮箱上的主列表中的类别可供您应用到邮件或约会。 您可以使用 API 添加、获取和删除主类别。

> [!IMPORTANT]
> 若要将外接程序管理类别主机列表，您必须将清单中`Permissions`的节点设置为`ReadWriteMailbox`。

### <a name="add-master-categories"></a>添加母版类别

下面的示例展示了如何添加名为 "Urgent！" 的类别。 通过在[masterCategories](/javascript/api/outlook/office.mailbox#mastercategories)上调用[addAsync](/javascript/api/outlook/office.mastercategories#addasync-categories--options--callback-)来指向主列表。

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

### <a name="get-master-categories"></a>获取主类别

下面的示例演示如何通过在[masterCategories](/javascript/api/outlook/office.mailbox#mastercategories)上调用[getAsync](/javascript/api/outlook/office.mastercategories#getasync-options--callback-)来获取类别的列表。

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

### <a name="remove-master-categories"></a>删除母版类别

下面的示例展示了如何删除名为 "Urgent！" 的类别。 通过在[masterCategories](/javascript/api/outlook/office.mailbox#mastercategories)上调用[removeAsync](/javascript/api/outlook/office.mastercategories#removeasync-categories--options--callback-)的主列表。

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

## <a name="manage-categories-on-a-message-or-appointment"></a>管理邮件或约会上的类别

您可以使用 API 添加、获取和删除邮件或约会项目的类别。

> [!IMPORTANT]
> 只有邮箱上的主列表中的类别可供您应用到邮件或约会。 有关详细信息，请参阅上文[中的管理主列表中的类别](#manage-categories-in-the-master-list)一节。
>
> 在 web 上的 Outlook 中，不能使用 API 在阅读模式下管理邮件的类别。

### <a name="add-categories-to-an-item"></a>将类别添加到项目

下面的示例展示了如何应用名为 "Urgent！" 的类别。 通过调用[addAsync](/javascript/api/outlook/office.categories#addasync-categories--options--callback-)的当前项`item.categories`。

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

### <a name="get-an-items-categories"></a>获取项目的类别

下面的示例演示如何通过在上`item.categories`调用[getAsync](/javascript/api/outlook/office.categories#getasync-options--callback-)来获取应用于当前项的类别。

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

### <a name="remove-categories-from-an-item"></a>从项目中删除类别

下面的示例展示了如何删除名为 "Urgent！" 的类别。 通过调用[removeAsync](/javascript/api/outlook/office.categories#removeasync-categories--options--callback-)的当前项目`item.categories`。

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

## <a name="see-also"></a>另请参阅

- [Outlook 权限](understanding-outlook-add-in-permissions.md)
- [清单中的权限元素](../reference/manifest/permissions.md)
