---
title: 获取和设置类别
description: 如何管理邮箱和项目的类别。
ms.date: 01/14/2020
ms.localizationpriority: medium
ms.openlocfilehash: 93f9167fcc31110543d08019e5428952beab0ccc
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/23/2022
ms.locfileid: "63746298"
---
# <a name="get-and-set-categories"></a>获取和设置类别

在Outlook中，用户可以将类别应用于邮件和约会，以用作组织其邮箱数据的方式。 用户定义其邮箱的颜色编码类别主列表，然后可以将这些类别的一个或多个应用于任何邮件或约会项目。 [主](/javascript/api/outlook/office.categorydetails)列表中的每个类别都由用户[指定的](/javascript/api/outlook/office.mailboxenums.categorycolor)名称和颜色表示。 可以使用 JavaScript API Office邮箱上的类别主列表以及应用于项目的类别。

> [!NOTE]
> 要求集 1.8 中引入了对此功能的支持。 请查看支持此要求集的[客户端和平台](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients)。

## <a name="manage-categories-in-the-master-list"></a>管理主列表中的类别

只有邮箱上主列表中的类别可以应用于邮件或约会。 可以使用 API 添加、获取和删除主类别。

> [!IMPORTANT]
> 若要使加载项管理类别主列表 `Permissions` ，必须将清单中的 `ReadWriteMailbox`节点设置为 。

### <a name="add-master-categories"></a>添加主类别

以下示例演示如何添加名为"Urgent！"的类别。 通过调用 [mailbox.masterCategories](/javascript/api/outlook/office.mailbox#outlook-office-mailbox-mastercategories-member) 上的 [addAsync](/javascript/api/outlook/office.mastercategories#outlook-office-mastercategories-addasync-member(1)) 来访问主列表。

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

以下示例演示如何通过调用 [mailbox.masterCategories](/javascript/api/outlook/office.mailbox#outlook-office-mailbox-mastercategories-member) 上的 [getAsync](/javascript/api/outlook/office.mastercategories#outlook-office-mastercategories-getasync-member(1)) 获取类别列表。

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

### <a name="remove-master-categories"></a>删除主类别

以下示例演示如何删除名为"Urgent！"的类别。 通过调用 [mailbox.masterCategories](/javascript/api/outlook/office.mailbox#outlook-office-mailbox-mastercategories-member) 上的 [removeAsync](/javascript/api/outlook/office.mastercategories#outlook-office-mastercategories-removeasync-member(1)) 从主列表。

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

可以使用 API 添加、获取和删除邮件或约会项目的类别。

> [!IMPORTANT]
> 只有邮箱上主列表中的类别可以应用于邮件或约会。 有关详细信息，请参阅上 [一节"管理主列表中的](#manage-categories-in-the-master-list) 类别"。
>
> 在Outlook 网页版中，你无法以阅读模式使用 API 管理邮件的类别。

### <a name="add-categories-to-an-item"></a>向项目添加类别

以下示例演示如何应用名为"Urgent！"的类别。 调用 上的 [addAsync，以访问当前](/javascript/api/outlook/office.categories#outlook-office-categories-addasync-member(1)) 项 `item.categories`。

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

以下示例演示如何通过调用 上的 [getAsync](/javascript/api/outlook/office.categories#outlook-office-categories-getasync-member(1)) 获取应用于当前项目的类别 `item.categories`。

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

以下示例演示如何删除名为"Urgent！"的类别。 通过调用 上的 [removeAsync 从当前](/javascript/api/outlook/office.categories#outlook-office-categories-removeasync-member(1))项。`item.categories`

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

- [Outlook权限](understanding-outlook-add-in-permissions.md)
- [清单中的 Permissions 元素](../reference/manifest/permissions.md)
