---
title: 获取和设置类别
description: 如何管理邮箱和项上的类别。
ms.date: 10/03/2022
ms.localizationpriority: medium
ms.openlocfilehash: a94aba61d513becf2fa1af27ff388b1286e94707
ms.sourcegitcommit: 005783ddd43cf6582233be1be6e3463d7ab9b0e5
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/05/2022
ms.locfileid: "68467102"
---
# <a name="get-and-set-categories"></a>获取和设置类别

在 Outlook 中，用户可以将类别应用于邮件和约会，作为组织其邮箱数据的一种手段。 用户为其邮箱定义颜色编码类别的主列表，然后可以将其中一个或多个类别应用于任何邮件或约会项目。 主列表中的每个 [类别](/javascript/api/outlook/office.categorydetails) 都由用户指定的名称和 [颜色](/javascript/api/outlook/office.mailboxenums.categorycolor) 表示。 可以使用 Office JavaScript API 管理邮箱上的类别主列表以及应用于项目的类别。

> [!NOTE]
> 要求集 1.8 中引入了对此功能的支持。 请查看支持此要求集的[客户端和平台](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets#requirement-sets-supported-by-exchange-servers-and-outlook-clients)。

## <a name="manage-categories-in-the-master-list"></a>管理主列表中的类别

只有邮箱上主列表中的类别可供你应用到邮件或约会。 可以使用 API 添加、获取和删除主类别。

> [!IMPORTANT]
> 若要使加载项管理类别主列表，它必须在清单中请求 **读/写邮箱** 权限。 标记因清单类型而异。
>
> - **XML 清单**：将 **\<Permissions\>** 元素设置为 **ReadWriteMailbox**。
> - **Teams 清单 (预览)**：将“authorization.permissions.resourceSpecific”数组中对象的“name”属性设置为“Mailbox.ReadWrite.User”。

### <a name="add-master-categories"></a>添加主类别

以下示例演示如何添加名为“紧急！” 的类别 通过在 [mailbox.masterCategories](/javascript/api/outlook/office.mailbox#outlook-office-mailbox-mastercategories-member) 上调用 [addAsync](/javascript/api/outlook/office.mastercategories#outlook-office-mastercategories-addasync-member(1)) 来访问主列表。

```js
const masterCategoriesToAdd = [
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

以下示例演示如何通过在 [mailbox.masterCategories](/javascript/api/outlook/office.mailbox#outlook-office-mailbox-mastercategories-member) 上调用 [getAsync](/javascript/api/outlook/office.mastercategories#outlook-office-mastercategories-getasync-member(1)) 来获取类别列表。

```js
Office.context.mailbox.masterCategories.getAsync(function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.log("Action failed with error: " + asyncResult.error.message);
    } else {
        const masterCategories = asyncResult.value;
        console.log("Master categories:");
        masterCategories.forEach(function (item) {
            console.log("-- " + JSON.stringify(item));
        });
    }
});
```

### <a name="remove-master-categories"></a>删除主类别

以下示例演示如何删除名为“紧急！” 的类别 通过在 [mailbox.masterCategories](/javascript/api/outlook/office.mailbox#outlook-office-mailbox-mastercategories-member) 上调用 [removeAsync](/javascript/api/outlook/office.mastercategories#outlook-office-mastercategories-removeasync-member(1)) 从主列表中。

```js
const masterCategoriesToRemove = ["Urgent!"];

Office.context.mailbox.masterCategories.removeAsync(masterCategoriesToRemove, function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        console.log("Successfully removed categories from master list");
    } else {
        console.log("masterCategories.removeAsync call failed with error: " + asyncResult.error.message);
    }
});
```

## <a name="manage-categories-on-a-message-or-appointment"></a>管理邮件或约会的类别

可以使用 API 添加、获取和删除邮件或约会项目的类别。

> [!IMPORTANT]
> 只有邮箱上主列表中的类别可供你应用到邮件或约会。 有关详细信息，请参阅 [母版列表中“管理类别](#manage-categories-in-the-master-list) ”的前面部分。
>
> 在Outlook 网页版中，不能使用 API 在读取模式下管理邮件的类别。

### <a name="add-categories-to-an-item"></a>向项添加类别

以下示例演示如何应用名为“紧急！” 的类别 通过调用 [addAsync](/javascript/api/outlook/office.categories#outlook-office-categories-addasync-member(1)) on `item.categories`调用当前项。

```js
const categoriesToAdd = ["Urgent!"];

Office.context.mailbox.item.categories.addAsync(categoriesToAdd, function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        console.log("Successfully added categories");
    } else {
        console.log("categories.addAsync call failed with error: " + asyncResult.error.message);
    }
});
```

### <a name="get-an-items-categories"></a>获取项的类别

以下示例演示如何通过调用 [getAsync](/javascript/api/outlook/office.categories#outlook-office-categories-getasync-member(1)) `item.categories`来获取应用于当前项的类别。

```js
Office.context.mailbox.item.categories.getAsync(function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.log("Action failed with error: " + asyncResult.error.message);
    } else {
        const categories = asyncResult.value;
        console.log("Categories:");
        categories.forEach(function (item) {
            console.log("-- " + JSON.stringify(item));
        });
    }
});
```

### <a name="remove-categories-from-an-item"></a>从项中删除类别

以下示例演示如何删除名为“紧急！” 的类别 通过调用 [removeAsync](/javascript/api/outlook/office.categories#outlook-office-categories-removeasync-member(1)) on 从当前项。`item.categories`

```js
const categoriesToRemove = ["Urgent!"];

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
