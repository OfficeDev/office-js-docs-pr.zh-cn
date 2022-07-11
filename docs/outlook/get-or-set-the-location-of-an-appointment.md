---
title: 在加载项获取或设置约会位置
description: 了解如何在 Outlook 加载项中获取或设置约会位置。
ms.date: 07/08/2022
ms.localizationpriority: medium
ms.openlocfilehash: 0a9b1ce064ff069aa48f6d9c979edc2dfa4fd3a4
ms.sourcegitcommit: d8ea4b761f44d3227b7f2c73e52f0d2233bf22e2
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/11/2022
ms.locfileid: "66713096"
---
# <a name="get-or-set-the-location-when-composing-an-appointment-in-outlook"></a>在 Outlook 中撰写约会时获取或设置位置

Office JavaScript API 提供属性和方法来管理用户正在撰写约会的位置。 目前，有两个属性提供约会的位置：

- [item.location](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)：可获取和设置位置的基本 API。
- [item.enhancedLocation](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)：增强的 API，可用于获取和设置位置，并包括指定 [位置类型](/javascript/api/outlook/office.mailboxenums.locationtype)。 类型是 `LocationType.Custom` 如果使用 `item.location`>a0/a0> 设置位置。

下表列出了位置 API 和模式 (，例如，Compose 或 Read) 可用的位置。

| API | 适用的约会模式 |
|---|---|
| [item.location](/javascript/api/outlook/office.appointmentread#outlook-office-appointmentread-location-member) | 与会者/阅读者 |
| [item.location.getAsync](/javascript/api/outlook/office.location#outlook-office-location-getasync-member(1)) | 组织者/撰写者 |
| [item.location.setAsync](/javascript/api/outlook/office.location#outlook-office-location-setasync-member(1)) | 组织者/撰写者 |
| [item.enhancedLocation.getAsync](/javascript/api/outlook/office.enhancedlocation#outlook-office-enhancedlocation-getasync-member(1)) | 组织者/撰写者，<br>与会者/阅读者 |
| [item.enhancedLocation.addAsync](/javascript/api/outlook/office.enhancedlocation#outlook-office-enhancedlocation-addasync-member(1)) | 组织者/撰写者 |
| [item.enhancedLocation.removeAsync](/javascript/api/outlook/office.enhancedlocation#outlook-office-enhancedlocation-removeasync-member(1)) | 组织者/撰写者 |

若要使用仅可用于撰写外接程序的方法，请将外接程序清单配置为在“组织者/撰写”模式下激活外接程序。 有关更多详细信息，请参阅 [“创建 Outlook 外接程序”以获取撰写表单](compose-scenario.md) 。

## <a name="use-the-enhancedlocation-api"></a>`enhancedLocation`使用 API

可以使用 `enhancedLocation` API 获取和设置约会的位置。 位置字段支持多个位置，对于每个位置，可以设置显示名称、类型和会议室电子邮件地址 (（如果适用）) 。 有关受支持的位置类型，请参阅 [LocationType](/javascript/api/outlook/office.mailboxenums.locationtype) 。

### <a name="add-location"></a>添加位置

以下示例演示如何通过在 [mailbox.item.enhancedLocation](/javascript/api/outlook/office.appointmentcompose#outlook-office-appointmentcompose-enhancedlocation-member) 上调用 [addAsync](/javascript/api/outlook/office.enhancedlocation#outlook-office-enhancedlocation-addasync-member(1)) 来添加位置。

```js
let item;
const locations = [
    {
        "id": "Contoso",
        "type": Office.MailboxEnums.LocationType.Custom
    }
];

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Check for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Add to the location of the item being composed.
        item.enhancedLocation.addAsync(locations);
    });
}
```

### <a name="get-location"></a>获取位置

以下示例演示如何通过在 [mailbox.item.enhancedLocation](/javascript/api/outlook/office.appointmentread#outlook-office-appointmentread-enhancedlocation-member) 上调用 [getAsync](/javascript/api/outlook/office.enhancedlocation#outlook-office-enhancedlocation-getasync-member(1)) 来获取位置。

```js
let item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Get the location of the item being composed.
        item.enhancedLocation.getAsync(callbackFunction);
    });
}

function callbackFunction(asyncResult) {
    asyncResult.value.forEach(function (place) {
        console.log("Display name: " + place.displayName);
        console.log("Type: " + place.locationIdentifier.type);
        if (place.locationIdentifier.type === Office.MailboxEnums.LocationType.Room) {
            console.log("Email address: " + place.emailAddress);
        }
    });
}
```

### <a name="remove-location"></a>删除位置

以下示例演示如何通过在 [mailbox.item.enhancedLocation](/javascript/api/outlook/office.appointmentcompose#outlook-office-appointmentcompose-enhancedlocation-member) 上调用 [removeAsync](/javascript/api/outlook/office.enhancedlocation#outlook-office-enhancedlocation-removeasync-member(1)) 来删除位置。

```js
let item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Get the location of the item being composed.
        item.enhancedLocation.getAsync(callbackFunction);
    });
}

function callbackFunction(asyncResult) {
    asyncResult.value.forEach(function (currentValue) {
        // Remove each location from the item being composed.
        item.enhancedLocation.removeAsync([currentValue.locationIdentifier]);
    });
}
```

## <a name="use-the-location-api"></a>`location`使用 API

可以使用 `location` API 获取和设置约会的位置。

### <a name="get-the-location"></a>获取位置

此部分显示了一个代码示例，用于获取用户正在撰写的约会的位置，并显示该位置。

若要使用 `item.location.getAsync`，请提供回调方法，用于检查异步调用的状态和结果。 可以通过 `asyncContext` 可选参数为回调方法提供任何必要的参数。 可以使用回调的输出参数 `asyncResult` 获取状态、结果和任何错误。 如果异步调用成功，可以使用 [AsyncResult.value](/javascript/api/office/office.asyncresult#office-office-asyncresult-value-member) 属性获取作为字符串的位置。

```js
let item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Get the location of the item being composed.
        getLocation();
    });
}

// Get the location of the item that the user is composing.
function getLocation() {
    item.location.getAsync(
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Successfully got the location, display it.
                write ('The location is: ' + asyncResult.value);
            }
        });
}

// Write to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message;
}
```

### <a name="set-the-location"></a>设置位置

此部分显示了一个代码示例，用于设置用户正在撰写的约会的位置。

若要使用 `item.location.setAsync`，请在数据参数中指定一个最多 255 个字符的字符串。 或者，可以提供一个回调方法，并在 `asyncContext` 参数中为该回调方法提供任何自变量。 应检查回调的输出参数中的 `asyncResult` 状态、结果和任何错误消息。 如果异步调用成功，`setAsync` 会将指定位置字符串作为纯文本插入，同时覆盖相应项的任何现有位置。

> [!NOTE]
> 可以使用分号作为分隔符设置多个位置 (例如“会议室 A;会议室 B') 。

```js
let item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Check for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Set the location of the item being composed.
        setLocation();
    });
}

// Set the location of the item that the user is composing.
function setLocation() {
    item.location.setAsync(
        'Conference room A',
        { asyncContext: { var1: 1, var2: 2 } },
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Successfully set the location.
                // Do whatever is appropriate for your scenario,
                // using the arguments var1 and var2 as applicable.
            }
        });
}

// Write to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message;
}
```

## <a name="see-also"></a>另请参阅

- [创建第一个 Outlook 加载项](../quickstarts/outlook-quickstart.md)
- [Office 外接程序中的异步编程](../develop/asynchronous-programming-in-office-add-ins.md)
