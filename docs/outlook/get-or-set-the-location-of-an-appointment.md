---
title: 在加载项获取或设置约会位置
description: 了解如何在 Outlook 加载项中获取或设置约会位置。
ms.date: 10/31/2019
localization_priority: Normal
ms.openlocfilehash: 5669f656348465baabb3e684b359261024a509ca
ms.sourcegitcommit: 3fa8c754a47bab909e559ae3e5d4237ba27fdbe4
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/30/2021
ms.locfileid: "53671833"
---
# <a name="get-or-set-the-location-when-composing-an-appointment-in-outlook"></a>在 Outlook 中撰写约会时获取或设置位置

JavaScript API Office JavaScript API 提供用于管理用户正在撰写的约会位置的属性和方法。 目前，有两个属性提供约会的位置：

- [item.location：](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)允许你获取和设置位置的基本 API。
- [item.enhancedLocation](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)：可让你获取和设置位置的增强型 API，包括指定 [位置类型](/javascript/api/outlook/office.mailboxenums.locationtype)。 类型是 `LocationType.Custom` ，如果使用 设置位置 `item.location` 。

下表列出了位置 API 和模式 (，即撰写或) 的可用模式。

| API | 适用的约会模式 |
|---|---|
| [item.location](/javascript/api/outlook/office.appointmentread#location) | Attendee/Read |
| [item.location.getAsync](/javascript/api/outlook/office.location#getAsync_options__callback_) | 组织者/撰写 |
| [item.location.setAsync](/javascript/api/outlook/office.location#setAsync_location__options__callback_) | 组织者/撰写 |
| [item.enhancedLocation.getAsync](/javascript/api/outlook/office.enhancedlocation#getAsync_options__callback_) | 组织者/撰写<br>Attendee/Read |
| [item.enhancedLocation.addAsync](/javascript/api/outlook/office.enhancedlocation#addAsync_locationIdentifiers__options__callback_) | 组织者/撰写 |
| [item.enhancedLocation.removeAsync](/javascript/api/outlook/office.enhancedlocation#removeAsync_locationIdentifiers__options__callback_) | 组织者/撰写 |

若要使用仅撰写外接程序的方法，请配置外接程序清单以在管理器/撰写模式下激活外接程序。 有关详细信息[，Outlook创建适用于撰写窗体的](compose-scenario.md)外接程序。

## <a name="use-the-enhancedlocation-api"></a>使用 `enhancedLocation` API

可以使用 API `enhancedLocation` 获取和设置约会的位置。 位置字段支持多个位置，并且对于每个位置，显示名称设置会议室电子邮件地址、类型和会议室 (（如果适用) ）。 有关支持的位置类型，请参阅[LocationType。](/javascript/api/outlook/office.mailboxenums.locationtype)

### <a name="add-location"></a>添加位置

下面的示例演示如何通过调用[mailbox.item.enhancedLocation](/javascript/api/outlook/office.appointmentcompose#enhancedLocation)上的[addAsync](/javascript/api/outlook/office.enhancedlocation#addAsync_locationIdentifiers__options__callback_)添加位置。

```js
var item;
var locations = [
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

下面的示例演示如何通过调用[mailbox.item.enhancedLocation](/javascript/api/outlook/office.appointmentread#enhancedLocation)上的[getAsync](/javascript/api/outlook/office.enhancedlocation#getAsync_options__callback_)获取位置。

```js
var item;

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

下面的示例演示如何通过调用[mailbox.item.enhancedLocation](/javascript/api/outlook/office.appointmentcompose#enhancedLocation)上的[removeAsync](/javascript/api/outlook/office.enhancedlocation#removeAsync_locationIdentifiers__options__callback_)来删除位置。

```js
var item;

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
        Office.context.mailbox.item.enhancedLocation.removeAsync([currentValue.locationIdentifier]);
    });
}
```

## <a name="use-the-location-api"></a>使用 `location` API

可以使用 API `location` 获取和设置约会的位置。

### <a name="get-the-location"></a>获取位置

此部分显示了一个代码示例，用于获取用户正在撰写的约会的位置，并显示该位置。

若要使用 `item.location.getAsync`，请提供回调方法，用于检查异步调用的状态和结果。 可以通过 `asyncContext` 可选参数为回调方法提供任何必要的参数。 可以使用回调的输出参数获取状态、结果 `asyncResult` 和任何错误。 如果异步调用成功，可以使用 [AsyncResult.value](/javascript/api/office/office.asyncresult#value) 属性获取作为字符串的位置。

```js
var item;

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

若要使用 `item.location.setAsync`，请在数据参数中指定一个最多 255 个字符的字符串。 或者，可以提供一个回调方法，并在 `asyncContext` 参数中为该回调方法提供任何自变量。 应检查回调的输出参数中的状态、结果 `asyncResult` 和任何错误消息。 如果异步调用成功，`setAsync` 会将指定位置字符串作为纯文本插入，同时覆盖相应项的任何现有位置。

> [!NOTE]
> 可以使用分号作为分隔符来设置多个 (例如，"会议室 A;会议室 B) 。

```js
var item;

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

- [创建首个Outlook加载项](../quickstarts/outlook-quickstart.md)
- [Office 外接程序中的异步编程](../develop/asynchronous-programming-in-office-add-ins.md)
