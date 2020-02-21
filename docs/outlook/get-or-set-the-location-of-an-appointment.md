---
title: 在加载项获取或设置约会位置
description: 了解如何在 Outlook 加载项中获取或设置约会位置。
ms.date: 10/31/2019
localization_priority: Normal
ms.openlocfilehash: 7e2c2b604948b7630581af03aa9f8fddc4c68da6
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/20/2020
ms.locfileid: "42166038"
---
# <a name="get-or-set-the-location-when-composing-an-appointment-in-outlook"></a><span data-ttu-id="97c25-103">在 Outlook 中撰写约会时获取或设置位置</span><span class="sxs-lookup"><span data-stu-id="97c25-103">Get or set the location when composing an appointment in Outlook</span></span>

<span data-ttu-id="97c25-104">适用于 Office 的 JavaScript API 提供了用于管理用户正在撰写的约会的位置的属性和方法。</span><span class="sxs-lookup"><span data-stu-id="97c25-104">The JavaScript API for Office provides properties and methods to manage the location of an appointment that the user is composing.</span></span> <span data-ttu-id="97c25-105">目前，有两个属性可提供约会的位置：</span><span class="sxs-lookup"><span data-stu-id="97c25-105">Currently, there are two properties that provide an appointment's location:</span></span>

- <span data-ttu-id="97c25-106">[item： location](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)：允许你获取和设置位置的基本 API。</span><span class="sxs-lookup"><span data-stu-id="97c25-106">[item.location](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties): Basic API that allows you to get and set the location.</span></span>
- <span data-ttu-id="97c25-107">[enhancedLocation](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)：增强 API，允许你获取和设置位置，并包括指定[位置类型](/javascript/api/outlook/office.mailboxenums.locationtype)。</span><span class="sxs-lookup"><span data-stu-id="97c25-107">[item.enhancedLocation](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties): Enhanced API that allows you to get and set the location, and includes specifying the [location type](/javascript/api/outlook/office.mailboxenums.locationtype).</span></span> <span data-ttu-id="97c25-108">键入的是`LocationType.Custom`使用`item.location`设置的位置。</span><span class="sxs-lookup"><span data-stu-id="97c25-108">The type is `LocationType.Custom` if you set the location using `item.location`.</span></span>

<span data-ttu-id="97c25-109">下表列出了位置 Api 以及它们的可用模式（即撰写或读取）。</span><span class="sxs-lookup"><span data-stu-id="97c25-109">The following table lists the location APIs and the modes (i.e., Compose or Read) where they are available.</span></span>

| <span data-ttu-id="97c25-110">API</span><span class="sxs-lookup"><span data-stu-id="97c25-110">API</span></span> | <span data-ttu-id="97c25-111">适用的约会模式</span><span class="sxs-lookup"><span data-stu-id="97c25-111">Applicable appointment modes</span></span> |
|---|---|
| [<span data-ttu-id="97c25-112">项。位置</span><span class="sxs-lookup"><span data-stu-id="97c25-112">item.location</span></span>](/javascript/api/outlook/office.appointmentread#location) | <span data-ttu-id="97c25-113">与会者/阅读</span><span class="sxs-lookup"><span data-stu-id="97c25-113">Attendee/Read</span></span> |
| [<span data-ttu-id="97c25-114">项。 getAsync</span><span class="sxs-lookup"><span data-stu-id="97c25-114">item.location.getAsync</span></span>](/javascript/api/outlook/office.location#getasync-options--callback-) | <span data-ttu-id="97c25-115">组织者/撰写</span><span class="sxs-lookup"><span data-stu-id="97c25-115">Organizer/Compose</span></span> |
| [<span data-ttu-id="97c25-116">item.location.setAsync</span><span class="sxs-lookup"><span data-stu-id="97c25-116">item.location.setAsync</span></span>](/javascript/api/outlook/office.location#setasync-location--options--callback-) | <span data-ttu-id="97c25-117">组织者/撰写</span><span class="sxs-lookup"><span data-stu-id="97c25-117">Organizer/Compose</span></span> |
| [<span data-ttu-id="97c25-118">enhancedLocation。 getAsync</span><span class="sxs-lookup"><span data-stu-id="97c25-118">item.enhancedLocation.getAsync</span></span>](/javascript/api/outlook/office.enhancedlocation#getasync-options--callback-) | <span data-ttu-id="97c25-119">组织者/撰写、</span><span class="sxs-lookup"><span data-stu-id="97c25-119">Organizer/Compose,</span></span><br><span data-ttu-id="97c25-120">与会者/阅读</span><span class="sxs-lookup"><span data-stu-id="97c25-120">Attendee/Read</span></span> |
| [<span data-ttu-id="97c25-121">enhancedLocation。 addAsync</span><span class="sxs-lookup"><span data-stu-id="97c25-121">item.enhancedLocation.addAsync</span></span>](/javascript/api/outlook/office.enhancedlocation#addasync-locationidentifiers--options--callback-) | <span data-ttu-id="97c25-122">组织者/撰写</span><span class="sxs-lookup"><span data-stu-id="97c25-122">Organizer/Compose</span></span> |
| [<span data-ttu-id="97c25-123">enhancedLocation。 removeAsync</span><span class="sxs-lookup"><span data-stu-id="97c25-123">item.enhancedLocation.removeAsync</span></span>](/javascript/api/outlook/office.enhancedlocation#removeasync-locationidentifiers--options--callback-) | <span data-ttu-id="97c25-124">组织者/撰写</span><span class="sxs-lookup"><span data-stu-id="97c25-124">Organizer/Compose</span></span> |

<span data-ttu-id="97c25-125">若要使用仅适用于撰写外接程序的方法，请配置外接程序清单以在组织者/撰写模式下激活外接程序。</span><span class="sxs-lookup"><span data-stu-id="97c25-125">To use the methods that are available only to compose add-ins, configure the add-in manifest to activate the add-in in Organizer/Compose mode.</span></span> <span data-ttu-id="97c25-126">有关详细信息，请参阅[创建适用于撰写窗体的 Outlook 外接程序](compose-scenario.md)。</span><span class="sxs-lookup"><span data-stu-id="97c25-126">See [Create Outlook add-ins for compose forms](compose-scenario.md) for more details.</span></span>

## <a name="use-the-enhancedlocation-api"></a><span data-ttu-id="97c25-127">使用`enhancedLocation` API</span><span class="sxs-lookup"><span data-stu-id="97c25-127">Use the `enhancedLocation` API</span></span>

<span data-ttu-id="97c25-128">您可以使用`enhancedLocation` API 来获取和设置约会的位置。</span><span class="sxs-lookup"><span data-stu-id="97c25-128">You can use the `enhancedLocation` API to get and set an appointment's location.</span></span> <span data-ttu-id="97c25-129">"位置" 字段支持多个位置，并且对于每个位置，可以设置显示名称、类型和会议室电子邮件地址（如果适用）。</span><span class="sxs-lookup"><span data-stu-id="97c25-129">The location field supports multiple locations and, for each location, you can set the display name, type, and conference room email address (if applicable).</span></span> <span data-ttu-id="97c25-130">有关支持的位置类型，请参阅[LocationType](/javascript/api/outlook/office.mailboxenums.locationtype) 。</span><span class="sxs-lookup"><span data-stu-id="97c25-130">See [LocationType](/javascript/api/outlook/office.mailboxenums.locationtype) for supported location types.</span></span>

### <a name="add-location"></a><span data-ttu-id="97c25-131">添加位置</span><span class="sxs-lookup"><span data-stu-id="97c25-131">Add location</span></span>

<span data-ttu-id="97c25-132">下面的示例演示如何通过对[enhancedLocation](/javascript/api/outlook/office.appointmentcompose#enhancedlocation)中的[addAsync](/javascript/api/outlook/office.enhancedlocation#addasync-locationidentifiers--options--callback-)调用的方式来添加位置。</span><span class="sxs-lookup"><span data-stu-id="97c25-132">The following example shows how to add a location by calling [addAsync](/javascript/api/outlook/office.enhancedlocation#addasync-locationidentifiers--options--callback-) on [mailbox.item.enhancedLocation](/javascript/api/outlook/office.appointmentcompose#enhancedlocation).</span></span>

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

### <a name="get-location"></a><span data-ttu-id="97c25-133">获取位置</span><span class="sxs-lookup"><span data-stu-id="97c25-133">Get location</span></span>

<span data-ttu-id="97c25-134">下面的示例演示如何通过在[enhancedLocation](/javascript/api/outlook/office.appointmentread#enhancedlocation)上调用[getAsync](/javascript/api/outlook/office.enhancedlocation#getasync-options--callback-)来获取位置。</span><span class="sxs-lookup"><span data-stu-id="97c25-134">The following example shows how to get the location by calling [getAsync](/javascript/api/outlook/office.enhancedlocation#getasync-options--callback-) on [mailbox.item.enhancedLocation](/javascript/api/outlook/office.appointmentread#enhancedlocation).</span></span>

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

### <a name="remove-location"></a><span data-ttu-id="97c25-135">删除位置</span><span class="sxs-lookup"><span data-stu-id="97c25-135">Remove location</span></span>

<span data-ttu-id="97c25-136">下面的示例演示如何通过对[enhancedLocation](/javascript/api/outlook/office.appointmentcompose#enhancedlocation)中的[removeAsync](/javascript/api/outlook/office.enhancedlocation#removeasync-locationidentifiers--options--callback-)调用的方式来删除该位置。</span><span class="sxs-lookup"><span data-stu-id="97c25-136">The following example shows how to remove the location by calling [removeAsync](/javascript/api/outlook/office.enhancedlocation#removeasync-locationidentifiers--options--callback-) on [mailbox.item.enhancedLocation](/javascript/api/outlook/office.appointmentcompose#enhancedlocation).</span></span>

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

## <a name="use-the-location-api"></a><span data-ttu-id="97c25-137">使用`location` API</span><span class="sxs-lookup"><span data-stu-id="97c25-137">Use the `location` API</span></span>

<span data-ttu-id="97c25-138">您可以使用`location` API 来获取和设置约会的位置。</span><span class="sxs-lookup"><span data-stu-id="97c25-138">You can use the `location` API to get and set an appointment's location.</span></span>

### <a name="get-the-location"></a><span data-ttu-id="97c25-139">获取位置</span><span class="sxs-lookup"><span data-stu-id="97c25-139">Get the location</span></span>

<span data-ttu-id="97c25-140">此部分显示了一个代码示例，用于获取用户正在撰写的约会的位置，并显示该位置。</span><span class="sxs-lookup"><span data-stu-id="97c25-140">This section shows a code sample that gets the location of the appointment that the user is composing, and displays the location.</span></span>

<span data-ttu-id="97c25-141">若要使用 `item.location.getAsync`，请提供回调方法，用于检查异步调用的状态和结果。</span><span class="sxs-lookup"><span data-stu-id="97c25-141">To use `item.location.getAsync`, provide a callback method that checks for the status and result of the asynchronous call.</span></span> <span data-ttu-id="97c25-142">可以通过 `asyncContext` 可选参数为回调方法提供任何必要的参数。</span><span class="sxs-lookup"><span data-stu-id="97c25-142">You can provide any necessary arguments to the callback method through the `asyncContext` optional parameter.</span></span> <span data-ttu-id="97c25-143">您可以使用回调的 output 参数`asyncResult`获取状态、结果和任何错误。</span><span class="sxs-lookup"><span data-stu-id="97c25-143">You can obtain status, results, and any error using the output parameter `asyncResult` of the callback.</span></span> <span data-ttu-id="97c25-144">如果异步调用成功，可以使用 [AsyncResult.value](/javascript/api/office/office.asyncresult#value) 属性获取作为字符串的位置。</span><span class="sxs-lookup"><span data-stu-id="97c25-144">If the asynchronous call is successful, you can get the location as a string using the [AsyncResult.value](/javascript/api/office/office.asyncresult#value) property.</span></span>

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

### <a name="set-the-location"></a><span data-ttu-id="97c25-145">设置位置</span><span class="sxs-lookup"><span data-stu-id="97c25-145">Set the location</span></span>

<span data-ttu-id="97c25-146">此部分显示了一个代码示例，用于设置用户正在撰写的约会的位置。</span><span class="sxs-lookup"><span data-stu-id="97c25-146">This section shows a code sample that sets the location of the appointment that the user is composing.</span></span>

<span data-ttu-id="97c25-147">若要使用 `item.location.setAsync`，请在数据参数中指定一个最多 255 个字符的字符串。</span><span class="sxs-lookup"><span data-stu-id="97c25-147">To use `item.location.setAsync`, specify a string of up to 255 characters in the data parameter.</span></span> <span data-ttu-id="97c25-148">或者，可以提供一个回调方法，并在 `asyncContext` 参数中为该回调方法提供任何自变量。</span><span class="sxs-lookup"><span data-stu-id="97c25-148">Optionally, you can provide a callback method and any arguments for the callback method in the `asyncContext` parameter.</span></span> <span data-ttu-id="97c25-149">应检查回调的`asyncResult` output 参数中的状态、结果和任何错误消息。</span><span class="sxs-lookup"><span data-stu-id="97c25-149">You should check the status, result, and any error message in the `asyncResult` output parameter of the callback.</span></span> <span data-ttu-id="97c25-150">如果异步调用成功，`setAsync` 会将指定位置字符串作为纯文本插入，同时覆盖相应项的任何现有位置。</span><span class="sxs-lookup"><span data-stu-id="97c25-150">If the asynchronous call is successful, `setAsync` inserts the specified location string as plain text, overwriting any existing location for that item.</span></span>

> [!NOTE]
> <span data-ttu-id="97c25-151">您可以使用分号作为分隔符（例如，"会议室 A"、"会议室"）来设置多个位置。会议室 B "）。</span><span class="sxs-lookup"><span data-stu-id="97c25-151">You can set multiple locations by using a semi-colon as the separator (e.g., 'Conference room A; Conference room B').</span></span>

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

## <a name="see-also"></a><span data-ttu-id="97c25-152">另请参阅</span><span class="sxs-lookup"><span data-stu-id="97c25-152">See also</span></span>

- [<span data-ttu-id="97c25-153">创建您的第一个 Outlook 外接程序</span><span class="sxs-lookup"><span data-stu-id="97c25-153">Create your first Outlook add-in</span></span>](../quickstarts/outlook-quickstart.md)
- [<span data-ttu-id="97c25-154">Office 外接程序中的异步编程</span><span class="sxs-lookup"><span data-stu-id="97c25-154">Asynchronous programming in Office Add-ins</span></span>](../develop/asynchronous-programming-in-office-add-ins.md)
