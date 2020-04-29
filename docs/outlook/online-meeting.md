---
title: 为联机会议提供商创建 Outlook mobile 外接程序（预览）
description: 讨论如何为联机会议服务提供商设置 Outlook 移动外接程序。
ms.topic: article
ms.date: 04/23/2020
localization_priority: Normal
ms.openlocfilehash: 8a54ddf96ca2b5e697198b4bc69b2ec5abee10d1
ms.sourcegitcommit: 0fdb78cefa669b727b817614a4147a46d249a0ed
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/28/2020
ms.locfileid: "43930321"
---
# <a name="create-an-outlook-mobile-add-in-for-an-online-meeting-provider-preview"></a><span data-ttu-id="6560c-103">为联机会议提供商创建 Outlook mobile 外接程序（预览）</span><span class="sxs-lookup"><span data-stu-id="6560c-103">Create an Outlook mobile add-in for an online-meeting provider (preview)</span></span>

<span data-ttu-id="6560c-104">设置联机会议是 Outlook 用户的核心体验，可轻松[创建使用 outlook mobile 的团队会议](/microsoftteams/teams-add-in-for-outlook)。</span><span class="sxs-lookup"><span data-stu-id="6560c-104">Setting up an online meeting is a core experience for an Outlook user, and it's easy to [create a Teams meeting with Outlook](/microsoftteams/teams-add-in-for-outlook) mobile.</span></span> <span data-ttu-id="6560c-105">但是，在 Outlook 中使用非 Microsoft 服务创建联机会议可能很麻烦。</span><span class="sxs-lookup"><span data-stu-id="6560c-105">However, creating an online meeting in Outlook with a non-Microsoft service can be cumbersome.</span></span> <span data-ttu-id="6560c-106">通过实施此功能，服务提供商可以为其 Outlook 外接程序用户简化联机会议创建体验。</span><span class="sxs-lookup"><span data-stu-id="6560c-106">By implementing this feature, service providers can streamline the online meeting creation experience for their Outlook add-in users.</span></span>

> [!NOTE]
> <span data-ttu-id="6560c-107">只有 Office 365 订阅的 Android 中的[预览](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md)才支持此功能。</span><span class="sxs-lookup"><span data-stu-id="6560c-107">This feature is only supported in [preview](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) on Android with an Office 365 subscription.</span></span>

<span data-ttu-id="6560c-108">在本文中，您将了解如何设置 Outlook 移动外接程序，以使用户能够使用您的联机会议服务来组织和加入会议。</span><span class="sxs-lookup"><span data-stu-id="6560c-108">In this article, you'll learn how to set up your Outlook mobile add-in to enable users to organize and join a meeting using your online-meeting service.</span></span> <span data-ttu-id="6560c-109">在整篇文章中，我们将使用虚构的联机会议服务提供商 "Contoso"。</span><span class="sxs-lookup"><span data-stu-id="6560c-109">Throughout this article, we'll be using a fictional online-meeting service provider, "Contoso".</span></span>

## <a name="configure-the-manifest"></a><span data-ttu-id="6560c-110">配置清单</span><span class="sxs-lookup"><span data-stu-id="6560c-110">Configure the manifest</span></span>

<span data-ttu-id="6560c-111">若要使用户能够使用您的外接程序创建联机会议，必须在`MobileOnlineMeetingCommandSurface`父元素`MobileFormFactor`下的清单中配置扩展点。</span><span class="sxs-lookup"><span data-stu-id="6560c-111">To enable users to create online meetings with your add-in, you must configure the `MobileOnlineMeetingCommandSurface` extension point in the manifest under the parent element `MobileFormFactor`.</span></span> <span data-ttu-id="6560c-112">不支持其他外观因素。</span><span class="sxs-lookup"><span data-stu-id="6560c-112">Other form factors are not supported.</span></span>

<span data-ttu-id="6560c-113">下面的示例展示了包含`MobileFormFactor`元素和`MobileOnlineMeetingCommandSurface`扩展点的清单中的摘录。</span><span class="sxs-lookup"><span data-stu-id="6560c-113">The following example shows an excerpt from the manifest that includes the `MobileFormFactor` element and `MobileOnlineMeetingCommandSurface` extension point.</span></span>

> [!TIP]
> <span data-ttu-id="6560c-114">若要了解有关 Outlook 外接程序清单的详细信息，请参阅[outlook 外接程序清单](manifests.md)和[添加对适用于 outlook Mobile 的外接程序命令的支持](add-mobile-support.md)。</span><span class="sxs-lookup"><span data-stu-id="6560c-114">To learn more about manifests for Outlook add-ins, see [Outlook add-in manifests](manifests.md) and [Add support for add-in commands for Outlook Mobile](add-mobile-support.md).</span></span>

```xml
...
<VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
    ...
    <Hosts>
      <Host xsi:type="MailHost">
        <MobileFormFactor>
          <FunctionFile resid="residMobileFuncUrl" />
          <ExtensionPoint xsi:type="MobileOnlineMeetingCommandSurface">
            <!-- Configure selected extension point. -->
            <Control xsi:type="MobileButton" id="onlineMeetingFunctionButton">
              <Label resid="residUILessButton0Name" />
              <Icon>
                <bt:Image resid="UiLessIcon" size="25" scale="1" />
                <bt:Image resid="UiLessIcon" size="25" scale="2" />
                <bt:Image resid="UiLessIcon" size="25" scale="3" />
                <bt:Image resid="UiLessIcon" size="32" scale="1" />
                <bt:Image resid="UiLessIcon" size="32" scale="2" />
                <bt:Image resid="UiLessIcon" size="32" scale="3" />
                <bt:Image resid="UiLessIcon" size="48" scale="1" />
                <bt:Image resid="UiLessIcon" size="48" scale="2" />
                <bt:Image resid="UiLessIcon" size="48" scale="3" />
              </Icon>
              <Action xsi:type="ExecuteFunction">
                <FunctionName>insertContosoMeeting</FunctionName>
              </Action>
            </Control>
          </ExtensionPoint>
        </MobileFormFactor>
      </Host>
    </Hosts>
    ...
  </VersionOverrides>
</VersionOverrides>
...
```

## <a name="implement-adding-online-meeting-details"></a><span data-ttu-id="6560c-115">实施添加联机会议详细信息</span><span class="sxs-lookup"><span data-stu-id="6560c-115">Implement adding online meeting details</span></span>

<span data-ttu-id="6560c-116">在本节中，了解外接程序脚本如何更新用户的会议以包含联机会议详细信息。</span><span class="sxs-lookup"><span data-stu-id="6560c-116">In this section, learn how your add-in script can update a user's meeting to include online meeting details.</span></span>

<span data-ttu-id="6560c-117">下面的示例展示了如何构建联机会议详细信息。</span><span class="sxs-lookup"><span data-stu-id="6560c-117">The following example shows how you construct online meeting details.</span></span> <span data-ttu-id="6560c-118">不显示：如何从服务中获取会议组织者的 ID 和其他详细信息。</span><span class="sxs-lookup"><span data-stu-id="6560c-118">Not shown is how to get the meeting organizer's ID and other details from your service.</span></span>

```js
const newBody = '<br>' +
    '<a href="https://contoso.com/meeting?id=123456789" target="_blank">Join Contoso meeting</a>' +
    '<br><br>' +
    'Phone Dial-in: +1(123)456-7890' +
    '<br><br>' +
    'Meeting ID: 123 456 789' +
    '<br><br>' +
    'Want to test your video connection?' +
    '<br><br>' +
    '<a href="https://contoso.com/testmeeting" target="_blank">Join test meeting</a>' +
    '<br><br>';
```

<span data-ttu-id="6560c-119">下面的示例演示如何定义清单中`insertContosoMeeting`引用的无 UI 的函数，以使用联机会议详细信息更新会议正文。</span><span class="sxs-lookup"><span data-stu-id="6560c-119">The following example shows how to define a UI-less function named `insertContosoMeeting` referenced in the manifest to update the meeting body with the online meeting details.</span></span>

```js
var mailboxItem;

// Office is ready.
Office.onReady(function () {
        mailboxItem = Office.context.mailbox.item;
    }
);

function insertContosoMeeting(event) {
    // Get HTML body from the client.
    mailboxItem.body.getAsync("html",
        { asyncContext: event },
        function (getBodyResult) {
            if (getBodyResult.status === Office.AsyncResultStatus.Succeeded) {
                updateBody(getBodyResult.asyncContext, getBodyResult.value);
            } else {
                console.error("Failed to get HTML body.");
                getBodyResult.asyncContext.completed({ allowEvent: false });
            }
        }
    );
}
```

<span data-ttu-id="6560c-120">下面的示例演示在上一示例中使用`updateBody`的支持函数的实现，该示例将联机会议详细信息追加到会议的当前正文。</span><span class="sxs-lookup"><span data-stu-id="6560c-120">The following example shows an implementation of the supporting function `updateBody` used in the previous example that appends the online meeting details to the current body of the meeting.</span></span>

```js
function updateBody(event, existingBody) {
    // Append new body to the existing body.
    mailboxItem.body.setAsync(existingBody + newBody,
        { asyncContext: event, coercionType: "html" },
        function (setBodyResult) {
            if (setBodyResult.status === Office.AsyncResultStatus.Succeeded) {
                setBodyResult.asyncContext.completed({ allowEvent: true });
            } else {
                console.error("Failed to set HTML body.");
                setBodyResult.asyncContext.completed({ allowEvent: false });
            }
        }
    );
}
```

## <a name="testing-and-validation"></a><span data-ttu-id="6560c-121">测试和验证</span><span class="sxs-lookup"><span data-stu-id="6560c-121">Testing and validation</span></span>

<span data-ttu-id="6560c-122">按照通常的指导来[测试和验证您的外接程序](testing-and-tips.md)。</span><span class="sxs-lookup"><span data-stu-id="6560c-122">Follow the usual guidance to [test and validate your add-in](testing-and-tips.md).</span></span> <span data-ttu-id="6560c-123">在 Outlook 网页版、Windows 版或 Mac 版中进行[旁加载](sideload-outlook-add-ins-for-testing.md)后，在你的 android 移动设备上重新启动 outlook （现在，android 是唯一受支持的客户端）。</span><span class="sxs-lookup"><span data-stu-id="6560c-123">After [sideloading](sideload-outlook-add-ins-for-testing.md) in Outlook on the web, Windows, or Mac, restart Outlook on your Android mobile device (Android is the only supported client for now).</span></span> <span data-ttu-id="6560c-124">然后，在新的会议屏幕上，验证 Microsoft 团队或 Skype 切换是否已替换为你自己的。</span><span class="sxs-lookup"><span data-stu-id="6560c-124">Then, on a new meeting screen, verify that the Microsoft Teams or Skype toggle is replaced with your own.</span></span>

### <a name="create-meeting-ui"></a><span data-ttu-id="6560c-125">创建会议用户界面</span><span class="sxs-lookup"><span data-stu-id="6560c-125">Create meeting UI</span></span>

<span data-ttu-id="6560c-126">作为会议组织者，在创建会议时，您应看到类似于以下三幅图像的屏幕。</span><span class="sxs-lookup"><span data-stu-id="6560c-126">As a meeting organizer, you should see screens similar to the following three images when you create a meeting.</span></span>

<span data-ttu-id="6560c-127">在 android 上创建会议屏幕的[ ![](../images/outlook-android-create-online-meeting-load.png)](../images/outlook-android-create-online-meeting-load-expanded.png#lightbox) [屏幕截图-contoso 切换关闭在 android 上创建会议屏幕的屏幕截图打开在 android 上创建会议屏幕的屏幕![](../images/outlook-android-create-online-meeting-off.png)](../images/outlook-android-create-online-meeting-off-expanded.png#lightbox)截图[ ![-contoso 切换](../images/outlook-android-create-online-meeting-on.png)](../images/outlook-android-create-online-meeting-on-expanded.png#lightbox)</span><span class="sxs-lookup"><span data-stu-id="6560c-127">[![screenshot of create meeting screen on Android - Contoso toggle off](../images/outlook-android-create-online-meeting-off.png)](../images/outlook-android-create-online-meeting-off-expanded.png#lightbox) [![screenshot of create meeting screen on Android - loading Contoso toggle](../images/outlook-android-create-online-meeting-load.png)](../images/outlook-android-create-online-meeting-load-expanded.png#lightbox) [![screenshot of create meeting screen on Android - Contoso toggle on](../images/outlook-android-create-online-meeting-on.png)](../images/outlook-android-create-online-meeting-on-expanded.png#lightbox)</span></span>

### <a name="join-meeting-ui"></a><span data-ttu-id="6560c-128">加入会议 UI</span><span class="sxs-lookup"><span data-stu-id="6560c-128">Join meeting UI</span></span>

<span data-ttu-id="6560c-129">作为会议与会者，查看会议时，您应该会看到类似于下图的屏幕。</span><span class="sxs-lookup"><span data-stu-id="6560c-129">As a meeting attendee, you should see a screen similar to the following image when you view the meeting.</span></span>

<span data-ttu-id="6560c-130">[![Android 上的加入会议屏幕的屏幕截图](../images/outlook-android-join-online-meeting-view-1.png)](../images/outlook-android-join-online-meeting-view-1-expanded.png#lightbox)</span><span class="sxs-lookup"><span data-stu-id="6560c-130">[![screenshot of join meeting screen on Android](../images/outlook-android-join-online-meeting-view-1.png)](../images/outlook-android-join-online-meeting-view-1-expanded.png#lightbox)</span></span>

## <a name="available-apis"></a><span data-ttu-id="6560c-131">可用 Api</span><span class="sxs-lookup"><span data-stu-id="6560c-131">Available APIs</span></span>

<span data-ttu-id="6560c-132">以下 Api 可用于此功能。</span><span class="sxs-lookup"><span data-stu-id="6560c-132">The following APIs are available for this feature.</span></span>

- <span data-ttu-id="6560c-133">约会组织者 Api</span><span class="sxs-lookup"><span data-stu-id="6560c-133">Appointment Organizer APIs</span></span>
  - <span data-ttu-id="6560c-134">" [Context.subname](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#subject) " （[subject](/javascript/api/outlook/office.subject?view=outlook-js-preview)）</span><span class="sxs-lookup"><span data-stu-id="6560c-134">[Office.context.mailbox.item.subject](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#subject) ([Subject](/javascript/api/outlook/office.subject?view=outlook-js-preview))</span></span>
  - <span data-ttu-id="6560c-135">"Context.subname" （Time）。[开始](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#start)（[Time](/javascript/api/outlook/office.time?view=outlook-js-preview)）</span><span class="sxs-lookup"><span data-stu-id="6560c-135">[Office.context.mailbox.item.start](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#start) ([Time](/javascript/api/outlook/office.time?view=outlook-js-preview))</span></span>
  - <span data-ttu-id="6560c-136">（Time）[结尾](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#end)（[Time](/javascript/api/outlook/office.time?view=outlook-js-preview)）</span><span class="sxs-lookup"><span data-stu-id="6560c-136">[Office.context.mailbox.item.end](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#end) ([Time](/javascript/api/outlook/office.time?view=outlook-js-preview))</span></span>
  - <span data-ttu-id="6560c-137">" [Context.subname](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#location) " （"[位置](/javascript/api/outlook/office.location?view=outlook-js-preview)"）</span><span class="sxs-lookup"><span data-stu-id="6560c-137">[Office.context.mailbox.item.location](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#location) ([Location](/javascript/api/outlook/office.location?view=outlook-js-preview))</span></span>
  - <span data-ttu-id="6560c-138">[OptionalAttendees](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#optionalattendees) （[收件人](/javascript/api/outlook/office.recipients?view=outlook-js-preview)）的内容</span><span class="sxs-lookup"><span data-stu-id="6560c-138">[Office.context.mailbox.item.optionalAttendees](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#optionalattendees) ([Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-preview))</span></span>
  - <span data-ttu-id="6560c-139">[RequiredAttendees](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#requiredattendees) （[收件人](/javascript/api/outlook/office.recipients?view=outlook-js-preview)）的内容</span><span class="sxs-lookup"><span data-stu-id="6560c-139">[Office.context.mailbox.item.requiredAttendees](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#requiredattendees) ([Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-preview))</span></span>
  - <span data-ttu-id="6560c-140">"[GetAsync](/javascript/api/outlook/office.body?view=outlook-js-preview#getasync-coerciontype--options--callback-) [" （"."](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#body) 、"setAsync"、"Body"、" [Body.setAsync](/javascript/api/outlook/office.body?view=outlook-js-preview#setasync-data--options--callback-)"）</span><span class="sxs-lookup"><span data-stu-id="6560c-140">[Office.context.mailbox.item.body](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#body) ([Body.getAsync](/javascript/api/outlook/office.body?view=outlook-js-preview#getasync-coerciontype--options--callback-), [Body.setAsync](/javascript/api/outlook/office.body?view=outlook-js-preview#setasync-data--options--callback-))</span></span>
  - <span data-ttu-id="6560c-141">[LoadCustomPropertiesAsync](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#loadcustompropertiesasync-callback--usercontext-) （[CustomProperties](/javascript/api/outlook/office.customproperties?view=outlook-js-preview)）的内容</span><span class="sxs-lookup"><span data-stu-id="6560c-141">[Office.context.mailbox.item.loadCustomPropertiesAsync](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#loadcustompropertiesasync-callback--usercontext-) ([CustomProperties](/javascript/api/outlook/office.customproperties?view=outlook-js-preview))</span></span>
  - <span data-ttu-id="6560c-142">[RoamingSettings](../reference/objectmodel/preview-requirement-set/office.context.md?view=outlook-js-preview#roamingsettings-roamingsettings) （[roamingSettings](/javascript/api/outlook/office.roamingsettings?view=outlook-js-preview)）</span><span class="sxs-lookup"><span data-stu-id="6560c-142">[Office.context.roamingSettings](../reference/objectmodel/preview-requirement-set/office.context.md?view=outlook-js-preview#roamingsettings-roamingsettings) ([RoamingSettings](/javascript/api/outlook/office.roamingsettings?view=outlook-js-preview))</span></span>
- <span data-ttu-id="6560c-143">处理身份验证流</span><span class="sxs-lookup"><span data-stu-id="6560c-143">Handle auth flow</span></span>
  - [<span data-ttu-id="6560c-144">Dialog API</span><span class="sxs-lookup"><span data-stu-id="6560c-144">Dialog APIs</span></span>](../develop/dialog-api-in-office-add-ins.md)

## <a name="restrictions"></a><span data-ttu-id="6560c-145">条件</span><span class="sxs-lookup"><span data-stu-id="6560c-145">Restrictions</span></span>

<span data-ttu-id="6560c-146">应用了多个限制。</span><span class="sxs-lookup"><span data-stu-id="6560c-146">Several restrictions apply.</span></span>

- <span data-ttu-id="6560c-147">仅适用于联机会议服务提供商。</span><span class="sxs-lookup"><span data-stu-id="6560c-147">Applicable only to online-meeting service providers.</span></span>
- <span data-ttu-id="6560c-148">当前在预览中，因此不应在生产外接程序中使用此功能。</span><span class="sxs-lookup"><span data-stu-id="6560c-148">Currently in preview so this feature shouldn't be used in production add-ins.</span></span>
- <span data-ttu-id="6560c-149">目前，Android 是唯一受支持的客户端。</span><span class="sxs-lookup"><span data-stu-id="6560c-149">At present, Android is the only supported client.</span></span> <span data-ttu-id="6560c-150">即将推出对 iOS 的支持。</span><span class="sxs-lookup"><span data-stu-id="6560c-150">Support on iOS is coming soon.</span></span>
- <span data-ttu-id="6560c-151">只有管理员安装的加载项才会显示在会议撰写屏幕上，替换默认团队或 Skype 选项。</span><span class="sxs-lookup"><span data-stu-id="6560c-151">Only admin-installed add-ins will appear on the meeting compose screen, replacing the default Teams or Skype option.</span></span> <span data-ttu-id="6560c-152">无法激活用户安装的外接程序。</span><span class="sxs-lookup"><span data-stu-id="6560c-152">User-installed add-ins won't activate.</span></span>
- <span data-ttu-id="6560c-153">外接端图标应使用十六进制代码`#919191`或以[其他颜色格式](https://convertingcolors.com/hex-color-919191.html)的等效项进行灰度。</span><span class="sxs-lookup"><span data-stu-id="6560c-153">The add-in icon should be in grayscale using hex code `#919191` or its equivalent in [other color formats](https://convertingcolors.com/hex-color-919191.html).</span></span>
- <span data-ttu-id="6560c-154">在约会组织者（撰写）模式下仅支持一个无 UI 的命令。</span><span class="sxs-lookup"><span data-stu-id="6560c-154">Only one UI-less command is supported in Appointment Organizer (compose) mode.</span></span>

## <a name="see-also"></a><span data-ttu-id="6560c-155">另请参阅</span><span class="sxs-lookup"><span data-stu-id="6560c-155">See also</span></span>

- [<span data-ttu-id="6560c-156">适用于 Outlook Mobile 的加载项</span><span class="sxs-lookup"><span data-stu-id="6560c-156">Add-ins for Outlook Mobile</span></span>](outlook-mobile-addins.md)
- [<span data-ttu-id="6560c-157">添加对适用于 Outlook Mobile 的外接程序命令的支持</span><span class="sxs-lookup"><span data-stu-id="6560c-157">Add support for add-in commands for Outlook Mobile</span></span>](add-mobile-support.md)
