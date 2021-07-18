---
title: 为Outlook会议提供商创建一个移动外接程序
description: 讨论如何为联机会议Outlook设置移动外接程序。
ms.topic: article
ms.date: 07/09/2021
localization_priority: Normal
ms.openlocfilehash: f0f9b69c2b8b515df3829ca3ba0714393df79fd1
ms.sourcegitcommit: 30a861ece18255e342725e31c47f01960b854532
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/16/2021
ms.locfileid: "53455500"
---
# <a name="create-an-outlook-mobile-add-in-for-an-online-meeting-provider"></a><span data-ttu-id="e8400-103">为Outlook会议提供商创建一个移动外接程序</span><span class="sxs-lookup"><span data-stu-id="e8400-103">Create an Outlook mobile add-in for an online-meeting provider</span></span>

<span data-ttu-id="e8400-104">设置联机会议是 Outlook 用户的核心体验，使用移动设备轻松创建Teams会议[Outlook](/microsoftteams/teams-add-in-for-outlook)体验。</span><span class="sxs-lookup"><span data-stu-id="e8400-104">Setting up an online meeting is a core experience for an Outlook user, and it's easy to [create a Teams meeting with Outlook](/microsoftteams/teams-add-in-for-outlook) mobile.</span></span> <span data-ttu-id="e8400-105">但是，使用非 Microsoft 服务Outlook联机会议可能会很麻烦。</span><span class="sxs-lookup"><span data-stu-id="e8400-105">However, creating an online meeting in Outlook with a non-Microsoft service can be cumbersome.</span></span> <span data-ttu-id="e8400-106">通过实现此功能，服务提供商可以简化其外接程序用户的联机会议Outlook体验。</span><span class="sxs-lookup"><span data-stu-id="e8400-106">By implementing this feature, service providers can streamline the online meeting creation experience for their Outlook add-in users.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="e8400-107">此功能仅在 Android 和 iOS 上受支持，Microsoft 365订阅。</span><span class="sxs-lookup"><span data-stu-id="e8400-107">This feature is only supported on Android and iOS with a Microsoft 365 subscription.</span></span>

<span data-ttu-id="e8400-108">本文将了解如何设置 Outlook 移动外接程序，以使用户能够使用联机会议服务组织和加入会议。</span><span class="sxs-lookup"><span data-stu-id="e8400-108">In this article, you'll learn how to set up your Outlook mobile add-in to enable users to organize and join a meeting using your online-meeting service.</span></span> <span data-ttu-id="e8400-109">在整篇文章中，我们将使用虚构的联机会议服务提供商"Contoso"。</span><span class="sxs-lookup"><span data-stu-id="e8400-109">Throughout this article, we'll be using a fictional online-meeting service provider, "Contoso".</span></span>

## <a name="set-up-your-environment"></a><span data-ttu-id="e8400-110">设置环境</span><span class="sxs-lookup"><span data-stu-id="e8400-110">Set up your environment</span></span>

<span data-ttu-id="e8400-111">完成[Outlook](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator)使用适用于加载项的 Yeoman 生成器创建加载项Office快速入门。</span><span class="sxs-lookup"><span data-stu-id="e8400-111">Complete the [Outlook quick start](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) which creates an add-in project with the Yeoman generator for Office Add-ins.</span></span>

## <a name="configure-the-manifest"></a><span data-ttu-id="e8400-112">配置清单</span><span class="sxs-lookup"><span data-stu-id="e8400-112">Configure the manifest</span></span>

<span data-ttu-id="e8400-113">若要使用户能够使用外接程序创建联机会议，您必须在清单中的父元素 下配置 [MobileOnlineMeetingCommandSurface](../reference/manifest/extensionpoint.md#mobileonlinemeetingcommandsurface) 扩展点 `MobileFormFactor` 。</span><span class="sxs-lookup"><span data-stu-id="e8400-113">To enable users to create online meetings with your add-in, you must configure the [MobileOnlineMeetingCommandSurface extension point](../reference/manifest/extensionpoint.md#mobileonlinemeetingcommandsurface) in the manifest under the parent element `MobileFormFactor`.</span></span> <span data-ttu-id="e8400-114">不支持其他外形因素。</span><span class="sxs-lookup"><span data-stu-id="e8400-114">Other form factors are not supported.</span></span>

1. <span data-ttu-id="e8400-115">在代码编辑器中，打开快速启动项目。</span><span class="sxs-lookup"><span data-stu-id="e8400-115">In your code editor, open the quick start project.</span></span>

1. <span data-ttu-id="e8400-116">打开 **manifest.xml** 根目录下的文件。</span><span class="sxs-lookup"><span data-stu-id="e8400-116">Open the **manifest.xml** file located at the root of your project.</span></span>

1. <span data-ttu-id="e8400-117">选择整个节点 (包括打开和 `<VersionOverrides>` 关闭标记) 并将其替换为以下 XML。</span><span class="sxs-lookup"><span data-stu-id="e8400-117">Select the entire `<VersionOverrides>` node (including open and close tags) and replace it with the following XML.</span></span>

```xml
<VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
    <Description resid="residDescription"></Description>
    <Requirements>
      <bt:Sets>
        <bt:Set Name="Mailbox" MinVersion="1.3"/>
      </bt:Sets>
    </Requirements>
    <Hosts>
      <Host xsi:type="MailHost">
        <DesktopFormFactor>
          <FunctionFile resid="residFunctionFile"/>
          <ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
            <OfficeTab id="TabDefault">
              <Group id="apptComposeGroup">
                <Label resid="residDescription"/>
                <Control xsi:type="Button" id="insertMeetingButton">
                  <Label resid="residLabel"/>
                  <Supertip>
                    <Title resid="residLabel"/>
                    <Description resid="residTooltip"/>
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="icon-16"/>
                    <bt:Image size="32" resid="icon-32"/>
                    <bt:Image size="64" resid="icon-64"/>
                    <bt:Image size="80" resid="icon-80"/>
                  </Icon>
                  <Action xsi:type="ExecuteFunction">
                    <FunctionName>insertContosoMeeting</FunctionName>
                  </Action>
                </Control>
              </Group>
            </OfficeTab>
          </ExtensionPoint>
        </DesktopFormFactor>

        <MobileFormFactor>
          <FunctionFile resid="residFunctionFile"/>
          <ExtensionPoint xsi:type="MobileOnlineMeetingCommandSurface">
            <Control xsi:type="MobileButton" id="insertMeetingButton">
              <Label resid="residLabel"/>
              <Icon>
                <bt:Image size="25" scale="1" resid="icon-16"/>
                <bt:Image size="25" scale="2" resid="icon-16"/>
                <bt:Image size="25" scale="3" resid="icon-16"/>

                <bt:Image size="32" scale="1" resid="icon-32"/>
                <bt:Image size="32" scale="2" resid="icon-32"/>
                <bt:Image size="32" scale="3" resid="icon-32"/>

                <bt:Image size="48" scale="1" resid="icon-48"/>
                <bt:Image size="48" scale="2" resid="icon-48"/>
                <bt:Image size="48" scale="3" resid="icon-48"/>
              </Icon>
              <Action xsi:type="ExecuteFunction">
                <FunctionName>insertContosoMeeting</FunctionName>
              </Action>
            </Control>
          </ExtensionPoint>
        </MobileFormFactor>
      </Host>
    </Hosts>
    <Resources>
      <bt:Images>
        <bt:Image id="icon-16" DefaultValue="https://contoso.com/assets/icon-16.png"/>
        <bt:Image id="icon-32" DefaultValue="https://contoso.com/assets/icon-32.png"/>
        <bt:Image id="icon-48" DefaultValue="https://contoso.com/assets/icon-48.png"/>
        <bt:Image id="icon-64" DefaultValue="https://contoso.com/assets/icon-64.png"/>
        <bt:Image id="icon-80" DefaultValue="https://contoso.com/assets/icon-80.png"/>
      </bt:Images>
      <bt:Urls>
        <bt:Url id="residFunctionFile" DefaultValue="https://contoso.com/commands.html"/>
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="residDescription" DefaultValue="Contoso meeting"/>
        <bt:String id="residLabel" DefaultValue="Add a contoso meeting"/>
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="residTooltip" DefaultValue="Add a contoso meeting to this appointment."/>
      </bt:LongStrings>
    </Resources>
  </VersionOverrides>
</VersionOverrides>
```

> [!TIP]
> <span data-ttu-id="e8400-118">若要了解有关外接程序的Outlook清单，请参阅[Outlook add-in manifests](manifests.md)和[Add support for add-in commands for Outlook Mobile。](add-mobile-support.md)</span><span class="sxs-lookup"><span data-stu-id="e8400-118">To learn more about manifests for Outlook add-ins, see [Outlook add-in manifests](manifests.md) and [Add support for add-in commands for Outlook Mobile](add-mobile-support.md).</span></span>

## <a name="implement-adding-online-meeting-details"></a><span data-ttu-id="e8400-119">实现添加联机会议详细信息</span><span class="sxs-lookup"><span data-stu-id="e8400-119">Implement adding online meeting details</span></span>

<span data-ttu-id="e8400-120">在此部分中，了解外接程序脚本如何更新用户会议以包含联机会议详细信息。</span><span class="sxs-lookup"><span data-stu-id="e8400-120">In this section, learn how your add-in script can update a user's meeting to include online meeting details.</span></span>

1. <span data-ttu-id="e8400-121">从同一快速启动项目中，在代码编辑器中打开 **commands.js./src/commands/commands.js** 文件。</span><span class="sxs-lookup"><span data-stu-id="e8400-121">From the same quick start project, open the file **./src/commands/commands.js** in your code editor.</span></span>

1. <span data-ttu-id="e8400-122">用以下 JavaScript **commands.js** 文件的全部内容。</span><span class="sxs-lookup"><span data-stu-id="e8400-122">Replace the entire content of the **commands.js** file with the following JavaScript.</span></span>

    ```js
    // 1. How to construct online meeting details.
    // Not shown: How to get the meeting organizer's ID and other details from your service.
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

    var mailboxItem;

    // Office is ready.
    Office.onReady(function () {
            mailboxItem = Office.context.mailbox.item;
        }
    );

    // 2. How to define a UI-less function named `insertContosoMeeting` (referenced in the manifest)
    //    to update the meeting body with the online meeting details.
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

    // 3. How to implement a supporting function `updateBody`
    //    that appends the online meeting details to the current body of the meeting.
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

    function getGlobal() {
      return typeof self !== "undefined"
        ? self
        : typeof window !== "undefined"
        ? window
        : typeof global !== "undefined"
        ? global
        : undefined;
    }

    const g = getGlobal();

    // The add-in command functions need to be available in global scope.
    g.insertContosoMeeting = insertContosoMeeting;
    ```

## <a name="testing-and-validation"></a><span data-ttu-id="e8400-123">测试和验证</span><span class="sxs-lookup"><span data-stu-id="e8400-123">Testing and validation</span></span>

<span data-ttu-id="e8400-124">按照常规指南 [测试和验证加载项](testing-and-tips.md)。</span><span class="sxs-lookup"><span data-stu-id="e8400-124">Follow the usual guidance to [test and validate your add-in](testing-and-tips.md).</span></span> <span data-ttu-id="e8400-125">在[Outlook 网页版、Windows](sideload-outlook-add-ins-for-testing.md)或 Mac 中旁加载后，Outlook Android 或 iOS 移动设备上重新启动。</span><span class="sxs-lookup"><span data-stu-id="e8400-125">After [sideloading](sideload-outlook-add-ins-for-testing.md) in Outlook on the web, Windows, or Mac, restart Outlook on your Android or iOS mobile device.</span></span> <span data-ttu-id="e8400-126">然后，在新的会议屏幕上，确认 Microsoft Teams 或 Skype 开关已替换为你自己的开关。</span><span class="sxs-lookup"><span data-stu-id="e8400-126">Then, on a new meeting screen, verify that the Microsoft Teams or Skype toggle is replaced with your own.</span></span>

### <a name="create-meeting-ui"></a><span data-ttu-id="e8400-127">创建会议 UI</span><span class="sxs-lookup"><span data-stu-id="e8400-127">Create meeting UI</span></span>

<span data-ttu-id="e8400-128">作为会议组织者，您应该在创建会议时看到类似于以下三个图像的屏幕。</span><span class="sxs-lookup"><span data-stu-id="e8400-128">As a meeting organizer, you should see screens similar to the following three images when you create a meeting.</span></span>

<span data-ttu-id="e8400-129">[![在 Android - Contoso 上创建会议屏幕切换为关闭。](../images/outlook-android-create-online-meeting-off.png)](../images/outlook-android-create-online-meeting-off-expanded.png#lightbox)</span><span class="sxs-lookup"><span data-stu-id="e8400-129">[![The create meeting screen on Android - Contoso toggle off.](../images/outlook-android-create-online-meeting-off.png)](../images/outlook-android-create-online-meeting-off-expanded.png#lightbox)</span></span> <span data-ttu-id="e8400-130">[![在 Android 上创建会议屏幕 - 加载 Contoso 切换。](../images/outlook-android-create-online-meeting-load.png)](../images/outlook-android-create-online-meeting-load-expanded.png#lightbox)</span><span class="sxs-lookup"><span data-stu-id="e8400-130">[![The create meeting screen on Android - loading Contoso toggle.](../images/outlook-android-create-online-meeting-load.png)](../images/outlook-android-create-online-meeting-load-expanded.png#lightbox)</span></span> <span data-ttu-id="e8400-131">[![在 Android 上创建会议屏幕 - Contoso 切换打开。](../images/outlook-android-create-online-meeting-on.png)](../images/outlook-android-create-online-meeting-on-expanded.png#lightbox)</span><span class="sxs-lookup"><span data-stu-id="e8400-131">[![The create meeting screen on Android - Contoso toggle on.](../images/outlook-android-create-online-meeting-on.png)](../images/outlook-android-create-online-meeting-on-expanded.png#lightbox)</span></span>

### <a name="join-meeting-ui"></a><span data-ttu-id="e8400-132">加入会议 UI</span><span class="sxs-lookup"><span data-stu-id="e8400-132">Join meeting UI</span></span>

<span data-ttu-id="e8400-133">作为与会者，在查看会议时应该会看到类似于下图的屏幕。</span><span class="sxs-lookup"><span data-stu-id="e8400-133">As a meeting attendee, you should see a screen similar to the following image when you view the meeting.</span></span>

<span data-ttu-id="e8400-134">[![Android 上加入会议屏幕的屏幕截图。](../images/outlook-android-join-online-meeting-view-1.png)](../images/outlook-android-join-online-meeting-view-1-expanded.png#lightbox)</span><span class="sxs-lookup"><span data-stu-id="e8400-134">[![Screenshot of join meeting screen on Android.](../images/outlook-android-join-online-meeting-view-1.png)](../images/outlook-android-join-online-meeting-view-1-expanded.png#lightbox)</span></span>

> [!IMPORTANT]
> <span data-ttu-id="e8400-135">如果未看到"加入"链接，可能是你的服务的联机会议模板未在我们的服务器上注册。</span><span class="sxs-lookup"><span data-stu-id="e8400-135">If you don't see the **Join** link, it may be that the online-meeting template for your service is not registered on our servers.</span></span> <span data-ttu-id="e8400-136">有关详细信息 [，请参阅注册联机会议](#register-your-online-meeting-template) 模板部分。</span><span class="sxs-lookup"><span data-stu-id="e8400-136">See the [Register your online-meeting template](#register-your-online-meeting-template) section for details.</span></span>

## <a name="register-your-online-meeting-template"></a><span data-ttu-id="e8400-137">注册联机会议模板</span><span class="sxs-lookup"><span data-stu-id="e8400-137">Register your online-meeting template</span></span>

<span data-ttu-id="e8400-138">如果要为服务注册联机会议模板，可以创建一个GitHub问题。</span><span class="sxs-lookup"><span data-stu-id="e8400-138">If you would like to register the online-meeting template for your service, you can create a GitHub issue with the details.</span></span> <span data-ttu-id="e8400-139">之后，我们将联系你以协调注册时间线。</span><span class="sxs-lookup"><span data-stu-id="e8400-139">After that, we'll contact you to coordinate registration timeline.</span></span>

1. <span data-ttu-id="e8400-140">转到 **本文** 末尾的"反馈"部分。</span><span class="sxs-lookup"><span data-stu-id="e8400-140">Go to the **Feedback** section at the end of this article.</span></span>
1. <span data-ttu-id="e8400-141">按" **此页面"** 链接。</span><span class="sxs-lookup"><span data-stu-id="e8400-141">Press the **This page** link.</span></span>
1. <span data-ttu-id="e8400-142">将 **新问题** 的标题设置为"为 my-service 注册联机会议模板"，并 `my-service` 替换为你的服务名称。</span><span class="sxs-lookup"><span data-stu-id="e8400-142">Set the **Title** of the new issue to "Register the online-meeting template for my-service", replacing `my-service` with your service name.</span></span>
1. <span data-ttu-id="e8400-143">在问题正文中，将字符串"[在此处输入反馈]"替换为你在本文前面实现添加联机会议详细信息部分或类似变量中设置的 `newBody` 字符串。 [](#implement-adding-online-meeting-details)</span><span class="sxs-lookup"><span data-stu-id="e8400-143">In the issue body, replace the string "[Enter feedback here]" with the string you set in the `newBody` or similar variable from the [Implement adding online meeting details](#implement-adding-online-meeting-details) section earlier in this article.</span></span>
1. <span data-ttu-id="e8400-144">单击 **"提交新问题"。**</span><span class="sxs-lookup"><span data-stu-id="e8400-144">Click **Submit new issue**.</span></span>

![包含 Contoso GitHub新问题屏幕的屏幕截图。](../images/outlook-request-to-register-online-meeting-template.png)

## <a name="available-apis"></a><span data-ttu-id="e8400-146">可用的 API</span><span class="sxs-lookup"><span data-stu-id="e8400-146">Available APIs</span></span>

<span data-ttu-id="e8400-147">以下 API 可用于此功能。</span><span class="sxs-lookup"><span data-stu-id="e8400-147">The following APIs are available for this feature.</span></span>

- <span data-ttu-id="e8400-148">约会管理器 API</span><span class="sxs-lookup"><span data-stu-id="e8400-148">Appointment Organizer APIs</span></span>
  - <span data-ttu-id="e8400-149">[Office.context.mailbox.item.body](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#body) ([Body.getAsync](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#getasync-coerciontype--options--callback-)、 [Body.setAsync](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#setasync-data--options--callback-)) </span><span class="sxs-lookup"><span data-stu-id="e8400-149">[Office.context.mailbox.item.body](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#body) ([Body.getAsync](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#getasync-coerciontype--options--callback-), [Body.setAsync](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#setasync-data--options--callback-))</span></span>
  - <span data-ttu-id="e8400-150">[Office.context.mailbox.item.end (](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#end) [Time](/javascript/api/outlook/office.time?view=outlook-js-preview&preserve-view=true)) </span><span class="sxs-lookup"><span data-stu-id="e8400-150">[Office.context.mailbox.item.end](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#end) ([Time](/javascript/api/outlook/office.time?view=outlook-js-preview&preserve-view=true))</span></span>
  - <span data-ttu-id="e8400-151">[Office.context.mailbox.item.loadCustomPropertiesAsync](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#loadcustompropertiesasync-callback--usercontext-) ([CustomProperties](/javascript/api/outlook/office.customproperties?view=outlook-js-preview&preserve-view=true)) </span><span class="sxs-lookup"><span data-stu-id="e8400-151">[Office.context.mailbox.item.loadCustomPropertiesAsync](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#loadcustompropertiesasync-callback--usercontext-) ([CustomProperties](/javascript/api/outlook/office.customproperties?view=outlook-js-preview&preserve-view=true))</span></span>
  - <span data-ttu-id="e8400-152">[Office.context.mailbox.item.location (](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#location) [Location) ](/javascript/api/outlook/office.location?view=outlook-js-preview&preserve-view=true)</span><span class="sxs-lookup"><span data-stu-id="e8400-152">[Office.context.mailbox.item.location](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#location) ([Location](/javascript/api/outlook/office.location?view=outlook-js-preview&preserve-view=true))</span></span>
  - <span data-ttu-id="e8400-153">[Office.context.mailbox.item.optionalAttendees](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#optionalattendees) ([Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-preview&preserve-view=true)) </span><span class="sxs-lookup"><span data-stu-id="e8400-153">[Office.context.mailbox.item.optionalAttendees](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#optionalattendees) ([Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-preview&preserve-view=true))</span></span>
  - <span data-ttu-id="e8400-154">[Office.context.mailbox.item.requiredAttendees](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#requiredattendees) ([Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-preview&preserve-view=true)) </span><span class="sxs-lookup"><span data-stu-id="e8400-154">[Office.context.mailbox.item.requiredAttendees](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#requiredattendees) ([Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-preview&preserve-view=true))</span></span>
  - <span data-ttu-id="e8400-155">[Office.context.mailbox.item.start (](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#start) [Time](/javascript/api/outlook/office.time?view=outlook-js-preview&preserve-view=true)) </span><span class="sxs-lookup"><span data-stu-id="e8400-155">[Office.context.mailbox.item.start](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#start) ([Time](/javascript/api/outlook/office.time?view=outlook-js-preview&preserve-view=true))</span></span>
  - <span data-ttu-id="e8400-156">[Office.context.mailbox.item.subject (](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#subject) [Subject) ](/javascript/api/outlook/office.subject?view=outlook-js-preview&preserve-view=true)</span><span class="sxs-lookup"><span data-stu-id="e8400-156">[Office.context.mailbox.item.subject](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#subject) ([Subject](/javascript/api/outlook/office.subject?view=outlook-js-preview&preserve-view=true))</span></span>
  - <span data-ttu-id="e8400-157">[Office.context.roamingSettings](../reference/objectmodel/preview-requirement-set/office.context.md?view=outlook-js-preview&preserve-view=true#roamingsettings-roamingsettings) ([RoamingSettings) ](/javascript/api/outlook/office.roamingsettings?view=outlook-js-preview&preserve-view=true)</span><span class="sxs-lookup"><span data-stu-id="e8400-157">[Office.context.roamingSettings](../reference/objectmodel/preview-requirement-set/office.context.md?view=outlook-js-preview&preserve-view=true#roamingsettings-roamingsettings) ([RoamingSettings](/javascript/api/outlook/office.roamingsettings?view=outlook-js-preview&preserve-view=true))</span></span>
- <span data-ttu-id="e8400-158">处理身份验证流</span><span class="sxs-lookup"><span data-stu-id="e8400-158">Handle auth flow</span></span>
  - [<span data-ttu-id="e8400-159">Dialog API</span><span class="sxs-lookup"><span data-stu-id="e8400-159">Dialog APIs</span></span>](../develop/dialog-api-in-office-add-ins.md)

## <a name="restrictions"></a><span data-ttu-id="e8400-160">限制</span><span class="sxs-lookup"><span data-stu-id="e8400-160">Restrictions</span></span>

<span data-ttu-id="e8400-161">有几个限制适用。</span><span class="sxs-lookup"><span data-stu-id="e8400-161">Several restrictions apply.</span></span>

- <span data-ttu-id="e8400-162">仅适用于联机会议服务提供商。</span><span class="sxs-lookup"><span data-stu-id="e8400-162">Applicable only to online-meeting service providers.</span></span>
- <span data-ttu-id="e8400-163">只有管理员安装的外接程序将显示在会议撰写屏幕上，以替换默认的Teams或Skype选项。</span><span class="sxs-lookup"><span data-stu-id="e8400-163">Only admin-installed add-ins will appear on the meeting compose screen, replacing the default Teams or Skype option.</span></span> <span data-ttu-id="e8400-164">用户安装的加载项不会激活。</span><span class="sxs-lookup"><span data-stu-id="e8400-164">User-installed add-ins won't activate.</span></span>
- <span data-ttu-id="e8400-165">加载项图标应该使用十六进制代码或其他颜色格式的等效 `#919191` 项以灰 [度显示](https://convertingcolors.com/hex-color-919191.html)。</span><span class="sxs-lookup"><span data-stu-id="e8400-165">The add-in icon should be in grayscale using hex code `#919191` or its equivalent in [other color formats](https://convertingcolors.com/hex-color-919191.html).</span></span>
- <span data-ttu-id="e8400-166">在以撰写模式撰写的约会管理器中，仅 (UI) 命令。</span><span class="sxs-lookup"><span data-stu-id="e8400-166">Only one UI-less command is supported in Appointment Organizer (compose) mode.</span></span>
- <span data-ttu-id="e8400-167">外接程序应在一分钟的超时时间内更新约会窗体中的会议详细信息。</span><span class="sxs-lookup"><span data-stu-id="e8400-167">The add-in should update the meeting details in the appointment form within the one-minute timeout period.</span></span> <span data-ttu-id="e8400-168">但是，在对话框中为进行身份验证而打开的外接程序等所花的任何时间都从超时期间排除。</span><span class="sxs-lookup"><span data-stu-id="e8400-168">However, any time spent in a dialog box the add-in opened for authentication, etc. is excluded from the timeout period.</span></span>

## <a name="see-also"></a><span data-ttu-id="e8400-169">另请参阅</span><span class="sxs-lookup"><span data-stu-id="e8400-169">See also</span></span>

- [<span data-ttu-id="e8400-170">适用于 Outlook Mobile 的加载项</span><span class="sxs-lookup"><span data-stu-id="e8400-170">Add-ins for Outlook Mobile</span></span>](outlook-mobile-addins.md)
- [<span data-ttu-id="e8400-171">添加对适用于 Outlook Mobile 的外接程序命令的支持</span><span class="sxs-lookup"><span data-stu-id="e8400-171">Add support for add-in commands for Outlook Mobile</span></span>](add-mobile-support.md)
