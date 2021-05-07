---
title: '配置Outlook加载项进行基于事件的激活和 (预览) '
description: 了解如何配置Outlook加载项进行基于事件的激活。
ms.topic: article
ms.date: 05/04/2021
localization_priority: Normal
ms.openlocfilehash: 0052f08e9c6a3903f4adb48efca3ff29a6d21467
ms.sourcegitcommit: 8fbc7c7eb47875bf022e402b13858695a8536ec5
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/06/2021
ms.locfileid: "52253318"
---
# <a name="configure-your-outlook-add-in-for-event-based-activation-preview"></a><span data-ttu-id="c0512-103">配置Outlook加载项进行基于事件的激活和 (预览) </span><span class="sxs-lookup"><span data-stu-id="c0512-103">Configure your Outlook add-in for event-based activation (preview)</span></span>

<span data-ttu-id="c0512-104">如果没有基于事件的激活功能，用户必须显式启动外接程序才能完成其任务。</span><span class="sxs-lookup"><span data-stu-id="c0512-104">Without the event-based activation feature, a user has to explicitly launch an add-in to complete their tasks.</span></span> <span data-ttu-id="c0512-105">此功能使加载项能够运行基于特定事件的任务，尤其是适用于每个项目的操作。</span><span class="sxs-lookup"><span data-stu-id="c0512-105">This feature enables your add-in to run tasks based on certain events, particularly for operations that apply to every item.</span></span> <span data-ttu-id="c0512-106">还可以与任务窗格和无 UI 功能集成。</span><span class="sxs-lookup"><span data-stu-id="c0512-106">You can also integrate with the task pane and UI-less functionality.</span></span>

<span data-ttu-id="c0512-107">在此演练结束时，您将具有一个加载项，只要创建一个新建项目并设置主题，就会运行该加载项。</span><span class="sxs-lookup"><span data-stu-id="c0512-107">By the end of this walkthrough, you'll have an add-in that runs whenever a new item is created and sets the subject.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="c0512-108">此功能仅支持在[Web](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md)上的 Outlook 和具有 Microsoft 365 订阅的 Windows 预览。</span><span class="sxs-lookup"><span data-stu-id="c0512-108">This feature is only supported for [preview](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) in Outlook on the web and on Windows with a Microsoft 365 subscription.</span></span> <span data-ttu-id="c0512-109">有关详细信息，请参阅本文中的如何预览基于 [事件的](#how-to-preview-the-event-based-activation-feature) 激活功能。</span><span class="sxs-lookup"><span data-stu-id="c0512-109">For more details, see [How to preview the event-based activation feature](#how-to-preview-the-event-based-activation-feature) in this article.</span></span>
>
> <span data-ttu-id="c0512-110">由于预览功能可能会随时更改，恕不另行通知，因此不应在生产外接程序中使用。</span><span class="sxs-lookup"><span data-stu-id="c0512-110">Because preview features are subject to change without notice, they shouldn't be used in production add-ins.</span></span>

## <a name="supported-events"></a><span data-ttu-id="c0512-111">支持的事件</span><span class="sxs-lookup"><span data-stu-id="c0512-111">Supported events</span></span>

<span data-ttu-id="c0512-112">目前，支持以下事件。</span><span class="sxs-lookup"><span data-stu-id="c0512-112">At present, the following events are supported.</span></span>

|<span data-ttu-id="c0512-113">事件</span><span class="sxs-lookup"><span data-stu-id="c0512-113">Event</span></span>|<span data-ttu-id="c0512-114">说明</span><span class="sxs-lookup"><span data-stu-id="c0512-114">Description</span></span>|<span data-ttu-id="c0512-115">客户端</span><span class="sxs-lookup"><span data-stu-id="c0512-115">Clients</span></span>|
|---|---|---|
|`OnNewMessageCompose`|<span data-ttu-id="c0512-116">撰写新邮件时 (包括答复、全部答复和转发) ，但不包括编辑时，例如草稿。</span><span class="sxs-lookup"><span data-stu-id="c0512-116">On composing a new message (includes reply, reply all, and forward) but not on editing, for example, a draft.</span></span>|<span data-ttu-id="c0512-117">Windows、Web</span><span class="sxs-lookup"><span data-stu-id="c0512-117">Windows, web</span></span>|
|`OnNewAppointmentOrganizer`|<span data-ttu-id="c0512-118">创建新约会但不编辑现有约会时。</span><span class="sxs-lookup"><span data-stu-id="c0512-118">On creating a new appointment but not on editing an existing one.</span></span>|<span data-ttu-id="c0512-119">Windows、Web</span><span class="sxs-lookup"><span data-stu-id="c0512-119">Windows, web</span></span>|
|`OnMessageAttachmentsChanged`|<span data-ttu-id="c0512-120">在撰写邮件时添加或删除附件。</span><span class="sxs-lookup"><span data-stu-id="c0512-120">On adding or removing attachments while composing a message.</span></span>|<span data-ttu-id="c0512-121">Windows</span><span class="sxs-lookup"><span data-stu-id="c0512-121">Windows</span></span>|
|`OnAppointmentAttachmentsChanged`|<span data-ttu-id="c0512-122">在撰写约会时添加或删除附件。</span><span class="sxs-lookup"><span data-stu-id="c0512-122">On adding or removing attachments while composing an appointment.</span></span>|<span data-ttu-id="c0512-123">Windows</span><span class="sxs-lookup"><span data-stu-id="c0512-123">Windows</span></span>|
|`OnMessageRecipientsChanged`|<span data-ttu-id="c0512-124">在撰写邮件时添加或删除收件人。</span><span class="sxs-lookup"><span data-stu-id="c0512-124">On adding or removing recipients while composing a message.</span></span>|<span data-ttu-id="c0512-125">Windows</span><span class="sxs-lookup"><span data-stu-id="c0512-125">Windows</span></span>|
|`OnAppointmentAttendeesChanged`|<span data-ttu-id="c0512-126">在撰写约会时添加或删除与会者。</span><span class="sxs-lookup"><span data-stu-id="c0512-126">On adding or removing attendees while composing an appointment.</span></span>|<span data-ttu-id="c0512-127">Windows</span><span class="sxs-lookup"><span data-stu-id="c0512-127">Windows</span></span>|
|`OnAppointmentTimeChanged`|<span data-ttu-id="c0512-128">在撰写约会时更改日期/时间。</span><span class="sxs-lookup"><span data-stu-id="c0512-128">On changing date/time while composing an appointment.</span></span>|<span data-ttu-id="c0512-129">Windows</span><span class="sxs-lookup"><span data-stu-id="c0512-129">Windows</span></span>|
|`OnAppointmentRecurrenceChanged`|<span data-ttu-id="c0512-130">在撰写约会时添加、更改或删除定期详细信息。</span><span class="sxs-lookup"><span data-stu-id="c0512-130">On adding, changing, or removing the recurrence details while composing an appointment.</span></span> <span data-ttu-id="c0512-131">如果日期/时间发生更改， `OnAppointmentTimeChanged` 也会触发该事件。</span><span class="sxs-lookup"><span data-stu-id="c0512-131">If the date/time is changed, the `OnAppointmentTimeChanged` event will also be fired.</span></span>|<span data-ttu-id="c0512-132">Windows</span><span class="sxs-lookup"><span data-stu-id="c0512-132">Windows</span></span>|
|`OnInfoBarDismissClicked`|<span data-ttu-id="c0512-133">在撰写邮件或约会项目时关闭通知。</span><span class="sxs-lookup"><span data-stu-id="c0512-133">On dismissing a notification while composing a message or appointment item.</span></span> <span data-ttu-id="c0512-134">仅通知添加了通知的外接程序。</span><span class="sxs-lookup"><span data-stu-id="c0512-134">Only the add-in that added the notification will be notified.</span></span>|<span data-ttu-id="c0512-135">Windows</span><span class="sxs-lookup"><span data-stu-id="c0512-135">Windows</span></span>|

## <a name="how-to-preview-the-event-based-activation-feature"></a><span data-ttu-id="c0512-136">如何预览基于事件的激活功能</span><span class="sxs-lookup"><span data-stu-id="c0512-136">How to preview the event-based activation feature</span></span>

<span data-ttu-id="c0512-137">我们邀请你试用基于事件的激活功能！</span><span class="sxs-lookup"><span data-stu-id="c0512-137">We invite you to try out the event-based activation feature!</span></span> <span data-ttu-id="c0512-138">请告诉我们你的方案，以及我们如何通过反馈提供反馈GitHub (请参阅此页面末尾的反馈部分) 。 </span><span class="sxs-lookup"><span data-stu-id="c0512-138">Let us know your scenarios and how we can improve by giving us feedback through GitHub (see the **Feedback** section at the end of this page).</span></span>

<span data-ttu-id="c0512-139">预览此功能：</span><span class="sxs-lookup"><span data-stu-id="c0512-139">To preview this feature:</span></span>

- <span data-ttu-id="c0512-140">For Outlook on the web：</span><span class="sxs-lookup"><span data-stu-id="c0512-140">For Outlook on the web:</span></span>
  - <span data-ttu-id="c0512-141">[在租户 上配置Microsoft 365版本](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)。</span><span class="sxs-lookup"><span data-stu-id="c0512-141">[Configure targeted release on your Microsoft 365 tenant](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center).</span></span>
  - <span data-ttu-id="c0512-142">在 上 **引用** beta https://appsforoffice.microsoft.com/lib/beta/hosted/office.js) CDN (。</span><span class="sxs-lookup"><span data-stu-id="c0512-142">Reference the **beta** library on the CDN (https://appsforoffice.microsoft.com/lib/beta/hosted/office.js).</span></span> <span data-ttu-id="c0512-143">TypeScript[编译和](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts)定义的类型IntelliSense位于 CDN[和 DefinitelyTyped 中](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts)。</span><span class="sxs-lookup"><span data-stu-id="c0512-143">The [type definition file](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts) for TypeScript compilation and IntelliSense is found at the CDN and [DefinitelyTyped](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts).</span></span> <span data-ttu-id="c0512-144">可以使用 安装这些类型 `npm install --save-dev @types/office-js-preview` 。</span><span class="sxs-lookup"><span data-stu-id="c0512-144">You can install these types with `npm install --save-dev @types/office-js-preview`.</span></span>
- <span data-ttu-id="c0512-145">有关Outlook Windows：</span><span class="sxs-lookup"><span data-stu-id="c0512-145">For Outlook on Windows:</span></span>
  - <span data-ttu-id="c0512-146">最低要求版本为 16.0.14026.20000。</span><span class="sxs-lookup"><span data-stu-id="c0512-146">The minimum required build is 16.0.14026.20000.</span></span> <span data-ttu-id="c0512-147">加入[Office 预览体验计划](https://insider.office.com)，以访问 Office beta 版本。</span><span class="sxs-lookup"><span data-stu-id="c0512-147">Join the [Office Insider program](https://insider.office.com) for access to Office beta builds.</span></span>
  - <span data-ttu-id="c0512-148">配置注册表：</span><span class="sxs-lookup"><span data-stu-id="c0512-148">Configure the registry:</span></span>
    1. <span data-ttu-id="c0512-149">创建注册表项 `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Outlook\Options\WebExt\Developer` 。</span><span class="sxs-lookup"><span data-stu-id="c0512-149">Create the registry key `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Outlook\Options\WebExt\Developer`.</span></span>
    1. <span data-ttu-id="c0512-150">添加一个名为 的 `EnableBetaAPIsInJavaScript` 条目，将值设置为 `1` 。</span><span class="sxs-lookup"><span data-stu-id="c0512-150">Add an entry named `EnableBetaAPIsInJavaScript` and set the value to `1`.</span></span>

## <a name="set-up-your-environment"></a><span data-ttu-id="c0512-151">设置环境</span><span class="sxs-lookup"><span data-stu-id="c0512-151">Set up your environment</span></span>

<span data-ttu-id="c0512-152">完成[Outlook](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator)使用适用于加载项的 Yeoman 生成器创建加载项Office快速入门。</span><span class="sxs-lookup"><span data-stu-id="c0512-152">Complete the [Outlook quick start](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) which creates an add-in project with the Yeoman generator for Office Add-ins.</span></span>

## <a name="configure-the-manifest"></a><span data-ttu-id="c0512-153">配置清单</span><span class="sxs-lookup"><span data-stu-id="c0512-153">Configure the manifest</span></span>

<span data-ttu-id="c0512-154">若要启用加载项的基于事件的激活，必须在清单节点中配置 [Runtimes](../reference/manifest/runtimes.md) 元素和 [LaunchEvent](../reference/manifest/extensionpoint.md#launchevent-preview) `VersionOverridesV1_1` 扩展点。</span><span class="sxs-lookup"><span data-stu-id="c0512-154">To enable event-based activation of your add-in, you must configure the [Runtimes](../reference/manifest/runtimes.md) element and [LaunchEvent](../reference/manifest/extensionpoint.md#launchevent-preview) extension point in the `VersionOverridesV1_1` node of the manifest.</span></span> <span data-ttu-id="c0512-155">目前， `DesktopFormFactor` 是唯一受支持的外形类型。</span><span class="sxs-lookup"><span data-stu-id="c0512-155">For now, `DesktopFormFactor` is the only supported form factor.</span></span>

1. <span data-ttu-id="c0512-156">在代码编辑器中，打开快速启动项目。</span><span class="sxs-lookup"><span data-stu-id="c0512-156">In your code editor, open the quick start project.</span></span>

1. <span data-ttu-id="c0512-157">打开 **manifest.xml** 根目录下的文件。</span><span class="sxs-lookup"><span data-stu-id="c0512-157">Open the **manifest.xml** file located at the root of your project.</span></span>

1. <span data-ttu-id="c0512-158">选择整个节点 (包括打开和关闭) `<VersionOverrides>` 并将其替换为以下 XML，然后保存更改。</span><span class="sxs-lookup"><span data-stu-id="c0512-158">Select the entire `<VersionOverrides>` node (including open and close tags) and replace it with the following XML, then save your changes.</span></span>

```XML
<VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
    <Requirements>
      <bt:Sets DefaultMinVersion="1.3">
        <bt:Set Name="Mailbox" />
      </bt:Sets>
    </Requirements>
    <Hosts>
      <Host xsi:type="MailHost">
        <!-- Event-based activation happens in a lightweight runtime.-->
        <Runtimes>
          <!-- HTML file including reference to or inline JavaScript event handlers.
               This is used by Outlook on the web. -->
          <Runtime resid="WebViewRuntime.Url">
            <!-- JavaScript file containing event handlers. This is used by Outlook Desktop. -->
            <Override type="javascript" resid="JSRuntime.Url"/>
          </Runtime>
        </Runtimes>
        <DesktopFormFactor>
          <FunctionFile resid="Commands.Url" />
          <ExtensionPoint xsi:type="MessageReadCommandSurface">
            <OfficeTab id="TabDefault">
              <Group id="msgReadGroup">
                <Label resid="GroupLabel" />
                <Control xsi:type="Button" id="msgReadOpenPaneButton">
                  <Label resid="TaskpaneButton.Label" />
                  <Supertip>
                    <Title resid="TaskpaneButton.Label" />
                    <Description resid="TaskpaneButton.Tooltip" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Icon.16x16" />
                    <bt:Image size="32" resid="Icon.32x32" />
                    <bt:Image size="80" resid="Icon.80x80" />
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <SourceLocation resid="Taskpane.Url" />
                  </Action>
                </Control>
                <Control xsi:type="Button" id="ActionButton">
                  <Label resid="ActionButton.Label"/>
                  <Supertip>
                    <Title resid="ActionButton.Label"/>
                    <Description resid="ActionButton.Tooltip"/>
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Icon.16x16"/>
                    <bt:Image size="32" resid="Icon.32x32"/>
                    <bt:Image size="80" resid="Icon.80x80"/>
                  </Icon>
                  <Action xsi:type="ExecuteFunction">
                    <FunctionName>action</FunctionName>
                  </Action>
                </Control>
              </Group>
            </OfficeTab>
          </ExtensionPoint>

          <!-- Can configure other command surface extension points for add-in command support. -->

          <!-- Enable launching the add-in on the included events. -->
          <ExtensionPoint xsi:type="LaunchEvent">
            <LaunchEvents>
              <LaunchEvent Type="OnNewMessageCompose" FunctionName="onMessageComposeHandler"/>
              <LaunchEvent Type="OnNewAppointmentOrganizer" FunctionName="onAppointmentComposeHandler"/>
            </LaunchEvents>
            <!-- Identifies the runtime to be used (also referenced by the Runtime element). -->
            <SourceLocation resid="WebViewRuntime.Url"/>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>
    </Hosts>
    <Resources>
      <bt:Images>
        <bt:Image id="Icon.16x16" DefaultValue="https://localhost:3000/assets/icon-16.png"/>
        <bt:Image id="Icon.32x32" DefaultValue="https://localhost:3000/assets/icon-32.png"/>
        <bt:Image id="Icon.80x80" DefaultValue="https://localhost:3000/assets/icon-80.png"/>
      </bt:Images>
      <bt:Urls>
        <bt:Url id="Commands.Url" DefaultValue="https://localhost:3000/commands.html" />
        <bt:Url id="Taskpane.Url" DefaultValue="https://localhost:3000/taskpane.html" />
        <bt:Url id="WebViewRuntime.Url" DefaultValue="https://localhost:3000/commands.html" />
        <!-- Entry needed for Outlook Desktop. -->
        <bt:Url id="JSRuntime.Url" DefaultValue="https://localhost:3000/src/commands/commands.js" />
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="GroupLabel" DefaultValue="Contoso Add-in"/>
        <bt:String id="TaskpaneButton.Label" DefaultValue="Show Taskpane"/>
        <bt:String id="ActionButton.Label" DefaultValue="Perform an action"/>
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="TaskpaneButton.Tooltip" DefaultValue="Opens a pane displaying all available properties."/>
        <bt:String id="ActionButton.Tooltip" DefaultValue="Perform an action when clicked."/>
      </bt:LongStrings>
    </Resources>
  </VersionOverrides>
</VersionOverrides>
```

<span data-ttu-id="c0512-159">Outlook Windows使用 JavaScript 文件，Outlook Web 上的开发人员使用可以引用同一 JavaScript 文件的 HTML 文件。</span><span class="sxs-lookup"><span data-stu-id="c0512-159">Outlook on Windows uses a JavaScript file, while Outlook on the web uses an HTML file that can reference the same JavaScript file.</span></span> <span data-ttu-id="c0512-160">你必须在清单的节点中提供对这两个文件的引用，因为 Outlook 平台最终确定是使用 HTML 还是基于 Outlook `Resources` 客户端的 JavaScript。</span><span class="sxs-lookup"><span data-stu-id="c0512-160">You must provide references to both these files in the `Resources` node of the manifest as the Outlook platform ultimately determines whether to use HTML or JavaScript based on the Outlook client.</span></span> <span data-ttu-id="c0512-161">因此，若要配置事件处理，请提供 HTML 在 元素中的位置，然后在其子元素中提供 JavaScript 文件内附或 HTML `Runtime` `Override` 引用的位置。</span><span class="sxs-lookup"><span data-stu-id="c0512-161">As such, to configure event handling, provide the location of the HTML in the `Runtime` element, then in its `Override` child element provide the location of the JavaScript file inlined or referenced by the HTML.</span></span>

> [!TIP]
> <span data-ttu-id="c0512-162">若要了解有关加载项清单Outlook，请参阅Outlook[加载项清单](manifests.md)。</span><span class="sxs-lookup"><span data-stu-id="c0512-162">To learn more about manifests for Outlook add-ins, see [Outlook add-in manifests](manifests.md).</span></span>

## <a name="implement-event-handling"></a><span data-ttu-id="c0512-163">实现事件处理</span><span class="sxs-lookup"><span data-stu-id="c0512-163">Implement event handling</span></span>

<span data-ttu-id="c0512-164">您必须对所选事件实现处理。</span><span class="sxs-lookup"><span data-stu-id="c0512-164">You have to implement handling for your selected events.</span></span>

<span data-ttu-id="c0512-165">在此方案中，您将添加用于撰写新项的处理。</span><span class="sxs-lookup"><span data-stu-id="c0512-165">In this scenario, you'll add handling for composing new items.</span></span>

1. <span data-ttu-id="c0512-166">从同一快速启动项目中，在代码编辑器中打开 **commands.js./src/commands/commands.js** 文件。</span><span class="sxs-lookup"><span data-stu-id="c0512-166">From the same quick start project, open the file **./src/commands/commands.js** in your code editor.</span></span>

1. <span data-ttu-id="c0512-167">在 函数 `action` 之后，插入以下 JavaScript 函数。</span><span class="sxs-lookup"><span data-stu-id="c0512-167">After the `action` function, insert the following JavaScript functions.</span></span>

    ```js
    function onMessageComposeHandler(event) {
      setSubject(event);
    }
    function onAppointmentComposeHandler(event) {
      setSubject(event);
    }
    function setSubject(event) {
      Office.context.mailbox.item.subject.setAsync(
        "Set by an event-based add-in!",
        {
          "asyncContext" : event
        },
        function (asyncResult) {
          // Handle success or error.
          if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
            console.error("Failed to set subject: " + JSON.stringify(asyncResult.error));
          }
    
          // Call event.completed() after all work is done.
          asyncResult.asyncContext.completed();
        });
    }
    ```

1. <span data-ttu-id="c0512-168">若要使用 Yeo  Office man Outlook加载项的 Yeoman 生成器生成的此项目在 Web 上运行的函数，在文件末尾添加以下语句。</span><span class="sxs-lookup"><span data-stu-id="c0512-168">For the functions to work in **Outlook on the web** with this project generated by the Yeoman generator for Office Add-ins, add the following statements at the end of the file.</span></span>

    ```js
    g.onMessageComposeHandler = onMessageComposeHandler;
    g.onAppointmentComposeHandler = onAppointmentComposeHandler;
    ```

1. <span data-ttu-id="c0512-169">若要使函数在 Outlook **中** Windows，在文件末尾添加以下 JavaScript 代码。</span><span class="sxs-lookup"><span data-stu-id="c0512-169">For the functions to work in **Outlook on Windows**, add the following JavaScript code at the end of the file.</span></span>

    ```js
    if (Office.actions) {
      // 1st parameter: FunctionName of LaunchEvent in the manifest; 2nd parameter: Its implementation in this .js file.
      Office.actions.associate("onMessageComposeHandler", onMessageComposeHandler);
      Office.actions.associate("onAppointmentComposeHandler", onAppointmentComposeHandler);
    }
    ```

    <span data-ttu-id="c0512-170">**注意**：检查 `Office.actions` 以确保Outlook忽略这些语句。</span><span class="sxs-lookup"><span data-stu-id="c0512-170">**Note**: Checking for `Office.actions` ensures that Outlook on the web ignores these statements.</span></span>

1. <span data-ttu-id="c0512-171">保存所做的更改。</span><span class="sxs-lookup"><span data-stu-id="c0512-171">Save your changes.</span></span>

## <a name="try-it-out"></a><span data-ttu-id="c0512-172">试用</span><span class="sxs-lookup"><span data-stu-id="c0512-172">Try it out</span></span>

1. <span data-ttu-id="c0512-173">在项目的根目录中运行以下命令。</span><span class="sxs-lookup"><span data-stu-id="c0512-173">Run the following command in the root directory of your project.</span></span> <span data-ttu-id="c0512-174">如果运行此命令，本地 Web 服务器将启动（如果尚未运行），并将旁加载加载项。</span><span class="sxs-lookup"><span data-stu-id="c0512-174">When you run this command, the local web server will start (if it's not already running) and your add-in will be sideloaded.</span></span>

    ```command&nbsp;line
    npm start
    ```

1. <span data-ttu-id="c0512-175">在 Outlook 网页版中，创建新邮件。</span><span class="sxs-lookup"><span data-stu-id="c0512-175">In Outlook on the web, create a new message.</span></span>

    ![在撰写时设置主题Outlook网页中的邮件窗口屏幕截图](../images/outlook-web-autolaunch-1.png)

1. <span data-ttu-id="c0512-177">在Outlook中Windows新建一封邮件。</span><span class="sxs-lookup"><span data-stu-id="c0512-177">In Outlook on Windows, create a new message.</span></span>

    ![撰写时主题集Outlook Windows中邮件窗口的屏幕截图](../images/outlook-win-autolaunch.png)

    > [!NOTE]
    > <span data-ttu-id="c0512-179">如果看到错误"无法从 localhost 打开此外接程序"，则需要启用环回豁免。</span><span class="sxs-lookup"><span data-stu-id="c0512-179">If you see the error "We can't open this add-in from localhost," you'll need to enable a loopback exemption.</span></span>
    >
    > 1. <span data-ttu-id="c0512-180">关闭 Outlook。</span><span class="sxs-lookup"><span data-stu-id="c0512-180">Close Outlook.</span></span>
    > 2. <span data-ttu-id="c0512-181">打开 **任务管理器** ， **并确保msoadfs.exe进程** 未运行。</span><span class="sxs-lookup"><span data-stu-id="c0512-181">Open the **Task Manager** and ensure that the **msoadfs.exe** process is not running.</span></span>
    > 3. <span data-ttu-id="c0512-182">运行以下命令。</span><span class="sxs-lookup"><span data-stu-id="c0512-182">Run the following command.</span></span>
    >
    >     ```command&nbsp;line
    >     call %SystemRoot%\System32\CheckNetIsolation.exe LoopbackExempt -a -n=1_http___localhost_300004ACA5EC-D79A-43EA-AB47-E50E47DD96FC
    >     ```
    >
    > 4. <span data-ttu-id="c0512-183">重新启动 Outlook。</span><span class="sxs-lookup"><span data-stu-id="c0512-183">Restart Outlook.</span></span>

## <a name="debug"></a><span data-ttu-id="c0512-184">Debug</span><span class="sxs-lookup"><span data-stu-id="c0512-184">Debug</span></span>

<span data-ttu-id="c0512-185">当你实现自己的功能时，你可能需要调试代码。</span><span class="sxs-lookup"><span data-stu-id="c0512-185">As you implement your own functionality, you may need to debug your code.</span></span> <span data-ttu-id="c0512-186">有关如何调试基于事件的外接程序激活的指南，请参阅调试基于事件Outlook[外接程序](debug-autolaunch.md)。</span><span class="sxs-lookup"><span data-stu-id="c0512-186">For guidance on how to debug event-based add-in activation, see [Debug your event-based Outlook add-in](debug-autolaunch.md).</span></span>

## <a name="event-based-activation-behavior-and-limitations"></a><span data-ttu-id="c0512-187">基于事件的激活行为和限制</span><span class="sxs-lookup"><span data-stu-id="c0512-187">Event-based activation behavior and limitations</span></span>

<span data-ttu-id="c0512-188">基于事件激活的加载项应尽量短运行、轻量且非轻量。</span><span class="sxs-lookup"><span data-stu-id="c0512-188">Add-ins that activate based on events are expected to be short-running, lightweight, and as non-invasive as possible.</span></span> <span data-ttu-id="c0512-189">若要表示加载项已完成处理启动事件，建议让加载项调用 `event.completed` 方法。</span><span class="sxs-lookup"><span data-stu-id="c0512-189">To signal that your add-in has completed processing the launch event, we recommend you have your add-in call the `event.completed` method.</span></span> <span data-ttu-id="c0512-190">如果未进行该调用，外接程序将在大约 300 秒（运行基于事件的外接程序所允许的最大时间长度）内退出。当用户关闭撰写窗口时，外接程序也将结束。</span><span class="sxs-lookup"><span data-stu-id="c0512-190">If that call is not made, the add-in will time out within approximately 300 seconds, the maximum length of time allowed for running event-based add-ins. The add-in also ends when the user closes the compose window.</span></span>

<span data-ttu-id="c0512-191">如果用户有多个订阅了同一事件的加载项，Outlook平台将按特定顺序启动加载项。</span><span class="sxs-lookup"><span data-stu-id="c0512-191">If the user has multiple add-ins that subscribed to the same event, the Outlook platform launches the add-ins in no particular order.</span></span> <span data-ttu-id="c0512-192">目前，只能主动运行五个基于事件的加载项。</span><span class="sxs-lookup"><span data-stu-id="c0512-192">Currently, only five event-based add-ins can be actively running.</span></span> <span data-ttu-id="c0512-193">任何其他外接程序将推送到队列，然后随着之前处于活动状态的外接程序完成或停用而运行。</span><span class="sxs-lookup"><span data-stu-id="c0512-193">Any additional add-ins are pushed to a queue then run as previously active add-ins are completed or deactivated.</span></span>

<span data-ttu-id="c0512-194">用户可以切换或导航离开外接程序开始运行的当前邮件项目。</span><span class="sxs-lookup"><span data-stu-id="c0512-194">The user can switch or navigate away from the current mail item where the add-in started running.</span></span> <span data-ttu-id="c0512-195">启动的加载项将在后台完成其操作。</span><span class="sxs-lookup"><span data-stu-id="c0512-195">The add-in that was launched will finish its operation in the background.</span></span>

<span data-ttu-id="c0512-196">某些Office.js更改或更改 UI 的 API 不允许来自基于事件的外接程序。以下是阻止的 API：</span><span class="sxs-lookup"><span data-stu-id="c0512-196">Some Office.js APIs that change or alter the UI are not allowed from event-based add-ins. The following are the blocked APIs:</span></span>

- <span data-ttu-id="c0512-197">在 `Office.context.auth` 下：</span><span class="sxs-lookup"><span data-stu-id="c0512-197">Under `Office.context.auth`:</span></span>
  - `getAccessToken`
  - `getAccessTokenAsync`
- <span data-ttu-id="c0512-198">在 `Office.context.mailbox` 下：</span><span class="sxs-lookup"><span data-stu-id="c0512-198">Under `Office.context.mailbox`:</span></span>
  - `displayAppointmentForm`
  - `displayMessageForm`
  - `displayNewAppointmentForm`
  - `displayNewMessageForm`
- <span data-ttu-id="c0512-199">在 `Office.context.mailbox.item` 下：</span><span class="sxs-lookup"><span data-stu-id="c0512-199">Under `Office.context.mailbox.item`:</span></span>
  - `close`
- <span data-ttu-id="c0512-200">在 `Office.context.ui` 下：</span><span class="sxs-lookup"><span data-stu-id="c0512-200">Under `Office.context.ui`:</span></span>
  - `displayDialogAsync`
  - `messageParent`

## <a name="see-also"></a><span data-ttu-id="c0512-201">另请参阅</span><span class="sxs-lookup"><span data-stu-id="c0512-201">See also</span></span>

- [<span data-ttu-id="c0512-202">Outlook 加载项清单</span><span class="sxs-lookup"><span data-stu-id="c0512-202">Outlook add-in manifests</span></span>](manifests.md)
- [<span data-ttu-id="c0512-203">如何调试基于事件的外接程序</span><span class="sxs-lookup"><span data-stu-id="c0512-203">How to debug event-based add-ins</span></span>](debug-autolaunch.md)
