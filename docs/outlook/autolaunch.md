---
title: 配置Outlook加载项进行基于事件的激活
description: 了解如何配置Outlook加载项进行基于事件的激活。
ms.topic: article
ms.date: 05/26/2021
localization_priority: Normal
ms.openlocfilehash: debf6db16adc8e0bc923142da1e85629b8a1daa8
ms.sourcegitcommit: a42ae8b804f944061c87bbd9d9f67990e4cf5e36
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/28/2021
ms.locfileid: "52697195"
---
# <a name="configure-your-outlook-add-in-for-event-based-activation"></a><span data-ttu-id="70c4d-103">配置Outlook加载项进行基于事件的激活</span><span class="sxs-lookup"><span data-stu-id="70c4d-103">Configure your Outlook add-in for event-based activation</span></span>

<span data-ttu-id="70c4d-104">如果没有基于事件的激活功能，用户必须显式启动外接程序才能完成其任务。</span><span class="sxs-lookup"><span data-stu-id="70c4d-104">Without the event-based activation feature, a user has to explicitly launch an add-in to complete their tasks.</span></span> <span data-ttu-id="70c4d-105">此功能使加载项能够运行基于特定事件的任务，尤其是适用于每个项目的操作。</span><span class="sxs-lookup"><span data-stu-id="70c4d-105">This feature enables your add-in to run tasks based on certain events, particularly for operations that apply to every item.</span></span> <span data-ttu-id="70c4d-106">还可以与任务窗格和无 UI 功能集成。</span><span class="sxs-lookup"><span data-stu-id="70c4d-106">You can also integrate with the task pane and UI-less functionality.</span></span>

<span data-ttu-id="70c4d-107">在此演练结束时，您将具有一个加载项，只要创建一个新建项目并设置主题，就会运行该加载项。</span><span class="sxs-lookup"><span data-stu-id="70c4d-107">By the end of this walkthrough, you'll have an add-in that runs whenever a new item is created and sets the subject.</span></span>

> [!NOTE]
> <span data-ttu-id="70c4d-108">要求集 [1.10](../reference/objectmodel/requirement-set-1.10/outlook-requirement-set-1.10.md)中引入了对此功能的支持。</span><span class="sxs-lookup"><span data-stu-id="70c4d-108">Support for this feature was introduced in [requirement set 1.10](../reference/objectmodel/requirement-set-1.10/outlook-requirement-set-1.10.md).</span></span> <span data-ttu-id="70c4d-109">请查看支持此要求集的[客户端和平台](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients)。</span><span class="sxs-lookup"><span data-stu-id="70c4d-109">See [clients and platforms](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) that support this requirement set.</span></span>

## <a name="supported-events"></a><span data-ttu-id="70c4d-110">支持的事件</span><span class="sxs-lookup"><span data-stu-id="70c4d-110">Supported events</span></span>

<span data-ttu-id="70c4d-111">目前，Web 和 Web 上支持以下Windows。</span><span class="sxs-lookup"><span data-stu-id="70c4d-111">At present, the following events are supported on the web and on Windows.</span></span>

|<span data-ttu-id="70c4d-112">事件</span><span class="sxs-lookup"><span data-stu-id="70c4d-112">Event</span></span>|<span data-ttu-id="70c4d-113">说明</span><span class="sxs-lookup"><span data-stu-id="70c4d-113">Description</span></span>|
|---|---|
|`OnNewMessageCompose`|<span data-ttu-id="70c4d-114">撰写新邮件时 (包括答复、全部答复和转发) ，但不包括编辑时，例如草稿。</span><span class="sxs-lookup"><span data-stu-id="70c4d-114">On composing a new message (includes reply, reply all, and forward) but not on editing, for example, a draft.</span></span>|
|`OnNewAppointmentOrganizer`|<span data-ttu-id="70c4d-115">创建新约会但不编辑现有约会时。</span><span class="sxs-lookup"><span data-stu-id="70c4d-115">On creating a new appointment but not on editing an existing one.</span></span>|
|`OnMessageAttachmentsChanged`\*|<span data-ttu-id="70c4d-116">在撰写邮件时添加或删除附件。</span><span class="sxs-lookup"><span data-stu-id="70c4d-116">On adding or removing attachments while composing a message.</span></span>|
|`OnAppointmentAttachmentsChanged`\*|<span data-ttu-id="70c4d-117">在撰写约会时添加或删除附件。</span><span class="sxs-lookup"><span data-stu-id="70c4d-117">On adding or removing attachments while composing an appointment.</span></span>|
|`OnMessageRecipientsChanged`\*|<span data-ttu-id="70c4d-118">在撰写邮件时添加或删除收件人。</span><span class="sxs-lookup"><span data-stu-id="70c4d-118">On adding or removing recipients while composing a message.</span></span>|
|`OnAppointmentAttendeesChanged`\*|<span data-ttu-id="70c4d-119">在撰写约会时添加或删除与会者。</span><span class="sxs-lookup"><span data-stu-id="70c4d-119">On adding or removing attendees while composing an appointment.</span></span>|
|`OnAppointmentTimeChanged`\*|<span data-ttu-id="70c4d-120">在撰写约会时更改日期/时间。</span><span class="sxs-lookup"><span data-stu-id="70c4d-120">On changing date/time while composing an appointment.</span></span>|
|`OnAppointmentRecurrenceChanged`\*|<span data-ttu-id="70c4d-121">在撰写约会时添加、更改或删除定期详细信息。</span><span class="sxs-lookup"><span data-stu-id="70c4d-121">On adding, changing, or removing the recurrence details while composing an appointment.</span></span> <span data-ttu-id="70c4d-122">如果日期/时间发生更改， `OnAppointmentTimeChanged` 也会触发该事件。</span><span class="sxs-lookup"><span data-stu-id="70c4d-122">If the date/time is changed, the `OnAppointmentTimeChanged` event will also be fired.</span></span>|
|`OnInfoBarDismissClicked`\*|<span data-ttu-id="70c4d-123">在撰写邮件或约会项目时关闭通知。</span><span class="sxs-lookup"><span data-stu-id="70c4d-123">On dismissing a notification while composing a message or appointment item.</span></span> <span data-ttu-id="70c4d-124">仅通知添加了通知的外接程序。</span><span class="sxs-lookup"><span data-stu-id="70c4d-124">Only the add-in that added the notification will be notified.</span></span>|

> [!IMPORTANT]
> <span data-ttu-id="70c4d-125">\*此事件仅在使用 Outlook[](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md)订阅的 Outlook 和 Windows 中Microsoft 365受支持。</span><span class="sxs-lookup"><span data-stu-id="70c4d-125">\* This event is only supported for [preview](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) in Outlook on the web and on Windows with a Microsoft 365 subscription.</span></span> <span data-ttu-id="70c4d-126">有关详细信息，请参阅 [本文中的](#how-to-preview) 如何预览。</span><span class="sxs-lookup"><span data-stu-id="70c4d-126">For more details, see [How to preview](#how-to-preview) in this article.</span></span>
>
> <span data-ttu-id="70c4d-127">由于预览功能可能会随时更改，恕不另行通知，因此不应在生产外接程序中使用。</span><span class="sxs-lookup"><span data-stu-id="70c4d-127">Because preview features are subject to change without notice, they shouldn't be used in production add-ins.</span></span>

## <a name="how-to-preview"></a><span data-ttu-id="70c4d-128">如何预览</span><span class="sxs-lookup"><span data-stu-id="70c4d-128">How to preview</span></span>

<span data-ttu-id="70c4d-129">我们邀请你试用新事件！</span><span class="sxs-lookup"><span data-stu-id="70c4d-129">We invite you to try out the new events!</span></span> <span data-ttu-id="70c4d-130">请告诉我们你的方案，以及我们如何通过反馈提供反馈GitHub (请参阅此页面末尾的反馈部分) 。 </span><span class="sxs-lookup"><span data-stu-id="70c4d-130">Let us know your scenarios and how we can improve by giving us feedback through GitHub (see the **Feedback** section at the end of this page).</span></span>

<span data-ttu-id="70c4d-131">预览此功能：</span><span class="sxs-lookup"><span data-stu-id="70c4d-131">To preview this feature:</span></span>

- <span data-ttu-id="70c4d-132">For Outlook on the web：</span><span class="sxs-lookup"><span data-stu-id="70c4d-132">For Outlook on the web:</span></span>
  - <span data-ttu-id="70c4d-133">[在租户 上配置Microsoft 365版本](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)。</span><span class="sxs-lookup"><span data-stu-id="70c4d-133">[Configure targeted release on your Microsoft 365 tenant](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center).</span></span>
  - <span data-ttu-id="70c4d-134">在 上 **引用** beta https://appsforoffice.microsoft.com/lib/beta/hosted/office.js) CDN (。</span><span class="sxs-lookup"><span data-stu-id="70c4d-134">Reference the **beta** library on the CDN (https://appsforoffice.microsoft.com/lib/beta/hosted/office.js).</span></span> <span data-ttu-id="70c4d-135">TypeScript[编译和](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts)定义的类型IntelliSense位于 CDN[和 DefinitelyTyped 中](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts)。</span><span class="sxs-lookup"><span data-stu-id="70c4d-135">The [type definition file](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts) for TypeScript compilation and IntelliSense is found at the CDN and [DefinitelyTyped](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts).</span></span> <span data-ttu-id="70c4d-136">可以使用 安装这些类型 `npm install --save-dev @types/office-js-preview` 。</span><span class="sxs-lookup"><span data-stu-id="70c4d-136">You can install these types with `npm install --save-dev @types/office-js-preview`.</span></span>
- <span data-ttu-id="70c4d-137">有关Outlook Windows：</span><span class="sxs-lookup"><span data-stu-id="70c4d-137">For Outlook on Windows:</span></span>
  - <span data-ttu-id="70c4d-138">最低要求版本为 16.0.14026.20000。</span><span class="sxs-lookup"><span data-stu-id="70c4d-138">The minimum required build is 16.0.14026.20000.</span></span> <span data-ttu-id="70c4d-139">加入[Office 预览体验计划](https://insider.office.com)，以访问 Office beta 版本。</span><span class="sxs-lookup"><span data-stu-id="70c4d-139">Join the [Office Insider program](https://insider.office.com) for access to Office beta builds.</span></span>
  - <span data-ttu-id="70c4d-140">配置注册表。</span><span class="sxs-lookup"><span data-stu-id="70c4d-140">Configure the registry.</span></span> <span data-ttu-id="70c4d-141">Outlook包括 Office.js 的生产和 beta 版本的本地副本，而不是从 CDN。</span><span class="sxs-lookup"><span data-stu-id="70c4d-141">Outlook includes a local copy of the production and beta versions of Office.js instead of loading from the CDN.</span></span> <span data-ttu-id="70c4d-142">默认情况下，将引用 API 的本地生产副本。</span><span class="sxs-lookup"><span data-stu-id="70c4d-142">By default, the local production copy of the API is referenced.</span></span> <span data-ttu-id="70c4d-143">若要切换到 JavaScript API 的本地 beta Outlook，需要添加此注册表项，否则可能无法找到 beta API。</span><span class="sxs-lookup"><span data-stu-id="70c4d-143">To switch to the local beta copy of the Outlook JavaScript APIs, you need to add this registry entry, otherwise beta APIs may not be found.</span></span>
    1. <span data-ttu-id="70c4d-144">创建注册表项 `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Outlook\Options\WebExt\Developer` 。</span><span class="sxs-lookup"><span data-stu-id="70c4d-144">Create the registry key `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Outlook\Options\WebExt\Developer`.</span></span>
    1. <span data-ttu-id="70c4d-145">添加一个名为 的 `EnableBetaAPIsInJavaScript` 条目，将值设置为 `1` 。</span><span class="sxs-lookup"><span data-stu-id="70c4d-145">Add an entry named `EnableBetaAPIsInJavaScript` and set the value to `1`.</span></span> <span data-ttu-id="70c4d-146">下图显示注册表应该呈现的状态。</span><span class="sxs-lookup"><span data-stu-id="70c4d-146">The following image shows what the registry should look like.</span></span>

        ![具有 EnableBetaAPIsInJavaScript 注册表项值的注册表编辑器的屏幕截图](../images/outlook-beta-registry-key.png)

## <a name="set-up-your-environment"></a><span data-ttu-id="70c4d-148">设置环境</span><span class="sxs-lookup"><span data-stu-id="70c4d-148">Set up your environment</span></span>

<span data-ttu-id="70c4d-149">完成[Outlook](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator)使用适用于加载项的 Yeoman 生成器创建加载项Office快速入门。</span><span class="sxs-lookup"><span data-stu-id="70c4d-149">Complete the [Outlook quick start](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) which creates an add-in project with the Yeoman generator for Office Add-ins.</span></span>

## <a name="configure-the-manifest"></a><span data-ttu-id="70c4d-150">配置清单</span><span class="sxs-lookup"><span data-stu-id="70c4d-150">Configure the manifest</span></span>

<span data-ttu-id="70c4d-151">若要启用加载项的基于事件的激活，必须在清单节点中配置 [Runtimes](../reference/manifest/runtimes.md) 元素和 [LaunchEvent](../reference/manifest/extensionpoint.md#launchevent) `VersionOverridesV1_1` 扩展点。</span><span class="sxs-lookup"><span data-stu-id="70c4d-151">To enable event-based activation of your add-in, you must configure the [Runtimes](../reference/manifest/runtimes.md) element and [LaunchEvent](../reference/manifest/extensionpoint.md#launchevent) extension point in the `VersionOverridesV1_1` node of the manifest.</span></span> <span data-ttu-id="70c4d-152">目前， `DesktopFormFactor` 是唯一受支持的外形类型。</span><span class="sxs-lookup"><span data-stu-id="70c4d-152">For now, `DesktopFormFactor` is the only supported form factor.</span></span>

1. <span data-ttu-id="70c4d-153">在代码编辑器中，打开快速启动项目。</span><span class="sxs-lookup"><span data-stu-id="70c4d-153">In your code editor, open the quick start project.</span></span>

1. <span data-ttu-id="70c4d-154">打开 **manifest.xml** 根目录下的文件。</span><span class="sxs-lookup"><span data-stu-id="70c4d-154">Open the **manifest.xml** file located at the root of your project.</span></span>

1. <span data-ttu-id="70c4d-155">选择整个节点 (包括打开和关闭) `<VersionOverrides>` 并将其替换为以下 XML，然后保存更改。</span><span class="sxs-lookup"><span data-stu-id="70c4d-155">Select the entire `<VersionOverrides>` node (including open and close tags) and replace it with the following XML, then save your changes.</span></span>

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
              <LaunchEvent Type="OnMessageAttachmentsChanged" FunctionName="onMessageAttachmentsChangedHandler" />
              <LaunchEvent Type="OnAppointmentAttachmentsChanged" FunctionName="onAppointmentAttachmentsChangedHandler" />
              <LaunchEvent Type="OnMessageRecipientsChanged" FunctionName="onMessageRecipientsChangedHandler" />
              <LaunchEvent Type="OnAppointmentAttendeesChanged" FunctionName="onAppointmentAttendeesChangedHandler" />
              <LaunchEvent Type="OnAppointmentTimeChanged" FunctionName="onAppointmentTimeChangedHandler" />
              <LaunchEvent Type="OnAppointmentRecurrenceChanged" FunctionName="onAppointmentRecurrenceChangedHandler" />
              <LaunchEvent Type="OnInfoBarDismissClicked" FunctionName="onInfobarDismissClickedHandler" />
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

<span data-ttu-id="70c4d-156">Outlook Windows使用 JavaScript 文件，Outlook Web 上的开发人员使用可以引用同一 JavaScript 文件的 HTML 文件。</span><span class="sxs-lookup"><span data-stu-id="70c4d-156">Outlook on Windows uses a JavaScript file, while Outlook on the web uses an HTML file that can reference the same JavaScript file.</span></span> <span data-ttu-id="70c4d-157">你必须在清单的节点中提供对这两个文件的引用，因为 Outlook 平台最终确定是使用 HTML 还是基于 Outlook `Resources` 客户端的 JavaScript。</span><span class="sxs-lookup"><span data-stu-id="70c4d-157">You must provide references to both these files in the `Resources` node of the manifest as the Outlook platform ultimately determines whether to use HTML or JavaScript based on the Outlook client.</span></span> <span data-ttu-id="70c4d-158">因此，若要配置事件处理，请提供 HTML 在 元素中的位置，然后在其子元素中提供 JavaScript 文件内附或 HTML `Runtime` `Override` 引用的位置。</span><span class="sxs-lookup"><span data-stu-id="70c4d-158">As such, to configure event handling, provide the location of the HTML in the `Runtime` element, then in its `Override` child element provide the location of the JavaScript file inlined or referenced by the HTML.</span></span>

> [!TIP]
> <span data-ttu-id="70c4d-159">若要了解有关加载项清单Outlook，请参阅Outlook[加载项清单](manifests.md)。</span><span class="sxs-lookup"><span data-stu-id="70c4d-159">To learn more about manifests for Outlook add-ins, see [Outlook add-in manifests](manifests.md).</span></span>

## <a name="implement-event-handling"></a><span data-ttu-id="70c4d-160">实现事件处理</span><span class="sxs-lookup"><span data-stu-id="70c4d-160">Implement event handling</span></span>

<span data-ttu-id="70c4d-161">您必须对所选事件实现处理。</span><span class="sxs-lookup"><span data-stu-id="70c4d-161">You have to implement handling for your selected events.</span></span>

<span data-ttu-id="70c4d-162">在此方案中，您将添加用于撰写新项的处理。</span><span class="sxs-lookup"><span data-stu-id="70c4d-162">In this scenario, you'll add handling for composing new items.</span></span>

1. <span data-ttu-id="70c4d-163">从同一快速启动项目中，在代码编辑器中打开 **commands.js./src/commands/commands.js** 文件。</span><span class="sxs-lookup"><span data-stu-id="70c4d-163">From the same quick start project, open the file **./src/commands/commands.js** in your code editor.</span></span>

1. <span data-ttu-id="70c4d-164">在 函数 `action` 之后，插入以下 JavaScript 函数。</span><span class="sxs-lookup"><span data-stu-id="70c4d-164">After the `action` function, insert the following JavaScript functions.</span></span>

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

1. <span data-ttu-id="70c4d-165">在文件末尾添加以下 JavaScript 代码。</span><span class="sxs-lookup"><span data-stu-id="70c4d-165">Add the following JavaScript code at the end of the file.</span></span>

    ```js
    // 1st parameter: FunctionName of LaunchEvent in the manifest; 2nd parameter: Its implementation in this .js file.
    Office.actions.associate("onMessageComposeHandler", onMessageComposeHandler);
    Office.actions.associate("onAppointmentComposeHandler", onAppointmentComposeHandler);
    ```

1. <span data-ttu-id="70c4d-166">保存所做的更改。</span><span class="sxs-lookup"><span data-stu-id="70c4d-166">Save your changes.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="70c4d-167">Windows：目前，在执行基于事件的激活的处理的 JavaScript 文件中不支持导入。</span><span class="sxs-lookup"><span data-stu-id="70c4d-167">Windows: At present, imports are not supported in the JavaScript file where you implement the handling for event-based activation.</span></span>

## <a name="try-it-out"></a><span data-ttu-id="70c4d-168">试用</span><span class="sxs-lookup"><span data-stu-id="70c4d-168">Try it out</span></span>

1. <span data-ttu-id="70c4d-169">在项目的根目录中运行以下命令。</span><span class="sxs-lookup"><span data-stu-id="70c4d-169">Run the following command in the root directory of your project.</span></span> <span data-ttu-id="70c4d-170">如果运行此命令，本地 Web 服务器将启动（如果尚未运行），并将旁加载加载项。</span><span class="sxs-lookup"><span data-stu-id="70c4d-170">When you run this command, the local web server will start (if it's not already running) and your add-in will be sideloaded.</span></span>

    ```command&nbsp;line
    npm start
    ```

    > [!NOTE]
    > <span data-ttu-id="70c4d-171">如果加载项未自动旁加载，请按照旁加载[Outlook](../outlook/sideload-outlook-add-ins-for-testing.md#sideload-manually)加载项进行测试中的说明，在加载项中手动旁加载Outlook。</span><span class="sxs-lookup"><span data-stu-id="70c4d-171">If your add-in wasn't automatically sideloaded, then follow the instructions in [Sideload Outlook add-ins for testing](../outlook/sideload-outlook-add-ins-for-testing.md#sideload-manually) to manually sideload the add-in in Outlook.</span></span>

1. <span data-ttu-id="70c4d-172">在 Outlook 网页版中，创建新邮件。</span><span class="sxs-lookup"><span data-stu-id="70c4d-172">In Outlook on the web, create a new message.</span></span>

    ![在撰写时设置主题Outlook网页中的邮件窗口屏幕截图](../images/outlook-web-autolaunch-1.png)

1. <span data-ttu-id="70c4d-174">在Outlook中Windows新建一封邮件。</span><span class="sxs-lookup"><span data-stu-id="70c4d-174">In Outlook on Windows, create a new message.</span></span>

    ![撰写时主题集Outlook Windows中邮件窗口的屏幕截图](../images/outlook-win-autolaunch.png)

    > [!NOTE]
    > <span data-ttu-id="70c4d-176">如果您从 localhost 运行您的外接程序，并且看到错误"很抱歉，我们无法访问 *{your-add-in-name-here}*。</span><span class="sxs-lookup"><span data-stu-id="70c4d-176">If you're running your add-in from localhost and see the error "We're sorry, we couldn't access *{your-add-in-name-here}*.</span></span> <span data-ttu-id="70c4d-177">请确保具有网络连接。</span><span class="sxs-lookup"><span data-stu-id="70c4d-177">Make sure you have a network connection.</span></span> <span data-ttu-id="70c4d-178">如果问题继续，请稍后重试。"，你可能需要启用环回豁免。</span><span class="sxs-lookup"><span data-stu-id="70c4d-178">If the problem continues, please try again later.", you may need to enable a loopback exemption.</span></span>
    >
    > 1. <span data-ttu-id="70c4d-179">关闭 Outlook。</span><span class="sxs-lookup"><span data-stu-id="70c4d-179">Close Outlook.</span></span>
    > 1. <span data-ttu-id="70c4d-180">打开 **任务管理器** ， **并确保msoadfsb.exe进程** 未运行。</span><span class="sxs-lookup"><span data-stu-id="70c4d-180">Open the **Task Manager** and ensure that the **msoadfsb.exe** process is not running.</span></span>
    > 1. <span data-ttu-id="70c4d-181">运行以下命令。</span><span class="sxs-lookup"><span data-stu-id="70c4d-181">Run the following command.</span></span>
    >
    >    ```command&nbsp;line
    >    call %SystemRoot%\System32\CheckNetIsolation.exe LoopbackExempt -a -n=1_http___localhost_300004ACA5EC-D79A-43EA-AB47-E50E47DD96FC
    >    ```
    >
    > 1. <span data-ttu-id="70c4d-182">重新启动 Outlook。</span><span class="sxs-lookup"><span data-stu-id="70c4d-182">Restart Outlook.</span></span>

## <a name="debug"></a><span data-ttu-id="70c4d-183">Debug</span><span class="sxs-lookup"><span data-stu-id="70c4d-183">Debug</span></span>

<span data-ttu-id="70c4d-184">在外接程序中对启动事件处理进行更改时，应注意：</span><span class="sxs-lookup"><span data-stu-id="70c4d-184">As you make changes to launch-event handling in your add-in, you should be aware that:</span></span>

- <span data-ttu-id="70c4d-185">如果更新了清单 [，请删除加载项，](sideload-outlook-add-ins-for-testing.md#remove-a-sideloaded-add-in) 然后再次旁加载它。</span><span class="sxs-lookup"><span data-stu-id="70c4d-185">If you updated the manifest, [remove the add-in](sideload-outlook-add-ins-for-testing.md#remove-a-sideloaded-add-in) then sideload it again.</span></span>
- <span data-ttu-id="70c4d-186">如果对清单外的文件进行了更改，请关闭并重新打开Outlook，Windows或刷新在 web 上Outlook浏览器选项卡。</span><span class="sxs-lookup"><span data-stu-id="70c4d-186">If you made changes to files other than the manifest, close and reopen Outlook on Windows, or refresh the browser tab running Outlook on the web.</span></span>

<span data-ttu-id="70c4d-187">实现自己的功能时，可能需要调试代码。</span><span class="sxs-lookup"><span data-stu-id="70c4d-187">While implementing your own functionality, you may need to debug your code.</span></span> <span data-ttu-id="70c4d-188">有关如何调试基于事件的外接程序激活的指南，请参阅调试基于事件Outlook[外接程序](debug-autolaunch.md)。</span><span class="sxs-lookup"><span data-stu-id="70c4d-188">For guidance on how to debug event-based add-in activation, see [Debug your event-based Outlook add-in](debug-autolaunch.md).</span></span>

<span data-ttu-id="70c4d-189">运行时日志记录还可用于 Windows。</span><span class="sxs-lookup"><span data-stu-id="70c4d-189">Runtime logging is also available for this feature on Windows.</span></span> <span data-ttu-id="70c4d-190">有关详细信息，请参阅使用运行时 [日志记录调试加载项](../testing/runtime-logging.md#runtime-logging-on-windows)。</span><span class="sxs-lookup"><span data-stu-id="70c4d-190">For more information, see [Debug your add-in with runtime logging](../testing/runtime-logging.md#runtime-logging-on-windows).</span></span>

## <a name="deploy-to-users"></a><span data-ttu-id="70c4d-191">部署到用户</span><span class="sxs-lookup"><span data-stu-id="70c4d-191">Deploy to users</span></span>

<span data-ttu-id="70c4d-192">通过管理中心上传清单，可以部署Microsoft 365加载项。</span><span class="sxs-lookup"><span data-stu-id="70c4d-192">You can deploy event-based add-ins by uploading the manifest through the Microsoft 365 admin center.</span></span> <span data-ttu-id="70c4d-193">在管理门户中，展开设置窗格中的"集成应用"部分，然后选择"**集成应用"。**</span><span class="sxs-lookup"><span data-stu-id="70c4d-193">In the admin portal, expand the **Settings** section in the navigation pane then select **Integrated apps**.</span></span> <span data-ttu-id="70c4d-194">在"**集成应用"** 页上，选择 **"Upload应用"** 操作。</span><span class="sxs-lookup"><span data-stu-id="70c4d-194">On the **Integrated apps** page, choose the **Upload custom apps** action.</span></span>

![管理中心上"集成应用"Microsoft 365的屏幕截图，包括Upload自定义应用操作](../images/outlook-deploy-event-based-add-ins.png)

<span data-ttu-id="70c4d-196">AppSource 和客户端存储：即将推出部署基于事件的加载项或更新现有加载项以包含基于事件的激活功能的功能。</span><span class="sxs-lookup"><span data-stu-id="70c4d-196">AppSource and inclient stores: The ability to deploy event-based add-ins or update existing add-ins to include the event-based activation feature should be available soon.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="70c4d-197">基于事件的外接程序仅限于管理员托管的部署。</span><span class="sxs-lookup"><span data-stu-id="70c4d-197">Event-based add-ins are restricted to admin-managed deployments only.</span></span> <span data-ttu-id="70c4d-198">目前，用户无法从 AppSource 或客户端应用商店获取基于事件的加载项。</span><span class="sxs-lookup"><span data-stu-id="70c4d-198">For now, users can't get event-based add-ins from AppSource or inclient stores.</span></span>

## <a name="event-based-activation-behavior-and-limitations"></a><span data-ttu-id="70c4d-199">基于事件的激活行为和限制</span><span class="sxs-lookup"><span data-stu-id="70c4d-199">Event-based activation behavior and limitations</span></span>

<span data-ttu-id="70c4d-200">加载项启动事件处理程序应尽量短运行、轻量且无影响。</span><span class="sxs-lookup"><span data-stu-id="70c4d-200">Add-in launch-event handlers are expected to be short-running, lightweight, and as noninvasive as possible.</span></span> <span data-ttu-id="70c4d-201">激活后，外接程序将在大约 300 秒（运行基于事件的外接程序所允许的最大时间长度）内退出。若要指示加载项已完成对启动事件的处理，我们建议让关联的处理程序调用 `event.completed` 方法。</span><span class="sxs-lookup"><span data-stu-id="70c4d-201">After activation, your add-in will time out within approximately 300 seconds, the maximum length of time allowed for running event-based add-ins. To signal that your add-in has completed processing a launch event, we recommend you have the associated handler call the `event.completed` method.</span></span> <span data-ttu-id="70c4d-202"> (请注意，语句后包含的代码不能保证运行。) 每次触发外接程序处理的事件时，外接程序将重新激活并运行关联的事件处理程序，超时窗口将重置。 `event.completed`</span><span class="sxs-lookup"><span data-stu-id="70c4d-202">(Note that code included after the `event.completed` statement is not guaranteed to run.) Each time an event that your add-in handles is triggered, the add-in is reactivated and runs the associated event handler, and the timeout window is reset.</span></span> <span data-ttu-id="70c4d-203">外接程序在时间结束后结束，或者用户关闭撰写窗口或发送项目。</span><span class="sxs-lookup"><span data-stu-id="70c4d-203">The add-in ends after it times out, or the user closes the compose window or sends the item.</span></span>

<span data-ttu-id="70c4d-204">如果用户有多个订阅了同一事件的加载项，Outlook平台将按特定顺序启动加载项。</span><span class="sxs-lookup"><span data-stu-id="70c4d-204">If the user has multiple add-ins that subscribed to the same event, the Outlook platform launches the add-ins in no particular order.</span></span> <span data-ttu-id="70c4d-205">目前，只能主动运行五个基于事件的加载项。</span><span class="sxs-lookup"><span data-stu-id="70c4d-205">Currently, only five event-based add-ins can be actively running.</span></span>

<span data-ttu-id="70c4d-206">用户可以切换或导航离开外接程序开始运行的当前邮件项目。</span><span class="sxs-lookup"><span data-stu-id="70c4d-206">The user can switch or navigate away from the current mail item where the add-in started running.</span></span> <span data-ttu-id="70c4d-207">启动的加载项将在后台完成其操作。</span><span class="sxs-lookup"><span data-stu-id="70c4d-207">The add-in that was launched will finish its operation in the background.</span></span>

<span data-ttu-id="70c4d-208">JavaScript 文件中不支持导入，在 JavaScript 文件中，您可以在 Windows 客户端中执行基于事件的激活。</span><span class="sxs-lookup"><span data-stu-id="70c4d-208">Imports are not supported in the JavaScript file where you implement the handling for event-based activation in the Windows client.</span></span>

<span data-ttu-id="70c4d-209">某些Office.js更改或更改 UI 的 API 不允许来自基于事件的外接程序。以下是阻止的 API：</span><span class="sxs-lookup"><span data-stu-id="70c4d-209">Some Office.js APIs that change or alter the UI are not allowed from event-based add-ins. The following are the blocked APIs:</span></span>

- <span data-ttu-id="70c4d-210">在 `OfficeRuntime.auth` 下：</span><span class="sxs-lookup"><span data-stu-id="70c4d-210">Under `OfficeRuntime.auth`:</span></span>
  - <span data-ttu-id="70c4d-211">`getAccessToken` (Windows仅) </span><span class="sxs-lookup"><span data-stu-id="70c4d-211">`getAccessToken` (Windows only)</span></span>
- <span data-ttu-id="70c4d-212">在 `Office.context.auth` 下：</span><span class="sxs-lookup"><span data-stu-id="70c4d-212">Under `Office.context.auth`:</span></span>
  - `getAccessToken`
  - `getAccessTokenAsync`
- <span data-ttu-id="70c4d-213">在 `Office.context.mailbox` 下：</span><span class="sxs-lookup"><span data-stu-id="70c4d-213">Under `Office.context.mailbox`:</span></span>
  - `displayAppointmentForm`
  - `displayMessageForm`
  - `displayNewAppointmentForm`
  - `displayNewMessageForm`
- <span data-ttu-id="70c4d-214">在 `Office.context.mailbox.item` 下：</span><span class="sxs-lookup"><span data-stu-id="70c4d-214">Under `Office.context.mailbox.item`:</span></span>
  - `close`
- <span data-ttu-id="70c4d-215">在 `Office.context.ui` 下：</span><span class="sxs-lookup"><span data-stu-id="70c4d-215">Under `Office.context.ui`:</span></span>
  - `displayDialogAsync`
  - `messageParent`

## <a name="see-also"></a><span data-ttu-id="70c4d-216">另请参阅</span><span class="sxs-lookup"><span data-stu-id="70c4d-216">See also</span></span>

- [<span data-ttu-id="70c4d-217">Outlook 加载项清单</span><span class="sxs-lookup"><span data-stu-id="70c4d-217">Outlook add-in manifests</span></span>](manifests.md)
- [<span data-ttu-id="70c4d-218">如何调试基于事件的外接程序</span><span class="sxs-lookup"><span data-stu-id="70c4d-218">How to debug event-based add-ins</span></span>](debug-autolaunch.md)
- <span data-ttu-id="70c4d-219">PnP 示例[：Outlook基于事件的激活设置签名](https://github.com/OfficeDev/PnP-OfficeAddins/tree/main/Samples/outlook-set-signature)</span><span class="sxs-lookup"><span data-stu-id="70c4d-219">PnP sample: [Use Outlook event-based activation to set the signature](https://github.com/OfficeDev/PnP-OfficeAddins/tree/main/Samples/outlook-set-signature)</span></span>