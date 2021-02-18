---
title: '为 Outlook 外接程序配置基于事件的激活 (预览) '
description: 了解如何配置 Outlook 外接程序进行基于事件的激活。
ms.topic: article
ms.date: 02/12/2021
localization_priority: Normal
ms.openlocfilehash: a3e2167adec824934d1bc20d0e6613f9057e5c70
ms.sourcegitcommit: 7cd501d0fdbbd4636bd08647b638dd5ca4c7c630
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/17/2021
ms.locfileid: "50282994"
---
# <a name="configure-your-outlook-add-in-for-event-based-activation-preview"></a><span data-ttu-id="0f327-103">为 Outlook 外接程序配置基于事件的激活 (预览) </span><span class="sxs-lookup"><span data-stu-id="0f327-103">Configure your Outlook add-in for event-based activation (preview)</span></span>

<span data-ttu-id="0f327-104">如果没有基于事件的激活功能，用户必须显式启动加载项才能完成其任务。</span><span class="sxs-lookup"><span data-stu-id="0f327-104">Without the event-based activation feature, a user has to explicitly launch an add-in to complete their tasks.</span></span> <span data-ttu-id="0f327-105">利用此功能，加载项能够基于特定事件运行任务，尤其是适用于每个项目的操作。</span><span class="sxs-lookup"><span data-stu-id="0f327-105">This feature enables your add-in to run tasks based on certain events, particularly for operations that apply to every item.</span></span> <span data-ttu-id="0f327-106">还可以与任务窗格和无 UI 功能集成。</span><span class="sxs-lookup"><span data-stu-id="0f327-106">You can also integrate with the task pane and UI-less functionality.</span></span> <span data-ttu-id="0f327-107">目前，支持以下事件。</span><span class="sxs-lookup"><span data-stu-id="0f327-107">At present, the following events are supported.</span></span>

- <span data-ttu-id="0f327-108">`OnNewMessageCompose`：撰写新邮件时 (包括答复、全部答复和转发) </span><span class="sxs-lookup"><span data-stu-id="0f327-108">`OnNewMessageCompose`: On composing a new message (includes reply, reply all, and forward)</span></span>
- <span data-ttu-id="0f327-109">`OnNewAppointmentOrganizer`：创建新约会时</span><span class="sxs-lookup"><span data-stu-id="0f327-109">`OnNewAppointmentOrganizer`: On creating a new appointment</span></span>

  > [!IMPORTANT]
  > <span data-ttu-id="0f327-110">在编辑 **项目** （例如草稿或现有约会）时，此功能不会激活。</span><span class="sxs-lookup"><span data-stu-id="0f327-110">This feature does **not** activate on editing an item, for example, a draft or an existing appointment.</span></span>

<span data-ttu-id="0f327-111">在此演练结束时，您将拥有一个在新建邮件时运行的外接程序。</span><span class="sxs-lookup"><span data-stu-id="0f327-111">By the end of this walkthrough, you'll have an add-in that runs whenever a new message is created.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="0f327-112">此功能仅在 Outlook [网页版](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) 和具有 Microsoft 365 订阅的 Windows 上受支持预览。</span><span class="sxs-lookup"><span data-stu-id="0f327-112">This feature is only supported for [preview](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) in Outlook on the web and on Windows with a Microsoft 365 subscription.</span></span> <span data-ttu-id="0f327-113">有关详细信息 [，请参阅](#how-to-preview-the-event-based-activation-feature) 本文中如何预览基于事件的激活功能。</span><span class="sxs-lookup"><span data-stu-id="0f327-113">See [How to preview the event-based activation feature](#how-to-preview-the-event-based-activation-feature) in this article for more details.</span></span>
>
> <span data-ttu-id="0f327-114">由于预览功能可能会随时更改，恕不另行通知，因此不应将其用于生产外接程序。</span><span class="sxs-lookup"><span data-stu-id="0f327-114">Because preview features are subject to change without notice, they shouldn't be used in production add-ins.</span></span>

## <a name="how-to-preview-the-event-based-activation-feature"></a><span data-ttu-id="0f327-115">如何预览基于事件的激活功能</span><span class="sxs-lookup"><span data-stu-id="0f327-115">How to preview the event-based activation feature</span></span>

<span data-ttu-id="0f327-116">我们邀请你试用基于事件的激活功能！</span><span class="sxs-lookup"><span data-stu-id="0f327-116">We invite you to try out the event-based activation feature!</span></span> <span data-ttu-id="0f327-117">请告诉我们你的方案以及如何通过 GitHub 提供反馈来改进 (请参阅此页面末尾的"反馈"部分) 。 </span><span class="sxs-lookup"><span data-stu-id="0f327-117">Let us know your scenarios and how we can improve by giving us feedback through GitHub (see the **Feedback** section at the end of this page).</span></span>

<span data-ttu-id="0f327-118">预览此功能：</span><span class="sxs-lookup"><span data-stu-id="0f327-118">To preview this feature:</span></span>

- <span data-ttu-id="0f327-119">对于 Outlook 网页：</span><span class="sxs-lookup"><span data-stu-id="0f327-119">For Outlook on the web:</span></span>
  - <span data-ttu-id="0f327-120">[在 Microsoft 365 租户上配置定向版本](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)。</span><span class="sxs-lookup"><span data-stu-id="0f327-120">[Configure targeted release on your Microsoft 365 tenant](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center).</span></span>
  - <span data-ttu-id="0f327-121">在 **CDN** 服务器上引用 beta https://appsforoffice.microsoft.com/lib/beta/hosted/office.js) (。</span><span class="sxs-lookup"><span data-stu-id="0f327-121">Reference the **beta** library on the CDN (https://appsforoffice.microsoft.com/lib/beta/hosted/office.js).</span></span> <span data-ttu-id="0f327-122">TypeScript [编译和键入](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts) IntelliSense在 CDN 和 [DefinitelyTyped 上找到](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts)。</span><span class="sxs-lookup"><span data-stu-id="0f327-122">The [type definition file](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts) for TypeScript compilation and IntelliSense is found at the CDN and [DefinitelyTyped](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts).</span></span> <span data-ttu-id="0f327-123">可以使用 .安装这些类型 `npm install --save-dev @types/office-js-preview` 。</span><span class="sxs-lookup"><span data-stu-id="0f327-123">You can install these types with `npm install --save-dev @types/office-js-preview`.</span></span>
- <span data-ttu-id="0f327-124">对于 Windows 版 Outlook：最低要求版本为 16.0.13729.20000。</span><span class="sxs-lookup"><span data-stu-id="0f327-124">For Outlook on Windows: The minimum required build is 16.0.13729.20000.</span></span> <span data-ttu-id="0f327-125">加入 [Office 预览体验计划](https://insider.office.com) 以访问 Office beta 版本。</span><span class="sxs-lookup"><span data-stu-id="0f327-125">Join the [Office Insider program](https://insider.office.com) for access to Office beta builds.</span></span>

## <a name="set-up-your-environment"></a><span data-ttu-id="0f327-126">设置环境</span><span class="sxs-lookup"><span data-stu-id="0f327-126">Set up your environment</span></span>

<span data-ttu-id="0f327-127">使用 [适用于 Office](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) 加载项的 Yeoman 生成器完成创建外接程序项目的 Outlook 快速入门。</span><span class="sxs-lookup"><span data-stu-id="0f327-127">Complete the [Outlook quick start](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) which creates an add-in project with the Yeoman generator for Office Add-ins.</span></span>

## <a name="configure-the-manifest"></a><span data-ttu-id="0f327-128">配置清单</span><span class="sxs-lookup"><span data-stu-id="0f327-128">Configure the manifest</span></span>

<span data-ttu-id="0f327-129">若要启用加载项的基于事件的激活，必须在清单节点中配置 [Runtimes](../reference/manifest/runtimes.md) 元素和 [LaunchEvent](../reference/manifest/extensionpoint.md#launchevent-preview) `VersionOverridesV1_1` 扩展点。</span><span class="sxs-lookup"><span data-stu-id="0f327-129">To enable event-based activation of your add-in, you must configure the [Runtimes](../reference/manifest/runtimes.md) element and [LaunchEvent](../reference/manifest/extensionpoint.md#launchevent-preview) extension point in the `VersionOverridesV1_1` node of the manifest.</span></span> <span data-ttu-id="0f327-130">目前， `DesktopFormFactor` 是唯一受支持的外形类型。</span><span class="sxs-lookup"><span data-stu-id="0f327-130">For now, `DesktopFormFactor` is the only supported form factor.</span></span>

1. <span data-ttu-id="0f327-131">在代码编辑器中，打开快速启动项目。</span><span class="sxs-lookup"><span data-stu-id="0f327-131">In your code editor, open the quick start project.</span></span>

1. <span data-ttu-id="0f327-132">打开 **manifest.xml** 根目录下的文件。</span><span class="sxs-lookup"><span data-stu-id="0f327-132">Open the **manifest.xml** file located at the root of your project.</span></span>

1. <span data-ttu-id="0f327-133">选择整个节点 (包括打开和关闭) `<VersionOverrides>` 并将其替换为以下 XML，然后保存更改。</span><span class="sxs-lookup"><span data-stu-id="0f327-133">Select the entire `<VersionOverrides>` node (including open and close tags) and replace it with the following XML, then save your changes.</span></span>

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

<span data-ttu-id="0f327-134">Windows 上的 Outlook 使用 JavaScript 文件，而 Web 上的 Outlook 使用可引用同一 JavaScript 文件的 HTML 文件。</span><span class="sxs-lookup"><span data-stu-id="0f327-134">Outlook on Windows uses a JavaScript file, while Outlook on the web uses an HTML file that can reference the same JavaScript file.</span></span> <span data-ttu-id="0f327-135">您必须在清单节点中提供对这两个文件的引用，因为 Outlook 平台最终确定是使用基于 Outlook 客户端的 HTML 还是 `Resources` JavaScript。</span><span class="sxs-lookup"><span data-stu-id="0f327-135">You must provide references to both these files in the `Resources` node of the manifest as the Outlook platform ultimately determines whether to use HTML or JavaScript based on the Outlook client.</span></span> <span data-ttu-id="0f327-136">因此，若要配置事件处理，请提供 HTML 在元素中的位置，然后在其子元素中提供 HTML 内附或引用 `Runtime` `Override` 的 JavaScript 文件的位置。</span><span class="sxs-lookup"><span data-stu-id="0f327-136">As such, to configure event handling, provide the location of the HTML in the `Runtime` element, then in its `Override` child element provide the location of the JavaScript file inlined or referenced by the HTML.</span></span>

> [!TIP]
> <span data-ttu-id="0f327-137">若要了解有关 Outlook 外接程序清单的更多信息，请参阅 [Outlook 外接程序清单](manifests.md)。</span><span class="sxs-lookup"><span data-stu-id="0f327-137">To learn more about manifests for Outlook add-ins, see [Outlook add-in manifests](manifests.md).</span></span>

## <a name="implement-event-handling"></a><span data-ttu-id="0f327-138">实现事件处理</span><span class="sxs-lookup"><span data-stu-id="0f327-138">Implement event handling</span></span>

<span data-ttu-id="0f327-139">您必须对所选事件实现处理。</span><span class="sxs-lookup"><span data-stu-id="0f327-139">You have to implement handling for your selected events.</span></span>

<span data-ttu-id="0f327-140">在此方案中，您将添加用于撰写新项的处理。</span><span class="sxs-lookup"><span data-stu-id="0f327-140">In this scenario, you'll add handling for composing new items.</span></span>

1. <span data-ttu-id="0f327-141">从同一快速启动项目中，在代码编辑器中commands.js **./src/commands/commands.js** 文件。</span><span class="sxs-lookup"><span data-stu-id="0f327-141">From the same quick start project, open the file **./src/commands/commands.js** in your code editor.</span></span>

1. <span data-ttu-id="0f327-142">在函数 `action` 后插入以下 JavaScript 函数。</span><span class="sxs-lookup"><span data-stu-id="0f327-142">After the `action` function, insert the following JavaScript functions.</span></span>

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

1. <span data-ttu-id="0f327-143">若要使用由 Office 外接程序的 Yeoman 生成器生成的此项目在 **Outlook** 网页 Outlook 中运行的函数，在文件末尾添加以下语句。</span><span class="sxs-lookup"><span data-stu-id="0f327-143">For the functions to work in **Outlook on the web** with this project generated by the Yeoman generator for Office Add-ins, add the following statements at the end of the file.</span></span>

    ```js
    g.onMessageComposeHandler = onMessageComposeHandler;
    g.onAppointmentComposeHandler = onAppointmentComposeHandler;
    ```

1. <span data-ttu-id="0f327-144">若要在 Windows 上的 **Outlook 中** 运行函数，在文件末尾添加以下 JavaScript 代码。</span><span class="sxs-lookup"><span data-stu-id="0f327-144">For the functions to work in **Outlook on Windows**, add the following JavaScript code at the end of the file.</span></span>

    ```js
    if (Office.actions) {
      // 1st parameter: FunctionName of LaunchEvent in the manifest; 2nd parameter: Its implementation in this .js file.
      Office.actions.associate("onMessageComposeHandler", onMessageComposeHandler);
      Office.actions.associate("onAppointmentComposeHandler", onAppointmentComposeHandler);
    }
    ```

    <span data-ttu-id="0f327-145">**注意**：检查 `Office.actions` 以确保 Web 上的 Outlook 忽略这些语句。</span><span class="sxs-lookup"><span data-stu-id="0f327-145">**Note**: Checking for `Office.actions` ensures that Outlook on the web ignores these statements.</span></span>

1. <span data-ttu-id="0f327-146">保存所做的更改。</span><span class="sxs-lookup"><span data-stu-id="0f327-146">Save your changes.</span></span>

## <a name="try-it-out"></a><span data-ttu-id="0f327-147">试用</span><span class="sxs-lookup"><span data-stu-id="0f327-147">Try it out</span></span>

1. <span data-ttu-id="0f327-148">在项目的根目录中运行以下命令。</span><span class="sxs-lookup"><span data-stu-id="0f327-148">Run the following command in the root directory of your project.</span></span> <span data-ttu-id="0f327-149">如果运行此命令，本地 Web 服务器将启动（如果尚未运行），并将旁加载加载项。</span><span class="sxs-lookup"><span data-stu-id="0f327-149">When you run this command, the local web server will start (if it's not already running) and your add-in will be sideloaded.</span></span>

    ```command&nbsp;line
    npm start
    ```

1. <span data-ttu-id="0f327-150">在 Outlook 网页版中，创建新邮件。</span><span class="sxs-lookup"><span data-stu-id="0f327-150">In Outlook on the web, create a new message.</span></span>

    ![Outlook 网页邮件窗口的屏幕截图，撰写时主题已设置](../images/outlook-web-autolaunch-1.png)

1. <span data-ttu-id="0f327-152">在 Windows 上的 Outlook 中，创建新邮件。</span><span class="sxs-lookup"><span data-stu-id="0f327-152">In Outlook on Windows, create a new message.</span></span>

    ![Windows 上的 Outlook 中邮件窗口的屏幕截图，撰写时主题已设置](../images/outlook-win-autolaunch.png)

## <a name="debug"></a><span data-ttu-id="0f327-154">Debug</span><span class="sxs-lookup"><span data-stu-id="0f327-154">Debug</span></span>

<span data-ttu-id="0f327-155">当你实现自己的功能时，你可能需要调试代码。</span><span class="sxs-lookup"><span data-stu-id="0f327-155">As you implement your own functionality, you may need to debug your code.</span></span> <span data-ttu-id="0f327-156">有关如何调试基于事件的外接程序激活的指南，请参阅"调试基于事件的[Outlook 外接程序"。](debug-autolaunch.md)</span><span class="sxs-lookup"><span data-stu-id="0f327-156">For guidance on how to debug event-based add-in activation, see [Debug your event-based Outlook add-in](debug-autolaunch.md).</span></span>

## <a name="event-based-activation-behavior-and-limitations"></a><span data-ttu-id="0f327-157">基于事件的激活行为和限制</span><span class="sxs-lookup"><span data-stu-id="0f327-157">Event-based activation behavior and limitations</span></span>

<span data-ttu-id="0f327-158">基于事件激活的加载项应尽可能短运行、轻型和非高空。</span><span class="sxs-lookup"><span data-stu-id="0f327-158">Add-ins that activate based on events are expected to be short-running, lightweight, and as non-invasive as possible.</span></span> <span data-ttu-id="0f327-159">若要指示加载项已完成对启动事件的处理，建议让加载项调用 `event.completed` 该方法。</span><span class="sxs-lookup"><span data-stu-id="0f327-159">To signal that your add-in has completed processing the launch event, we recommend you have your add-in call the `event.completed` method.</span></span> <span data-ttu-id="0f327-160">如果未进行该调用，加载项将在大约 300 秒（运行基于事件的加载项所允许的最大时间长度）内退出。当用户关闭撰写窗口时，加载项也会结束。</span><span class="sxs-lookup"><span data-stu-id="0f327-160">If that call is not made, the add-in will time out within approximately 300 seconds, the maximum length of time allowed for running event-based add-ins. The add-in also ends when the user closes the compose window.</span></span>

<span data-ttu-id="0f327-161">如果用户有多个订阅同一事件的加载项，则 Outlook 平台将启动外接程序，而没有任何特定顺序。</span><span class="sxs-lookup"><span data-stu-id="0f327-161">If the user has multiple add-ins that subscribed to the same event, the Outlook platform launches the add-ins in no particular order.</span></span> <span data-ttu-id="0f327-162">目前，只能主动运行五个基于事件的加载项。</span><span class="sxs-lookup"><span data-stu-id="0f327-162">Currently, only five event-based add-ins can be actively running.</span></span> <span data-ttu-id="0f327-163">任何其他加载项将推送到队列，然后随着之前处于活动状态的加载项完成或停用而运行。</span><span class="sxs-lookup"><span data-stu-id="0f327-163">Any additional add-ins are pushed to a queue then run as previously active add-ins are completed or deactivated.</span></span>

<span data-ttu-id="0f327-164">用户可以切换或导航离开加载项开始运行的当前邮件项目。</span><span class="sxs-lookup"><span data-stu-id="0f327-164">The user can switch or navigate away from the current mail item where the add-in started running.</span></span> <span data-ttu-id="0f327-165">启动的加载项将在后台完成其操作。</span><span class="sxs-lookup"><span data-stu-id="0f327-165">The add-in that was launched will finish its operation in the background.</span></span>

<span data-ttu-id="0f327-166">某些Office.js更改或更改 UI 的 API 不允许来自基于事件的加载项。以下是阻止的 API：</span><span class="sxs-lookup"><span data-stu-id="0f327-166">Some Office.js APIs that change or alter the UI are not allowed from event-based add-ins. The following are the blocked APIs:</span></span>

- <span data-ttu-id="0f327-167">下 `Office.context.auth` ：</span><span class="sxs-lookup"><span data-stu-id="0f327-167">Under `Office.context.auth`:</span></span>
  - `getAccessToken`
  - `getAccessTokenAsync`
- <span data-ttu-id="0f327-168">下 `Office.context.mailbox` ：</span><span class="sxs-lookup"><span data-stu-id="0f327-168">Under `Office.context.mailbox`:</span></span>
  - `displayAppointmentForm`
  - `displayMessageForm`
  - `displayNewAppointmentForm`
  - `displayNewMessageForm`
- <span data-ttu-id="0f327-169">下 `Office.context.mailbox.item` ：</span><span class="sxs-lookup"><span data-stu-id="0f327-169">Under `Office.context.mailbox.item`:</span></span>
  - `close`
- <span data-ttu-id="0f327-170">下 `Office.context.ui` ：</span><span class="sxs-lookup"><span data-stu-id="0f327-170">Under `Office.context.ui`:</span></span>
  - `displayDialogAsync`
  - `messageParent`

## <a name="see-also"></a><span data-ttu-id="0f327-171">另请参阅</span><span class="sxs-lookup"><span data-stu-id="0f327-171">See also</span></span>

<span data-ttu-id="0f327-172">[Outlook 外接程序清单](manifests.md) 
[如何调试基于事件的加载项](debug-autolaunch.md)</span><span class="sxs-lookup"><span data-stu-id="0f327-172">[Outlook add-in manifests](manifests.md)
[How to debug event-based add-ins](debug-autolaunch.md)</span></span>
