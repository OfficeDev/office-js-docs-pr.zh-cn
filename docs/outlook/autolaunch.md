---
title: '为 Outlook 外接程序配置基于事件的激活和 (预览) '
description: 了解如何为基于事件的激活配置 Outlook 外接程序。
ms.topic: article
ms.date: 02/03/2021
localization_priority: Normal
ms.openlocfilehash: a4fce335738d1bcff2be43e4e609998be89fca20
ms.sourcegitcommit: 8546889a759590c3798ce56e311d9e46f0171413
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/04/2021
ms.locfileid: "50104848"
---
# <a name="configure-your-outlook-add-in-for-event-based-activation-preview"></a><span data-ttu-id="1da94-103">为 Outlook 外接程序配置基于事件的激活和 (预览) </span><span class="sxs-lookup"><span data-stu-id="1da94-103">Configure your Outlook add-in for event-based activation (preview)</span></span>

<span data-ttu-id="1da94-104">如果没有基于事件的激活功能，用户必须显式启动加载项才能完成其任务。</span><span class="sxs-lookup"><span data-stu-id="1da94-104">Without the event-based activation feature, a user has to explicitly launch an add-in to complete their tasks.</span></span> <span data-ttu-id="1da94-105">利用此功能，您的外接程序可以基于特定事件运行任务，尤其是适用于每个项目的操作。</span><span class="sxs-lookup"><span data-stu-id="1da94-105">This feature enables your add-in to run tasks based on certain events, particularly for operations that apply to every item.</span></span> <span data-ttu-id="1da94-106">还可以与任务窗格和无 UI 功能集成。</span><span class="sxs-lookup"><span data-stu-id="1da94-106">You can also integrate with the task pane and UI-less functionality.</span></span> <span data-ttu-id="1da94-107">目前，支持以下事件。</span><span class="sxs-lookup"><span data-stu-id="1da94-107">At present, the following events are supported.</span></span>

- <span data-ttu-id="1da94-108">`OnNewMessageCompose`：撰写新邮件时 (包括答复、全部答复和转发) </span><span class="sxs-lookup"><span data-stu-id="1da94-108">`OnNewMessageCompose`: On composing a new message (includes reply, reply all, and forward)</span></span>
- <span data-ttu-id="1da94-109">`OnNewAppointmentOrganizer`：创建新约会时</span><span class="sxs-lookup"><span data-stu-id="1da94-109">`OnNewAppointmentOrganizer`: On creating a new appointment</span></span>

  > [!IMPORTANT]
  > <span data-ttu-id="1da94-110">在编辑 **项目** （例如草稿或现有约会）时，此功能不会激活。</span><span class="sxs-lookup"><span data-stu-id="1da94-110">This feature does **not** activate on editing an item, for example, a draft or an existing appointment.</span></span>

<span data-ttu-id="1da94-111">在此演练结束时，您将具有一个在新建邮件时运行的外接程序。</span><span class="sxs-lookup"><span data-stu-id="1da94-111">By the end of this walkthrough, you'll have an add-in that runs whenever a new message is created.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="1da94-112">此功能仅在 Outlook [网页版](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) 和具有 Microsoft 365 订阅的 Windows 中受支持预览。</span><span class="sxs-lookup"><span data-stu-id="1da94-112">This feature is only supported for [preview](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) in Outlook on the web and Windows with a Microsoft 365 subscription.</span></span> <span data-ttu-id="1da94-113">请参阅 [本文中](#how-to-preview-the-event-based-activation-feature) 如何预览基于事件的激活功能，了解更多详细信息。</span><span class="sxs-lookup"><span data-stu-id="1da94-113">See [How to preview the event-based activation feature](#how-to-preview-the-event-based-activation-feature) in this article for more details.</span></span>
>
> <span data-ttu-id="1da94-114">由于预览功能可能会随时更改，恕不另行通知，因此不应在生产外接程序中使用。</span><span class="sxs-lookup"><span data-stu-id="1da94-114">Because preview features are subject to change without notice, they shouldn't be used in production add-ins.</span></span>

## <a name="how-to-preview-the-event-based-activation-feature"></a><span data-ttu-id="1da94-115">如何预览基于事件的激活功能</span><span class="sxs-lookup"><span data-stu-id="1da94-115">How to preview the event-based activation feature</span></span>

<span data-ttu-id="1da94-116">我们邀请你试用基于事件的激活功能！</span><span class="sxs-lookup"><span data-stu-id="1da94-116">We invite you to try out the event-based activation feature!</span></span> <span data-ttu-id="1da94-117">通过 GitHub 提供反馈，让我们了解你的方案以及如何改进 (请参阅此页面末尾的"反馈"部分) 。 </span><span class="sxs-lookup"><span data-stu-id="1da94-117">Let us know your scenarios and how we can improve by giving us feedback through GitHub (see the **Feedback** section at the end of this page).</span></span>

<span data-ttu-id="1da94-118">预览此功能：</span><span class="sxs-lookup"><span data-stu-id="1da94-118">To preview this feature:</span></span>

- <span data-ttu-id="1da94-119">对于 Outlook 网页：</span><span class="sxs-lookup"><span data-stu-id="1da94-119">For Outlook on the web:</span></span>
  - <span data-ttu-id="1da94-120">[在 Microsoft 365 租户](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)上配置定向发布。</span><span class="sxs-lookup"><span data-stu-id="1da94-120">[Configure targeted release on your Microsoft 365 tenant](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center).</span></span>
  - <span data-ttu-id="1da94-121">引用 **CDN** 版本上的 beta https://appsforoffice.microsoft.com/lib/beta/hosted/office.js) (。</span><span class="sxs-lookup"><span data-stu-id="1da94-121">Reference the **beta** library on the CDN (https://appsforoffice.microsoft.com/lib/beta/hosted/office.js).</span></span> <span data-ttu-id="1da94-122">TypeScript [编译和](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts) 编译的类型IntelliSense CDN 和 [DefinitelyTyped 找到](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts)。</span><span class="sxs-lookup"><span data-stu-id="1da94-122">The [type definition file](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts) for TypeScript compilation and IntelliSense is found at the CDN and [DefinitelyTyped](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts).</span></span> <span data-ttu-id="1da94-123">可以安装这些类型 `npm install --save-dev @types/office-js-preview` 。</span><span class="sxs-lookup"><span data-stu-id="1da94-123">You can install these types with `npm install --save-dev @types/office-js-preview`.</span></span>
- <span data-ttu-id="1da94-124">对于 Windows 版 Outlook：最低要求版本为 16.0.13729.20000。</span><span class="sxs-lookup"><span data-stu-id="1da94-124">For Outlook on Windows: The minimum required build is 16.0.13729.20000.</span></span> <span data-ttu-id="1da94-125">加入 [Office 预览体验计划](https://insider.office.com) 以访问 Office beta 版本。</span><span class="sxs-lookup"><span data-stu-id="1da94-125">Join the [Office Insider program](https://insider.office.com) for access to Office beta builds.</span></span>

## <a name="set-up-your-environment"></a><span data-ttu-id="1da94-126">设置环境</span><span class="sxs-lookup"><span data-stu-id="1da94-126">Set up your environment</span></span>

<span data-ttu-id="1da94-127">使用 [适用于 Office](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) 加载项的 Yeoman 生成器完成创建外接程序项目的 Outlook 快速入门。</span><span class="sxs-lookup"><span data-stu-id="1da94-127">Complete the [Outlook quick start](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) which creates an add-in project with the Yeoman generator for Office Add-ins.</span></span>

## <a name="configure-the-manifest"></a><span data-ttu-id="1da94-128">配置清单</span><span class="sxs-lookup"><span data-stu-id="1da94-128">Configure the manifest</span></span>

<span data-ttu-id="1da94-129">若要启用加载项的基于事件的激活，必须在清单节点中配置 [Runtimes](../reference/manifest/runtimes.md) 元素和 [LaunchEvent](../reference/manifest/extensionpoint.md#launchevent-preview) `VersionOverridesV1_1` 扩展点。</span><span class="sxs-lookup"><span data-stu-id="1da94-129">To enable event-based activation of your add-in, you must configure the [Runtimes](../reference/manifest/runtimes.md) element and [LaunchEvent](../reference/manifest/extensionpoint.md#launchevent-preview) extension point in the `VersionOverridesV1_1` node of the manifest.</span></span> <span data-ttu-id="1da94-130">目前， `DesktopFormFactor` 是唯一受支持的外形类型。</span><span class="sxs-lookup"><span data-stu-id="1da94-130">For now, `DesktopFormFactor` is the only supported form factor.</span></span>

1. <span data-ttu-id="1da94-131">在代码编辑器中，打开快速启动项目。</span><span class="sxs-lookup"><span data-stu-id="1da94-131">In your code editor, open the quick start project.</span></span>

1. <span data-ttu-id="1da94-132">打开 **manifest.xml** 根目录下的文件。</span><span class="sxs-lookup"><span data-stu-id="1da94-132">Open the **manifest.xml** file located at the root of your project.</span></span>

1. <span data-ttu-id="1da94-133">选择整个 `<VersionOverrides>` 节点 (包括打开和关闭) 并将其替换为以下 XML。</span><span class="sxs-lookup"><span data-stu-id="1da94-133">Select the entire `<VersionOverrides>` node (including open and close tags) and replace it with the following XML.</span></span>

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

<span data-ttu-id="1da94-134">Windows 上的 Outlook 使用 JavaScript 文件，而 Web 上的 Outlook 使用可引用同一 JavaScript 文件的 HTML 文件。</span><span class="sxs-lookup"><span data-stu-id="1da94-134">Outlook on Windows uses a JavaScript file, while Outlook on the web uses an HTML file that can reference the same JavaScript file.</span></span> <span data-ttu-id="1da94-135">您必须在清单节点中提供对这两个文件的引用，因为 Outlook 平台最终决定是使用 HTML 还是基于 Outlook 客户端 `Resources` 的 JavaScript。</span><span class="sxs-lookup"><span data-stu-id="1da94-135">You must provide references to both these files in the `Resources` node of the manifest as the Outlook platform ultimately determines whether to use HTML or JavaScript based on the Outlook client.</span></span> <span data-ttu-id="1da94-136">因此，若要配置事件处理，请提供 HTML 在元素中的位置，然后在其子元素中提供 HTML 内附或引用 `Runtime` `Override` 的 JavaScript 文件的位置。</span><span class="sxs-lookup"><span data-stu-id="1da94-136">As such, to configure event handling, provide the location of the HTML in the `Runtime` element, then in its `Override` child element provide the location of the JavaScript file inlined or referenced by the HTML.</span></span>

> [!TIP]
> <span data-ttu-id="1da94-137">若要了解有关 Outlook 外接程序清单的更多信息，请参阅 [Outlook 外接程序清单](manifests.md)。</span><span class="sxs-lookup"><span data-stu-id="1da94-137">To learn more about manifests for Outlook add-ins, see [Outlook add-in manifests](manifests.md).</span></span>

## <a name="implement-event-handling"></a><span data-ttu-id="1da94-138">实现事件处理</span><span class="sxs-lookup"><span data-stu-id="1da94-138">Implement event handling</span></span>

<span data-ttu-id="1da94-139">您必须对选定的事件实现处理。</span><span class="sxs-lookup"><span data-stu-id="1da94-139">You have to implement handling for your selected events.</span></span>

<span data-ttu-id="1da94-140">在此方案中，将添加用于撰写新项的处理。</span><span class="sxs-lookup"><span data-stu-id="1da94-140">In this scenario, you'll add handling for composing new items.</span></span>

1. <span data-ttu-id="1da94-141">从同一快速启动项目中，在代码编辑器中打开commands.js **./src/commands/commands.js** 文件。</span><span class="sxs-lookup"><span data-stu-id="1da94-141">From the same quick start project, open the file **./src/commands/commands.js** in your code editor.</span></span>

1. <span data-ttu-id="1da94-142">在函数 `action` 后插入以下 JavaScript 函数。</span><span class="sxs-lookup"><span data-stu-id="1da94-142">After the `action` function, insert the following JavaScript functions.</span></span>

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

1. <span data-ttu-id="1da94-143">若要使用由 Office 加载项的 Yeoman 生成器生成的此项目在 **Outlook** 网页 Outlook 中运行的函数，在文件末尾添加以下语句。</span><span class="sxs-lookup"><span data-stu-id="1da94-143">For the functions to work in **Outlook on the web** with this project generated by the Yeoman generator for Office Add-ins, add the following statements at the end of the file.</span></span>

    ```js
    g.onMessageComposeHandler = onMessageComposeHandler;
    g.onAppointmentComposeHandler = onAppointmentComposeHandler;
    ```

1. <span data-ttu-id="1da94-144">若要使函数在 Windows 上的 **Outlook 中运行**，在文件末尾添加以下 JavaScript 代码。</span><span class="sxs-lookup"><span data-stu-id="1da94-144">For the functions to work in **Outlook on Windows**, add the following JavaScript code at the end of the file.</span></span>

    ```js
    if (Office.actions) {
      // 1st parameter: FunctionName of LaunchEvent in the manifest; 2nd parameter: Its implementation in this .js file.
      Office.actions.associate("onMessageComposeHandler", onMessageComposeHandler);
      Office.actions.associate("onAppointmentComposeHandler", onAppointmentComposeHandler);
    }
    ```

    <span data-ttu-id="1da94-145">**注意**：检查 `Office.actions` 以确保 Web 上的 Outlook 忽略这些语句。</span><span class="sxs-lookup"><span data-stu-id="1da94-145">**Note**: Checking for `Office.actions` ensures that Outlook on the web ignores these statements.</span></span>

## <a name="try-it-out"></a><span data-ttu-id="1da94-146">试用</span><span class="sxs-lookup"><span data-stu-id="1da94-146">Try it out</span></span>

1. <span data-ttu-id="1da94-147">在项目的根目录中运行以下命令。</span><span class="sxs-lookup"><span data-stu-id="1da94-147">Run the following command in the root directory of your project.</span></span> <span data-ttu-id="1da94-148">运行此命令时，本地 Web 服务器将启动（如果尚未运行）。</span><span class="sxs-lookup"><span data-stu-id="1da94-148">When you run this command, the local web server will start (if it's not already running).</span></span>

    ```command&nbsp;line
    npm run dev-server
    ```

1. <span data-ttu-id="1da94-149">按照[旁加载 Outlook 加载项以供测试](sideload-outlook-add-ins-for-testing.md)中的说明操作，旁加载加载项。</span><span class="sxs-lookup"><span data-stu-id="1da94-149">Follow the instructions in [Sideload Outlook add-ins for testing](sideload-outlook-add-ins-for-testing.md) to sideload the add-in in Outlook.</span></span>

1. <span data-ttu-id="1da94-150">在 Outlook 网页版中，创建新邮件。</span><span class="sxs-lookup"><span data-stu-id="1da94-150">In Outlook on the web, create a new message.</span></span>

    ![Outlook 网页中的邮件窗口的屏幕截图，撰写时主题已设置](../images/outlook-web-autolaunch-1.png)

1. <span data-ttu-id="1da94-152">在 Windows 上的 Outlook 中，创建新邮件。</span><span class="sxs-lookup"><span data-stu-id="1da94-152">In Outlook on Windows, create a new message.</span></span>

    ![在撰写时设置了主题的 Windows 上的 Outlook 中邮件窗口的屏幕截图](../images/outlook-win-autolaunch.png)

## <a name="event-based-activation-behavior-and-limitations"></a><span data-ttu-id="1da94-154">基于事件的激活行为和限制</span><span class="sxs-lookup"><span data-stu-id="1da94-154">Event-based activation behavior and limitations</span></span>

<span data-ttu-id="1da94-155">基于事件激活的外接程序应尽可能短运行、轻型且非高速度。</span><span class="sxs-lookup"><span data-stu-id="1da94-155">Add-ins that activate based on events are expected to be short-running, lightweight, and as non-invasive as possible.</span></span> <span data-ttu-id="1da94-156">若要表明加载项已完成对启动事件的处理，我们建议你让加载项调用 `event.completed` 该方法。</span><span class="sxs-lookup"><span data-stu-id="1da94-156">To signal that your add-in has completed processing the launch event, we recommend you have your add-in call the `event.completed` method.</span></span> <span data-ttu-id="1da94-157">如果未进行该调用，加载项将在大约 300 秒（运行基于事件的加载项所允许的最大时间长度）内退出。当用户关闭撰写窗口时，外接程序也会结束。</span><span class="sxs-lookup"><span data-stu-id="1da94-157">If that call is not made, the add-in will time out within approximately 300 seconds, the maximum length of time allowed for running event-based add-ins. The add-in also ends when the user closes the compose window.</span></span>

<span data-ttu-id="1da94-158">如果用户有多个订阅同一事件的加载项，Outlook 平台将按特定顺序启动外接程序。</span><span class="sxs-lookup"><span data-stu-id="1da94-158">If the user has multiple add-ins that subscribed to the same event, the Outlook platform launches the add-ins in no particular order.</span></span> <span data-ttu-id="1da94-159">目前，只能主动运行五个基于事件的加载项。</span><span class="sxs-lookup"><span data-stu-id="1da94-159">Currently, only five event-based add-ins can be actively running.</span></span> <span data-ttu-id="1da94-160">任何其他加载项将推送到队列，然后随着之前处于活动状态的加载项完成或停用而运行。</span><span class="sxs-lookup"><span data-stu-id="1da94-160">Any additional add-ins are pushed to a queue then run as previously active add-ins are completed or deactivated.</span></span>

<span data-ttu-id="1da94-161">用户可以切换或导航离开外接程序开始运行的当前邮件项目。</span><span class="sxs-lookup"><span data-stu-id="1da94-161">The user can switch or navigate away from the current mail item where the add-in started running.</span></span> <span data-ttu-id="1da94-162">启动的加载项将在后台完成其操作。</span><span class="sxs-lookup"><span data-stu-id="1da94-162">The add-in that was launched will finish its operation in the background.</span></span>

<span data-ttu-id="1da94-163">某些Office.js更改或更改 UI 的 API 不允许来自基于事件的加载项。以下是阻止的 API：</span><span class="sxs-lookup"><span data-stu-id="1da94-163">Some Office.js APIs that change or alter the UI are not allowed from event-based add-ins. The following are the blocked APIs:</span></span>

- <span data-ttu-id="1da94-164">Under `Office.context.mailbox` ：</span><span class="sxs-lookup"><span data-stu-id="1da94-164">Under `Office.context.mailbox`:</span></span>
  - `displayAppointmentForm`
  - `displayMessageForm`
  - `displayNewAppointmentForm`
  - `displayNewMessageForm`
- <span data-ttu-id="1da94-165">Under `Office.context.ui` ：</span><span class="sxs-lookup"><span data-stu-id="1da94-165">Under `Office.context.ui`:</span></span>
  - `displayDialogAsync`
  - `messageParent`
- <span data-ttu-id="1da94-166">Under `Office.context.auth` ：</span><span class="sxs-lookup"><span data-stu-id="1da94-166">Under `Office.context.auth`:</span></span>
  - `getAccessToken`
  - `getAccessTokenAsync`

## <a name="see-also"></a><span data-ttu-id="1da94-167">另请参阅</span><span class="sxs-lookup"><span data-stu-id="1da94-167">See also</span></span>

[<span data-ttu-id="1da94-168">Outlook 加载项清单</span><span class="sxs-lookup"><span data-stu-id="1da94-168">Outlook add-in manifests</span></span>](manifests.md)