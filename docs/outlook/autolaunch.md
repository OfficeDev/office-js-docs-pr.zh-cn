---
title: '为 Outlook 外接程序配置基于事件的激活和 (预览) '
description: 了解如何为基于事件的激活配置 Outlook 外接程序。
ms.topic: article
ms.date: 01/06/2021
localization_priority: Normal
ms.openlocfilehash: d6893733af52bba7917531b2e8d5a442ce3dcd77
ms.sourcegitcommit: d28392721958555d6edea48cea000470bd27fcf7
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/13/2021
ms.locfileid: "49839829"
---
# <a name="configure-your-outlook-add-in-for-event-based-activation-preview"></a><span data-ttu-id="2ef61-103">为 Outlook 外接程序配置基于事件的激活和 (预览) </span><span class="sxs-lookup"><span data-stu-id="2ef61-103">Configure your Outlook add-in for event-based activation (preview)</span></span>

<span data-ttu-id="2ef61-104">如果没有基于事件的激活功能，用户必须显式启动加载项才能完成其任务。</span><span class="sxs-lookup"><span data-stu-id="2ef61-104">Without the event-based activation feature, a user has to explicitly launch an add-in to complete their tasks.</span></span> <span data-ttu-id="2ef61-105">利用此功能，您的外接程序可以基于特定事件运行任务，尤其是适用于每个项目的操作。</span><span class="sxs-lookup"><span data-stu-id="2ef61-105">This feature enables your add-in to run tasks based on certain events, particularly for operations that apply to every item.</span></span> <span data-ttu-id="2ef61-106">还可以与任务窗格和无 UI 功能集成。</span><span class="sxs-lookup"><span data-stu-id="2ef61-106">You can also integrate with the task pane and UI-less functionality.</span></span> <span data-ttu-id="2ef61-107">目前，支持以下事件。</span><span class="sxs-lookup"><span data-stu-id="2ef61-107">At present, the following events are supported.</span></span>

- <span data-ttu-id="2ef61-108">`OnNewMessageCompose`：撰写新邮件时 (包括答复、全部答复和转发) </span><span class="sxs-lookup"><span data-stu-id="2ef61-108">`OnNewMessageCompose`: On composing a new message (includes reply, reply all, and forward)</span></span>
- <span data-ttu-id="2ef61-109">`OnNewAppointmentOrganizer`：创建新约会时</span><span class="sxs-lookup"><span data-stu-id="2ef61-109">`OnNewAppointmentOrganizer`: On creating a new appointment</span></span>

  > [!IMPORTANT]
  > <span data-ttu-id="2ef61-110">在编辑 **项目** （例如草稿或现有约会）时，此功能不会激活。</span><span class="sxs-lookup"><span data-stu-id="2ef61-110">This feature does **not** activate on editing an item, for example, a draft or an existing appointment.</span></span>

<span data-ttu-id="2ef61-111">在此演练结束时，您将具有一个在新建邮件时运行的外接程序。</span><span class="sxs-lookup"><span data-stu-id="2ef61-111">By the end of this walkthrough, you'll have an add-in that runs whenever a new message is created.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="2ef61-112">此功能仅在具有 Microsoft [](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) 365 订阅的 Outlook 网页版中受预览支持。</span><span class="sxs-lookup"><span data-stu-id="2ef61-112">This feature is only supported for [preview](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) in Outlook on the web with a Microsoft 365 subscription.</span></span> <span data-ttu-id="2ef61-113">请参阅 [本文中](#how-to-preview-the-event-based-activation-feature) 如何预览基于事件的激活功能，了解更多详细信息。</span><span class="sxs-lookup"><span data-stu-id="2ef61-113">See [How to preview the event-based activation feature](#how-to-preview-the-event-based-activation-feature) in this article for more details.</span></span>
>
> <span data-ttu-id="2ef61-114">由于预览功能可能会随时更改，恕不另行通知，因此不应在生产外接程序中使用。</span><span class="sxs-lookup"><span data-stu-id="2ef61-114">Because preview features are subject to change without notice, they shouldn't be used in production add-ins.</span></span>

## <a name="how-to-preview-the-event-based-activation-feature"></a><span data-ttu-id="2ef61-115">如何预览基于事件的激活功能</span><span class="sxs-lookup"><span data-stu-id="2ef61-115">How to preview the event-based activation feature</span></span>

<span data-ttu-id="2ef61-116">我们邀请你试用基于事件的激活功能！</span><span class="sxs-lookup"><span data-stu-id="2ef61-116">We invite you to try out the event-based activation feature!</span></span> <span data-ttu-id="2ef61-117">通过 GitHub 提供反馈，让我们了解你的方案以及如何改进 (请参阅此页面末尾的"反馈"部分) 。 </span><span class="sxs-lookup"><span data-stu-id="2ef61-117">Let us know your scenarios and how we can improve by giving us feedback through GitHub (see the **Feedback** section at the end of this page).</span></span>

<span data-ttu-id="2ef61-118">预览此功能：</span><span class="sxs-lookup"><span data-stu-id="2ef61-118">To preview this feature:</span></span>

- <span data-ttu-id="2ef61-119">引用 **CDN** 版本上的 beta https://appsforoffice.microsoft.com/lib/beta/hosted/office.js) (。</span><span class="sxs-lookup"><span data-stu-id="2ef61-119">Reference the **beta** library on the CDN (https://appsforoffice.microsoft.com/lib/beta/hosted/office.js).</span></span> <span data-ttu-id="2ef61-120">TypeScript [编译和](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts) 编译的类型IntelliSense CDN 和 [DefinitelyTyped 找到](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts)。</span><span class="sxs-lookup"><span data-stu-id="2ef61-120">The [type definition file](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts) for TypeScript compilation and IntelliSense is found at the CDN and [DefinitelyTyped](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts).</span></span> <span data-ttu-id="2ef61-121">可以安装这些类型 `npm install --save-dev @types/office-js-preview` 。</span><span class="sxs-lookup"><span data-stu-id="2ef61-121">You can install these types with `npm install --save-dev @types/office-js-preview`.</span></span>
- <span data-ttu-id="2ef61-122">[在 Microsoft 365 租户](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)上配置定向发布。</span><span class="sxs-lookup"><span data-stu-id="2ef61-122">[Configure targeted release on your Microsoft 365 tenant](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center).</span></span>

## <a name="set-up-your-environment"></a><span data-ttu-id="2ef61-123">设置环境</span><span class="sxs-lookup"><span data-stu-id="2ef61-123">Set up your environment</span></span>

<span data-ttu-id="2ef61-124">使用 [适用于 Office](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) 加载项的 Yeoman 生成器完成创建外接程序项目的 Outlook 快速入门。</span><span class="sxs-lookup"><span data-stu-id="2ef61-124">Complete the [Outlook quick start](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) which creates an add-in project with the Yeoman generator for Office Add-ins.</span></span>

## <a name="configure-the-manifest"></a><span data-ttu-id="2ef61-125">配置清单</span><span class="sxs-lookup"><span data-stu-id="2ef61-125">Configure the manifest</span></span>

<span data-ttu-id="2ef61-126">若要启用加载项的基于事件的激活，必须在清单中配置 [Runtimes](../reference/manifest/runtimes.md) 元素和 [LaunchEvent](../reference/manifest/extensionpoint.md#launchevent-preview) 扩展点。</span><span class="sxs-lookup"><span data-stu-id="2ef61-126">To enable event-based activation of your add-in, you must configure the [Runtimes](../reference/manifest/runtimes.md) element and [LaunchEvent](../reference/manifest/extensionpoint.md#launchevent-preview) extension point in the manifest.</span></span> <span data-ttu-id="2ef61-127">目前， `DesktopFormFactor` 是唯一受支持的外形类型。</span><span class="sxs-lookup"><span data-stu-id="2ef61-127">For now, `DesktopFormFactor` is the only supported form factor.</span></span>

1. <span data-ttu-id="2ef61-128">在代码编辑器中，打开快速启动项目。</span><span class="sxs-lookup"><span data-stu-id="2ef61-128">In your code editor, open the quick start project.</span></span>

1. <span data-ttu-id="2ef61-129">打开 **manifest.xml** 根目录下的文件。</span><span class="sxs-lookup"><span data-stu-id="2ef61-129">Open the **manifest.xml** file located at the root of your project.</span></span>

1. <span data-ttu-id="2ef61-130">选择整个 `<VersionOverrides>` 节点 (包括打开和关闭) 并将其替换为以下 XML。</span><span class="sxs-lookup"><span data-stu-id="2ef61-130">Select the entire `<VersionOverrides>` node (including open and close tags) and replace it with the following XML.</span></span>

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
        <bt:Url id="JSRuntime.Url" DefaultValue="https://localhost:3000/runtime.js" />
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

<span data-ttu-id="2ef61-131">Windows 上的 Outlook 使用 JavaScript 文件，而 Web 上的 Outlook 使用引用同一 JavaScript 文件的 HTML 文件。</span><span class="sxs-lookup"><span data-stu-id="2ef61-131">Outlook on Windows uses a JavaScript file, while Outlook on the web uses an HTML file that references the same JavaScript file.</span></span> <span data-ttu-id="2ef61-132">您必须在清单中提供对这两个文件的引用，因为 Outlook 平台最终确定是使用基于 Outlook 客户端的 HTML 还是 JavaScript。</span><span class="sxs-lookup"><span data-stu-id="2ef61-132">You must provide references to both these files in the manifest as the Outlook platform ultimately determines whether to use HTML or JavaScript based on the Outlook client.</span></span> <span data-ttu-id="2ef61-133">因此，若要配置事件处理，请提供 HTML 在元素中的位置，然后在其子元素中提供 HTML 内附或引用 `Runtime` `Override` 的 JavaScript 文件的位置。</span><span class="sxs-lookup"><span data-stu-id="2ef61-133">As such, to configure event handling, provide the location of the HTML in the `Runtime` element, then in its `Override` child element provide the location of the JavaScript file inlined or referenced by the HTML.</span></span>

> [!TIP]
> <span data-ttu-id="2ef61-134">若要了解有关 Outlook 外接程序清单的更多信息，请参阅 [Outlook 外接程序清单](manifests.md)。</span><span class="sxs-lookup"><span data-stu-id="2ef61-134">To learn more about manifests for Outlook add-ins, see [Outlook add-in manifests](manifests.md).</span></span>

## <a name="implement-event-handling"></a><span data-ttu-id="2ef61-135">实现事件处理</span><span class="sxs-lookup"><span data-stu-id="2ef61-135">Implement event handling</span></span>

<span data-ttu-id="2ef61-136">您必须对选定的事件实现处理。</span><span class="sxs-lookup"><span data-stu-id="2ef61-136">You have to implement handling for your selected events.</span></span>

<span data-ttu-id="2ef61-137">在此方案中，将添加用于撰写新项的处理。</span><span class="sxs-lookup"><span data-stu-id="2ef61-137">In this scenario, you'll add handling for composing new items.</span></span>

1. <span data-ttu-id="2ef61-138">从同一快速启动项目中，在代码编辑器中打开commands.js **./src/commands/commands.js** 文件。</span><span class="sxs-lookup"><span data-stu-id="2ef61-138">From the same quick start project, open the file **./src/commands/commands.js** in your code editor.</span></span>

1. <span data-ttu-id="2ef61-139">在函数 `action` 后插入以下 JavaScript 函数。</span><span class="sxs-lookup"><span data-stu-id="2ef61-139">After the `action` function, insert the following JavaScript functions.</span></span>

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

1. <span data-ttu-id="2ef61-140">在文件末尾，添加以下语句。</span><span class="sxs-lookup"><span data-stu-id="2ef61-140">At the end of the file, add the following statements.</span></span>

    ```js
    g.onMessageComposeHandler = onMessageComposeHandler;
    g.onAppointmentComposeHandler = onAppointmentComposeHandler;
    ```

## <a name="try-it-out"></a><span data-ttu-id="2ef61-141">试用</span><span class="sxs-lookup"><span data-stu-id="2ef61-141">Try it out</span></span>

1. <span data-ttu-id="2ef61-142">在项目的根目录中运行以下命令。</span><span class="sxs-lookup"><span data-stu-id="2ef61-142">Run the following command in the root directory of your project.</span></span> <span data-ttu-id="2ef61-143">运行此命令时，本地 Web 服务器将启动（如果尚未运行）。</span><span class="sxs-lookup"><span data-stu-id="2ef61-143">When you run this command, the local web server will start (if it's not already running).</span></span>

    ```command&nbsp;line
    npm run dev-server
    ```

1. <span data-ttu-id="2ef61-144">按照[旁加载 Outlook 加载项以供测试](sideload-outlook-add-ins-for-testing.md)中的说明操作，旁加载加载项。</span><span class="sxs-lookup"><span data-stu-id="2ef61-144">Follow the instructions in [Sideload Outlook add-ins for testing](sideload-outlook-add-ins-for-testing.md) to sideload the add-in in Outlook.</span></span>

1. <span data-ttu-id="2ef61-145">在 Outlook 网页版中，创建新邮件。</span><span class="sxs-lookup"><span data-stu-id="2ef61-145">In Outlook on the web, create a new message.</span></span>

    ![Outlook 网页中的邮件窗口的屏幕截图，主题在撰写时设置。](../images/outlook-web-autolaunch.png)

## <a name="event-based-activation-behavior-and-limitations"></a><span data-ttu-id="2ef61-147">基于事件的激活行为和限制</span><span class="sxs-lookup"><span data-stu-id="2ef61-147">Event-based activation behavior and limitations</span></span>

<span data-ttu-id="2ef61-148">基于事件激活的外接程序设计为短时间运行，最多 330 秒。</span><span class="sxs-lookup"><span data-stu-id="2ef61-148">Add-ins that activate based on events are designed to be short-running, up to 330 seconds only.</span></span> <span data-ttu-id="2ef61-149">我们建议你让加载项调用该方法，以表明它 `event.completed` 已完成处理启动事件。</span><span class="sxs-lookup"><span data-stu-id="2ef61-149">We recommend you have your add-in call the `event.completed` method to signal it has completed processing the launch event.</span></span> <span data-ttu-id="2ef61-150">当用户关闭撰写窗口时，外接程序也会结束。</span><span class="sxs-lookup"><span data-stu-id="2ef61-150">The add-in also ends when the user closes the compose window.</span></span>

<span data-ttu-id="2ef61-151">如果用户有多个订阅同一事件的加载项，Outlook 平台将按特定顺序启动外接程序。</span><span class="sxs-lookup"><span data-stu-id="2ef61-151">If the user has multiple add-ins that subscribed to the same event, the Outlook platform launches the add-ins in no particular order.</span></span> <span data-ttu-id="2ef61-152">目前，只能主动运行五个基于事件的加载项。</span><span class="sxs-lookup"><span data-stu-id="2ef61-152">Currently, only five event-based add-ins can be actively running.</span></span> <span data-ttu-id="2ef61-153">任何其他加载项将推送到队列，然后随着之前处于活动状态的加载项完成或停用而运行。</span><span class="sxs-lookup"><span data-stu-id="2ef61-153">Any additional add-ins are pushed to a queue then run as previously active add-ins are completed or deactivated.</span></span>

<span data-ttu-id="2ef61-154">用户可以切换或导航离开外接程序开始运行的当前邮件项目。</span><span class="sxs-lookup"><span data-stu-id="2ef61-154">The user can switch or navigate away from the current mail item where the add-in started running.</span></span> <span data-ttu-id="2ef61-155">启动的加载项将在后台完成其操作。</span><span class="sxs-lookup"><span data-stu-id="2ef61-155">The add-in that was launched will finish its operation in the background.</span></span>

<span data-ttu-id="2ef61-156">某些Office.js更改或更改 UI 的 API 不允许来自基于事件的加载项。下面是阻止的 API。</span><span class="sxs-lookup"><span data-stu-id="2ef61-156">Some Office.js APIs that change or alter the UI are not allowed from event-based add-ins. The following are the blocked APIs.</span></span>

- <span data-ttu-id="2ef61-157">Under `Office.context.mailbox` ：</span><span class="sxs-lookup"><span data-stu-id="2ef61-157">Under `Office.context.mailbox`:</span></span>
  - `displayAppointmentForm`
  - `displayMessageForm`
  - `displayNewAppointmentForm`
  - `displayNewMessageForm`
- <span data-ttu-id="2ef61-158">Under `Office.context.ui` ：</span><span class="sxs-lookup"><span data-stu-id="2ef61-158">Under `Office.context.ui`:</span></span>
  - `displayDialogAsync`
  - `messageParent`
- <span data-ttu-id="2ef61-159">Under `Office.context.auth` ：</span><span class="sxs-lookup"><span data-stu-id="2ef61-159">Under `Office.context.auth`:</span></span>
  - `getAccessToken`
  - `getAccessTokenAsync`

## <a name="see-also"></a><span data-ttu-id="2ef61-160">另请参阅</span><span class="sxs-lookup"><span data-stu-id="2ef61-160">See also</span></span>

[<span data-ttu-id="2ef61-161">Outlook 加载项清单</span><span class="sxs-lookup"><span data-stu-id="2ef61-161">Outlook add-in manifests</span></span>](manifests.md)
