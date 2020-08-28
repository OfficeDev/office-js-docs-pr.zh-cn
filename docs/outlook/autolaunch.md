---
title: '配置 Outlook 外接程序以进行基于事件的激活 (预览) '
description: 了解如何配置 Outlook 外接程序以进行基于事件的激活。
ms.topic: article
ms.date: 08/24/2020
localization_priority: Normal
ms.openlocfilehash: 0131cafa8315315d63b6319ecad4fd41b1168073
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/28/2020
ms.locfileid: "47293924"
---
# <a name="configure-your-outlook-add-in-for-event-based-activation-preview"></a><span data-ttu-id="d2ca6-103">配置 Outlook 外接程序以进行基于事件的激活 (预览) </span><span class="sxs-lookup"><span data-stu-id="d2ca6-103">Configure your Outlook add-in for event-based activation (preview)</span></span>

<span data-ttu-id="d2ca6-104">如果没有基于事件的激活功能，用户必须显式启动外接程序以完成其任务。</span><span class="sxs-lookup"><span data-stu-id="d2ca6-104">Without the event-based activation feature, a user has to explicitly launch an add-in to complete their tasks.</span></span> <span data-ttu-id="d2ca6-105">此功能使加载项能够根据特定事件（尤其是适用于每个项目的操作）运行任务。</span><span class="sxs-lookup"><span data-stu-id="d2ca6-105">This feature enables your add-in to run tasks based on certain events, particularly for operations that apply to every item.</span></span> <span data-ttu-id="d2ca6-106">您还可以与任务窗格和无 UI 功能集成。</span><span class="sxs-lookup"><span data-stu-id="d2ca6-106">You can also integrate with the task pane and UI-less functionality.</span></span> <span data-ttu-id="d2ca6-107">目前，支持以下事件。</span><span class="sxs-lookup"><span data-stu-id="d2ca6-107">At present, the following events are supported.</span></span>

- <span data-ttu-id="d2ca6-108">`OnNewMessageCompose`：撰写新邮件时 (包括答复、全部答复和转发) </span><span class="sxs-lookup"><span data-stu-id="d2ca6-108">`OnNewMessageCompose`: On composing a new message (includes reply, reply all, and forward)</span></span>
- <span data-ttu-id="d2ca6-109">`OnNewAppointmentOrganizer`：创建新约会时</span><span class="sxs-lookup"><span data-stu-id="d2ca6-109">`OnNewAppointmentOrganizer`: On creating a new appointment</span></span>

  > [!IMPORTANT]
  > <span data-ttu-id="d2ca6-110">此 **功能不会激活编辑** 项目（例如，草稿或现有约会）。</span><span class="sxs-lookup"><span data-stu-id="d2ca6-110">This feature does **not** activate on editing an item, for example, a draft or an existing appointment.</span></span>

<span data-ttu-id="d2ca6-111">本演练结束时，您将拥有一个在创建新邮件时运行的外接程序。</span><span class="sxs-lookup"><span data-stu-id="d2ca6-111">By the end of this walkthrough, you'll have an add-in that runs whenever a new message is created.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="d2ca6-112">只有使用 Microsoft 365 订阅的 Outlook 网页版中的 [预览](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) 才支持此功能。</span><span class="sxs-lookup"><span data-stu-id="d2ca6-112">This feature is only supported for [preview](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) in Outlook on the web with a Microsoft 365 subscription.</span></span> <span data-ttu-id="d2ca6-113">有关更多详细信息，请参阅 [如何预览本文中基于事件的激活功能](#how-to-preview-the-event-based-activation-feature) 。</span><span class="sxs-lookup"><span data-stu-id="d2ca6-113">See [How to preview the event-based activation feature](#how-to-preview-the-event-based-activation-feature) in this article for more details.</span></span>
>
> <span data-ttu-id="d2ca6-114">由于预览功能可能会发生更改，恕不另行通知，它们不应在生产外接程序中使用。</span><span class="sxs-lookup"><span data-stu-id="d2ca6-114">Because preview features are subject to change without notice, they shouldn't be used in production add-ins.</span></span>

## <a name="how-to-preview-the-event-based-activation-feature"></a><span data-ttu-id="d2ca6-115">如何预览基于事件的激活功能</span><span class="sxs-lookup"><span data-stu-id="d2ca6-115">How to preview the event-based activation feature</span></span>

<span data-ttu-id="d2ca6-116">我们邀请你试用基于事件的激活功能！</span><span class="sxs-lookup"><span data-stu-id="d2ca6-116">We invite you to try out the event-based activation feature!</span></span> <span data-ttu-id="d2ca6-117">请通过 GitHub 向我们提供反馈，告知我们你的方案以及我们如何改进， (请参阅本页结尾处的 **反馈** 部分) 。</span><span class="sxs-lookup"><span data-stu-id="d2ca6-117">Let us know your scenarios and how we can improve by giving us feedback through GitHub (see the **Feedback** section at the end of this page).</span></span>

<span data-ttu-id="d2ca6-118">若要预览此功能：</span><span class="sxs-lookup"><span data-stu-id="d2ca6-118">To preview this feature:</span></span>

- <span data-ttu-id="d2ca6-119">参考 CDN (上的 **beta** 库 https://appsforoffice.microsoft.com/lib/beta/hosted/office.js) 。</span><span class="sxs-lookup"><span data-stu-id="d2ca6-119">Reference the **beta** library on the CDN (https://appsforoffice.microsoft.com/lib/beta/hosted/office.js).</span></span> <span data-ttu-id="d2ca6-120">在 CDN 和[jquery.typescript.definitelytyped](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts)中找到 TypeScript 编译和智能感知的[类型定义文件](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts)。</span><span class="sxs-lookup"><span data-stu-id="d2ca6-120">The [type definition file](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts) for TypeScript compilation and IntelliSense is found at the CDN and [DefinitelyTyped](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts).</span></span> <span data-ttu-id="d2ca6-121">您可以使用安装这些类型 `npm install --save-dev @types/office-js-preview` 。</span><span class="sxs-lookup"><span data-stu-id="d2ca6-121">You can install these types with `npm install --save-dev @types/office-js-preview`.</span></span>
- <span data-ttu-id="d2ca6-122">通过完成并提交 [此请求表单](https://aka.ms/OWAPreview)，请求使用 Microsoft 365 帐户对在 web 上的 Outlook 的预览位进行访问。</span><span class="sxs-lookup"><span data-stu-id="d2ca6-122">Request access to preview bits for Outlook on the web using your Microsoft 365 account by completing and submitting [this request form](https://aka.ms/OWAPreview).</span></span> <span data-ttu-id="d2ca6-123">我们将在你的租户准备就绪时通知你。</span><span class="sxs-lookup"><span data-stu-id="d2ca6-123">We'll let you know when your tenant is ready.</span></span>

## <a name="set-up-your-environment"></a><span data-ttu-id="d2ca6-124">设置环境</span><span class="sxs-lookup"><span data-stu-id="d2ca6-124">Set up your environment</span></span>

<span data-ttu-id="d2ca6-125">完成 [Outlook 快速入门](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) ，它将使用 Office 外接程序的 Yeoman 生成器创建外接程序项目。</span><span class="sxs-lookup"><span data-stu-id="d2ca6-125">Complete the [Outlook quick start](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) which creates an add-in project with the Yeoman generator for Office Add-ins.</span></span>

## <a name="configure-the-manifest"></a><span data-ttu-id="d2ca6-126">配置清单</span><span class="sxs-lookup"><span data-stu-id="d2ca6-126">Configure the manifest</span></span>

<span data-ttu-id="d2ca6-127">若要启用您的外接程序的基于事件的激活，必须在清单中配置 [运行时](../reference/manifest/runtimes.md) 元素和 [LaunchEvent](../reference/manifest/extensionpoint.md#launchevent-preview) 扩展点。</span><span class="sxs-lookup"><span data-stu-id="d2ca6-127">To enable event-based activation of your add-in, you must configure the [Runtimes](../reference/manifest/runtimes.md) element and [LaunchEvent](../reference/manifest/extensionpoint.md#launchevent-preview) extension point in the manifest.</span></span> <span data-ttu-id="d2ca6-128">目前， `DesktopFormFactor` 是唯一受支持的板型。</span><span class="sxs-lookup"><span data-stu-id="d2ca6-128">For now, `DesktopFormFactor` is the only supported form factor.</span></span>

1. <span data-ttu-id="d2ca6-129">在代码编辑器中，打开 "快速启动" 项目。</span><span class="sxs-lookup"><span data-stu-id="d2ca6-129">In your code editor, open the quick start project.</span></span>

1. <span data-ttu-id="d2ca6-130">打开位于项目根目录中的 **manifest.xml** 文件。</span><span class="sxs-lookup"><span data-stu-id="d2ca6-130">Open the **manifest.xml** file located at the root of your project.</span></span>

1. <span data-ttu-id="d2ca6-131">选择整个 `<VersionOverrides>` 节点 (包括 "打开" 和 "关闭" 标记) 并将其替换为以下 XML。</span><span class="sxs-lookup"><span data-stu-id="d2ca6-131">Select the entire `<VersionOverrides>` node (including open and close tags) and replace it with the following XML.</span></span>

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

<span data-ttu-id="d2ca6-132">Windows 上的 outlook 使用 JavaScript 文件，而 web 上的 Outlook 使用引用相同 JavaScript 文件的 HTML 文件。</span><span class="sxs-lookup"><span data-stu-id="d2ca6-132">Outlook on Windows uses a JavaScript file, while Outlook on the web uses an HTML file that references the same JavaScript file.</span></span> <span data-ttu-id="d2ca6-133">由于 Outlook 平台最终决定是使用基于 Outlook 客户端的 HTML 还是 JavaScript，因此您必须在清单中提供对这些文件的引用。</span><span class="sxs-lookup"><span data-stu-id="d2ca6-133">You must provide references to both these files in the manifest as the Outlook platform ultimately determines whether to use HTML or JavaScript based on the Outlook client.</span></span> <span data-ttu-id="d2ca6-134">因此，若要配置事件处理，请在元素中提供 HTML 的位置 `Runtime` ，然后在其子 `Override` 元素中提供由 html 内联或引用的 JavaScript 文件的位置。</span><span class="sxs-lookup"><span data-stu-id="d2ca6-134">As such, to configure event handling, provide the location of the HTML in the `Runtime` element, then in its `Override` child element provide the location of the JavaScript file inlined or referenced by the HTML.</span></span>

> [!TIP]
> <span data-ttu-id="d2ca6-135">若要了解有关 Outlook 外接程序的清单的详细信息，请参阅 [outlook 外接程序清单](manifests.md)。</span><span class="sxs-lookup"><span data-stu-id="d2ca6-135">To learn more about manifests for Outlook add-ins, see [Outlook add-in manifests](manifests.md).</span></span>

## <a name="implement-event-handling"></a><span data-ttu-id="d2ca6-136">实现事件处理</span><span class="sxs-lookup"><span data-stu-id="d2ca6-136">Implement event handling</span></span>

<span data-ttu-id="d2ca6-137">您必须为选定的事件实现处理。</span><span class="sxs-lookup"><span data-stu-id="d2ca6-137">You have to implement handling for your selected events.</span></span>

<span data-ttu-id="d2ca6-138">在这种情况下，您将添加用于撰写新项目的处理。</span><span class="sxs-lookup"><span data-stu-id="d2ca6-138">In this scenario, you'll add handling for composing new items.</span></span>

1. <span data-ttu-id="d2ca6-139">在同一 "快速启动" 项目中，在代码编辑器中打开 **/src/commands/commands.js** 。</span><span class="sxs-lookup"><span data-stu-id="d2ca6-139">From the same quick start project, open the file **./src/commands/commands.js** in your code editor.</span></span>

1. <span data-ttu-id="d2ca6-140">在 `action` 函数后面，插入以下 JavaScript 函数。</span><span class="sxs-lookup"><span data-stu-id="d2ca6-140">After the `action` function, insert the following JavaScript functions.</span></span>

    ```js
    function onMessageComposeHandler(event) {
      setSubject();
      event.completed();
    }
    function onAppointmentComposeHandler(event) {
      setSubject();
      event.completed();
    }
    function setSubject() {
      Office.context.mailbox.item.subject.setAsync("Set by an event-based add-in!");
    }
    ```

1. <span data-ttu-id="d2ca6-141">在文件末尾，添加以下语句。</span><span class="sxs-lookup"><span data-stu-id="d2ca6-141">At the end of the file, add the following statements.</span></span>

    ```js
    g.onMessageComposeHandler = onMessageComposeHandler;
    g.onAppointmentComposeHandler = onAppointmentComposeHandler;
    ```

## <a name="try-it-out"></a><span data-ttu-id="d2ca6-142">试用</span><span class="sxs-lookup"><span data-stu-id="d2ca6-142">Try it out</span></span>

1. <span data-ttu-id="d2ca6-143">在项目的根目录中运行以下命令。</span><span class="sxs-lookup"><span data-stu-id="d2ca6-143">Run the following command in the root directory of your project.</span></span> <span data-ttu-id="d2ca6-144">运行此命令时，本地 Web 服务器将启动（如果尚未运行）。</span><span class="sxs-lookup"><span data-stu-id="d2ca6-144">When you run this command, the local web server will start (if it's not already running).</span></span>

    ```command&nbsp;line
    npm start
    ```

    > [!IMPORTANT]
    > <span data-ttu-id="d2ca6-145">如果看到 "旁加载不受支持" 错误，则可以忽略它并继续。</span><span class="sxs-lookup"><span data-stu-id="d2ca6-145">If you see a "Sideload is not supported" error, you can ignore it and proceed.</span></span>

1. <span data-ttu-id="d2ca6-146">按照[旁加载 Outlook 加载项以供测试](sideload-outlook-add-ins-for-testing.md)中的说明操作，旁加载加载项。</span><span class="sxs-lookup"><span data-stu-id="d2ca6-146">Follow the instructions in [Sideload Outlook add-ins for testing](sideload-outlook-add-ins-for-testing.md) to sideload the add-in in Outlook.</span></span>

1. <span data-ttu-id="d2ca6-147">在 Outlook 网页版中，创建新邮件。</span><span class="sxs-lookup"><span data-stu-id="d2ca6-147">In Outlook on the web, create a new message.</span></span>

    ![Outlook 网页版中邮件窗口的屏幕截图，其中的主题设置为撰写。](../images/outlook-web-autolaunch.png)

## <a name="event-based-activation-behavior-and-limitations"></a><span data-ttu-id="d2ca6-149">基于事件的激活行为和限制</span><span class="sxs-lookup"><span data-stu-id="d2ca6-149">Event-based activation behavior and limitations</span></span>

<span data-ttu-id="d2ca6-150">根据事件激活的加载项设计为运行时间较短，最长为330秒。</span><span class="sxs-lookup"><span data-stu-id="d2ca6-150">Add-ins that activate based on events are designed to be short-running, up to 330 seconds only.</span></span> <span data-ttu-id="d2ca6-151">我们建议您让您的外接程序调用 `event.completed` 方法，以通知其已完成启动事件的处理。</span><span class="sxs-lookup"><span data-stu-id="d2ca6-151">We recommend you have your add-in call the `event.completed` method to signal it has completed processing the launch event.</span></span> <span data-ttu-id="d2ca6-152">当用户关闭撰写窗口时，外接端也会结束。</span><span class="sxs-lookup"><span data-stu-id="d2ca6-152">The add-in also ends when the user closes the compose window.</span></span>

<span data-ttu-id="d2ca6-153">如果用户具有多个订阅同一事件的加载项，则 Outlook 平台将以无特定的顺序启动外接程序。</span><span class="sxs-lookup"><span data-stu-id="d2ca6-153">If the user has multiple add-ins that subscribed to the same event, the Outlook platform launches the add-ins in no particular order.</span></span> <span data-ttu-id="d2ca6-154">目前，只有五个基于事件的外接程序可以处于活动状态。</span><span class="sxs-lookup"><span data-stu-id="d2ca6-154">Currently, only five event-based add-ins can be actively running.</span></span> <span data-ttu-id="d2ca6-155">任何其他外接程序将被推送到队列中，然后运行之前的活动外接程序已完成或停用。</span><span class="sxs-lookup"><span data-stu-id="d2ca6-155">Any additional add-ins are pushed to a queue then run as previously active add-ins are completed or deactivated.</span></span>

<span data-ttu-id="d2ca6-156">用户可以从加载项开始运行的当前邮件项目中进行切换或导航。</span><span class="sxs-lookup"><span data-stu-id="d2ca6-156">The user can switch or navigate away from the current mail item where the add-in started running.</span></span> <span data-ttu-id="d2ca6-157">启动的外接程序将在后台完成其操作。</span><span class="sxs-lookup"><span data-stu-id="d2ca6-157">The add-in that was launched will finish its operation in the background.</span></span>

<span data-ttu-id="d2ca6-158">基于事件的外接程序不允许更改或更改 UI 的一些 Office.js Api。以下是阻止的 Api。</span><span class="sxs-lookup"><span data-stu-id="d2ca6-158">Some Office.js APIs that change or alter the UI are not allowed from event-based add-ins. The following are the blocked APIs.</span></span>

- <span data-ttu-id="d2ca6-159">在 `Office.context.mailbox` ：</span><span class="sxs-lookup"><span data-stu-id="d2ca6-159">Under `Office.context.mailbox`:</span></span>
  - `displayAppointmentForm`
  - `displayMessageForm`
  - `displayNewAppointmentForm`
  - `displayNewMessageForm`
- <span data-ttu-id="d2ca6-160">在 `Office.context.ui` ：</span><span class="sxs-lookup"><span data-stu-id="d2ca6-160">Under `Office.context.ui`:</span></span>
  - `displayDialogAsync`
  - `messageParent`
- <span data-ttu-id="d2ca6-161">在 `Office.context.auth` ：</span><span class="sxs-lookup"><span data-stu-id="d2ca6-161">Under `Office.context.auth`:</span></span>
  - `getAccessTokenAsync`

## <a name="see-also"></a><span data-ttu-id="d2ca6-162">另请参阅</span><span class="sxs-lookup"><span data-stu-id="d2ca6-162">See also</span></span>

[<span data-ttu-id="d2ca6-163">Outlook 加载项清单</span><span class="sxs-lookup"><span data-stu-id="d2ca6-163">Outlook add-in manifests</span></span>](manifests.md)
