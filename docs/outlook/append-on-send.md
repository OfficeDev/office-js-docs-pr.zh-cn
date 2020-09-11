---
title: '在 Outlook 加载项中实现向外接程序发送 (预览) '
description: 了解如何在 Outlook 外接程序中实现 "发送时发送" 功能。
ms.topic: article
ms.date: 09/09/2020
localization_priority: Normal
ms.openlocfilehash: 2199f837351c1030e6f6d0d23db7bf81e498d433
ms.sourcegitcommit: 83f9a2fdff81ca421cd23feea103b9b60895cab4
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/11/2020
ms.locfileid: "47430931"
---
# <a name="implement-append-on-send-in-your-outlook-add-in-preview"></a><span data-ttu-id="d02bc-103">在 Outlook 加载项中实现向外接程序发送 (预览) </span><span class="sxs-lookup"><span data-stu-id="d02bc-103">Implement append-on-send in your Outlook add-in (preview)</span></span>

<span data-ttu-id="d02bc-104">本演练结束时，您将拥有一个可在发送邮件时插入免责声明的 Outlook 外接程序。</span><span class="sxs-lookup"><span data-stu-id="d02bc-104">By the end of this walkthrough, you'll have an Outlook add-in that can insert a disclaimer when a message is sent.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="d02bc-105">此功能目前仅支持在 Outlook 网页版和使用 Microsoft 365 订阅的 Windows 中进行 [预览](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) 。</span><span class="sxs-lookup"><span data-stu-id="d02bc-105">This feature is currently supported for [preview](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) in Outlook on the web and Windows with a Microsoft 365 subscription.</span></span> <span data-ttu-id="d02bc-106">有关更多详细信息，请参阅 [如何预览本文中的追加发送功能](#how-to-preview-the-append-on-send-feature) 。</span><span class="sxs-lookup"><span data-stu-id="d02bc-106">See [How to preview the append-on-send feature](#how-to-preview-the-append-on-send-feature) in this article for more details.</span></span>
>
> <span data-ttu-id="d02bc-107">由于预览功能可能会发生更改，恕不另行通知，它们不应在生产外接程序中使用。</span><span class="sxs-lookup"><span data-stu-id="d02bc-107">Because preview features are subject to change without notice, they shouldn't be used in production add-ins.</span></span>

## <a name="how-to-preview-the-append-on-send-feature"></a><span data-ttu-id="d02bc-108">如何预览追加发送功能</span><span class="sxs-lookup"><span data-stu-id="d02bc-108">How to preview the append-on-send feature</span></span>

<span data-ttu-id="d02bc-109">我们邀请你试用 "发送时追加" 功能！</span><span class="sxs-lookup"><span data-stu-id="d02bc-109">We invite you to try out the append-on-send feature!</span></span> <span data-ttu-id="d02bc-110">请通过 GitHub 向我们提供反馈，告知我们你的方案以及我们如何改进， (请参阅本页结尾处的 **反馈** 部分) 。</span><span class="sxs-lookup"><span data-stu-id="d02bc-110">Let us know your scenarios and how we can improve by giving us feedback through GitHub (see the **Feedback** section at the end of this page).</span></span>

<span data-ttu-id="d02bc-111">若要预览此功能：</span><span class="sxs-lookup"><span data-stu-id="d02bc-111">To preview this feature:</span></span>

- <span data-ttu-id="d02bc-112">参考 CDN (上的 **beta** 库 https://appsforoffice.microsoft.com/lib/beta/hosted/office.js) 。</span><span class="sxs-lookup"><span data-stu-id="d02bc-112">Reference the **beta** library on the CDN (https://appsforoffice.microsoft.com/lib/beta/hosted/office.js).</span></span> <span data-ttu-id="d02bc-113">在 CDN 和[jquery.typescript.definitelytyped](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts)中找到 TypeScript 编译和智能感知的[类型定义文件](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts)。</span><span class="sxs-lookup"><span data-stu-id="d02bc-113">The [type definition file](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts) for TypeScript compilation and IntelliSense is found at the CDN and [DefinitelyTyped](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts).</span></span> <span data-ttu-id="d02bc-114">您可以使用安装这些类型 `npm install --save-dev @types/office-js-preview` 。</span><span class="sxs-lookup"><span data-stu-id="d02bc-114">You can install these types with `npm install --save-dev @types/office-js-preview`.</span></span>
- <span data-ttu-id="d02bc-115">对于 Windows，你可能需要加入 [Office 预览体验成员计划](https://insider.office.com) ，以访问更多最近的 office 版本。</span><span class="sxs-lookup"><span data-stu-id="d02bc-115">For Windows, you may need to join the [Office Insider program](https://insider.office.com) to access more recent Office builds.</span></span>
- <span data-ttu-id="d02bc-116">对于 web 上的 Outlook， [在 Microsoft 365 租户上配置目标版本](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)。</span><span class="sxs-lookup"><span data-stu-id="d02bc-116">For Outlook on the web, [configure targeted release on your Microsoft 365 tenant](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center).</span></span>

## <a name="set-up-your-environment"></a><span data-ttu-id="d02bc-117">设置环境</span><span class="sxs-lookup"><span data-stu-id="d02bc-117">Set up your environment</span></span>

<span data-ttu-id="d02bc-118">完成 [Outlook 快速入门](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) ，它将使用 Office 外接程序的 Yeoman 生成器创建外接程序项目。</span><span class="sxs-lookup"><span data-stu-id="d02bc-118">Complete the [Outlook quick start](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) which creates an add-in project with the Yeoman generator for Office Add-ins.</span></span>

## <a name="configure-the-manifest"></a><span data-ttu-id="d02bc-119">配置清单</span><span class="sxs-lookup"><span data-stu-id="d02bc-119">Configure the manifest</span></span>

<span data-ttu-id="d02bc-120">若要在您的外接程序中启用 "追加发送" 功能，必须 `AppendOnSend` 在 [ExtendedPermissions](../reference/manifest/extendedpermissions.md)集合中包含该权限。</span><span class="sxs-lookup"><span data-stu-id="d02bc-120">To enable the append-on-send feature in your add-in, you must include the `AppendOnSend` permission in the collection of [ExtendedPermissions](../reference/manifest/extendedpermissions.md).</span></span>

<span data-ttu-id="d02bc-121">对于此方案， `action` 您将运行函数，而不是在选择 " **执行操作** " 按钮时运行函数 `appendOnSend` 。</span><span class="sxs-lookup"><span data-stu-id="d02bc-121">For this scenario, instead of running the `action` function on choosing the **Perform an action** button, you'll be running the `appendOnSend` function.</span></span>

1. <span data-ttu-id="d02bc-122">在代码编辑器中，打开 "快速启动" 项目。</span><span class="sxs-lookup"><span data-stu-id="d02bc-122">In your code editor, open the quick start project.</span></span>

1. <span data-ttu-id="d02bc-123">打开位于项目根目录中的 **manifest.xml** 文件。</span><span class="sxs-lookup"><span data-stu-id="d02bc-123">Open the **manifest.xml** file located at the root of your project.</span></span>

1. <span data-ttu-id="d02bc-124">选择整个 `<VersionOverrides>` 节点 (包括 "打开" 和 "关闭" 标记) 并将其替换为以下 XML。</span><span class="sxs-lookup"><span data-stu-id="d02bc-124">Select the entire `<VersionOverrides>` node (including open and close tags) and replace it with the following XML.</span></span>

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
            <DesktopFormFactor>
              <FunctionFile resid="Commands.Url" />
              <ExtensionPoint xsi:type="MessageComposeCommandSurface">
                <OfficeTab id="TabDefault">
                  <Group id="msgComposeGroup">
                    <Label resid="GroupLabel" />
                    <Control xsi:type="Button" id="msgComposeOpenPaneButton">
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
                        <FunctionName>appendDisclaimerOnSend</FunctionName>
                      </Action>
                    </Control>
                  </Group>
                </OfficeTab>
              </ExtensionPoint>

              <!-- Configure AppointmentOrganizerCommandSurface extension point to support
              append on sending a new appointment. -->

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
        <ExtendedPermissions>
          <ExtendedPermission>AppendOnSend</ExtendedPermission>
        </ExtendedPermissions>
      </VersionOverrides>
    </VersionOverrides>
    ```

> [!TIP]
> <span data-ttu-id="d02bc-125">若要了解有关 Outlook 外接程序的清单的详细信息，请参阅 [outlook 外接程序清单](manifests.md)。</span><span class="sxs-lookup"><span data-stu-id="d02bc-125">To learn more about manifests for Outlook add-ins, see [Outlook add-in manifests](manifests.md).</span></span>

## <a name="implement-append-on-send-handling"></a><span data-ttu-id="d02bc-126">实现附加发送前处理</span><span class="sxs-lookup"><span data-stu-id="d02bc-126">Implement append-on-send handling</span></span>

<span data-ttu-id="d02bc-127">接下来，实现在 send 事件上追加。</span><span class="sxs-lookup"><span data-stu-id="d02bc-127">Next, implement appending on the send event.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="d02bc-128">如果您的外接程序还实现了[使用 `ItemSend` 中的发送事件处理](outlook-on-send-addins.md)，则在 `AppendOnSendAsync` 发送时处理程序中调用将返回错误，因为这种情况不受支持。</span><span class="sxs-lookup"><span data-stu-id="d02bc-128">If your add-in also implements [on-send event handling using `ItemSend`](outlook-on-send-addins.md), calling `AppendOnSendAsync` in the on-send handler returns an error as this scenario isn't supported.</span></span>

<span data-ttu-id="d02bc-129">在这种情况下，您将实现在用户发送时向项目追加免责声明。</span><span class="sxs-lookup"><span data-stu-id="d02bc-129">For this scenario, you'll implement appending a disclaimer to the item when the user sends.</span></span>

1. <span data-ttu-id="d02bc-130">在同一 "快速启动" 项目中，在代码编辑器中打开 **/src/commands/commands.js** 。</span><span class="sxs-lookup"><span data-stu-id="d02bc-130">From the same quick start project, open the file **./src/commands/commands.js** in your code editor.</span></span>

1. <span data-ttu-id="d02bc-131">在 `action` 函数后面，插入以下 JavaScript 函数。</span><span class="sxs-lookup"><span data-stu-id="d02bc-131">After the `action` function, insert the following JavaScript function.</span></span>

    ```js
    function appendDisclaimerOnSend(event) {
      var appendText =
        '<p style = "color:blue"> <i>This and subsequent emails on the same topic are for discussion and information purposes only. Only those matters set out in a fully executed agreement are legally binding. This email may contain confidential information and should not be shared with any third party without the prior written agreement of Contoso. If you are not the intended recipient, take no action and contact the sender immediately.<br><br>Contoso Limited (company number 01624297) is a company registered in England and Wales whose registered office is at Contoso Campus, Thames Valley Park, Reading RG6 1WG</i></p>';  
      /**
        *************************************************************
         Ideal Usage - Call the getBodyType API. Use the coercionType
         it returns as the parameter value below.
        *************************************************************
      */
      Office.context.mailbox.item.body.appendOnSendAsync(
        appendText,
        {
          coercionType: Office.CoercionType.Html
        },
        function(asyncResult) {
          console.log(asyncResult);
        }
      );

      event.completed();
    }
    ```

1. <span data-ttu-id="d02bc-132">在文件末尾，添加以下语句。</span><span class="sxs-lookup"><span data-stu-id="d02bc-132">At the end of the file, add the following statement.</span></span>

    ```js
    g.appendDisclaimerOnSend = appendDisclaimerOnSend;
    ```

## <a name="try-it-out"></a><span data-ttu-id="d02bc-133">试用</span><span class="sxs-lookup"><span data-stu-id="d02bc-133">Try it out</span></span>

1. <span data-ttu-id="d02bc-134">在项目的根目录中运行以下命令。</span><span class="sxs-lookup"><span data-stu-id="d02bc-134">Run the following command in the root directory of your project.</span></span> <span data-ttu-id="d02bc-135">运行此命令时，本地 web 服务器将启动（如果它尚未运行）。</span><span class="sxs-lookup"><span data-stu-id="d02bc-135">When you run this command, the local web server will start if it's not already running.</span></span>

    ```command&nbsp;line
    npm run dev-server
    ```

1. <span data-ttu-id="d02bc-136">按照 [旁加载 Outlook 外接程序](sideload-outlook-add-ins-for-testing.md)中的说明进行操作，以进行测试。</span><span class="sxs-lookup"><span data-stu-id="d02bc-136">Follow the instructions in [Sideload Outlook add-ins for testing](sideload-outlook-add-ins-for-testing.md).</span></span>

1. <span data-ttu-id="d02bc-137">创建新邮件，并将自己添加到 " **to** " 行。</span><span class="sxs-lookup"><span data-stu-id="d02bc-137">Create a new message, and add yourself to the **To** line.</span></span>

1. <span data-ttu-id="d02bc-138">从 "功能区" 或 "溢出" 菜单中，选择 " **执行操作**"。</span><span class="sxs-lookup"><span data-stu-id="d02bc-138">From the ribbon or overflow menu, choose **Perform an action**.</span></span>

1. <span data-ttu-id="d02bc-139">发送邮件，然后从 **"收件箱" 或 "** **已发送邮件** " 文件夹中打开它以查看追加的免责声明。</span><span class="sxs-lookup"><span data-stu-id="d02bc-139">Send the message, then open it from your **Inbox** or **Sent Items** folder to view the appended disclaimer.</span></span>

    ![在 Outlook 网页版上追加的包含免责声明的示例邮件的屏幕截图。](../images/outlook-web-append-disclaimer.png)

## <a name="see-also"></a><span data-ttu-id="d02bc-141">另请参阅</span><span class="sxs-lookup"><span data-stu-id="d02bc-141">See also</span></span>

[<span data-ttu-id="d02bc-142">Outlook 加载项清单</span><span class="sxs-lookup"><span data-stu-id="d02bc-142">Outlook add-in manifests</span></span>](manifests.md)
