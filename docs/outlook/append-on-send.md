---
title: 在 Outlook 外接程序中实现追加发送
description: 了解如何在 Outlook 外接程序中实现 "发送时发送" 功能。
ms.topic: article
ms.date: 10/14/2020
localization_priority: Normal
ms.openlocfilehash: 62234f580f6ff6be418f1c252510f234e297b0c6
ms.sourcegitcommit: 4e7c74ad67ea8bf6b47d65b2fde54a967090f65b
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/20/2020
ms.locfileid: "48626454"
---
# <a name="implement-append-on-send-in-your-outlook-add-in"></a><span data-ttu-id="7494a-103">在 Outlook 外接程序中实现追加发送</span><span class="sxs-lookup"><span data-stu-id="7494a-103">Implement append-on-send in your Outlook add-in</span></span>

<span data-ttu-id="7494a-104">本演练结束时，您将拥有一个可在发送邮件时插入免责声明的 Outlook 外接程序。</span><span class="sxs-lookup"><span data-stu-id="7494a-104">By the end of this walkthrough, you'll have an Outlook add-in that can insert a disclaimer when a message is sent.</span></span>

> [!NOTE]
> <span data-ttu-id="7494a-105">对此功能的支持是在要求集1.9 中引入的。</span><span class="sxs-lookup"><span data-stu-id="7494a-105">Support for this feature was introduced in requirement set 1.9.</span></span> <span data-ttu-id="7494a-106">请查看支持此要求集的[客户端和平台](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients)。</span><span class="sxs-lookup"><span data-stu-id="7494a-106">See [clients and platforms](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) that support this requirement set.</span></span>

## <a name="set-up-your-environment"></a><span data-ttu-id="7494a-107">设置环境</span><span class="sxs-lookup"><span data-stu-id="7494a-107">Set up your environment</span></span>

<span data-ttu-id="7494a-108">完成 [Outlook 快速入门](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) ，它将使用 Office 外接程序的 Yeoman 生成器创建外接程序项目。</span><span class="sxs-lookup"><span data-stu-id="7494a-108">Complete the [Outlook quick start](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) which creates an add-in project with the Yeoman generator for Office Add-ins.</span></span>

## <a name="configure-the-manifest"></a><span data-ttu-id="7494a-109">配置清单</span><span class="sxs-lookup"><span data-stu-id="7494a-109">Configure the manifest</span></span>

<span data-ttu-id="7494a-110">若要在您的外接程序中启用 "追加发送" 功能，必须 `AppendOnSend` 在 [ExtendedPermissions](../reference/manifest/extendedpermissions.md)集合中包含该权限。</span><span class="sxs-lookup"><span data-stu-id="7494a-110">To enable the append-on-send feature in your add-in, you must include the `AppendOnSend` permission in the collection of [ExtendedPermissions](../reference/manifest/extendedpermissions.md).</span></span>

<span data-ttu-id="7494a-111">对于此方案， `action` 您将运行函数，而不是在选择 " **执行操作** " 按钮时运行函数 `appendOnSend` 。</span><span class="sxs-lookup"><span data-stu-id="7494a-111">For this scenario, instead of running the `action` function on choosing the **Perform an action** button, you'll be running the `appendOnSend` function.</span></span>

1. <span data-ttu-id="7494a-112">在代码编辑器中，打开 "快速启动" 项目。</span><span class="sxs-lookup"><span data-stu-id="7494a-112">In your code editor, open the quick start project.</span></span>

1. <span data-ttu-id="7494a-113">打开位于项目根目录中的 **manifest.xml** 文件。</span><span class="sxs-lookup"><span data-stu-id="7494a-113">Open the **manifest.xml** file located at the root of your project.</span></span>

1. <span data-ttu-id="7494a-114">选择整个 `<VersionOverrides>` 节点 (包括 "打开" 和 "关闭" 标记) 并将其替换为以下 XML。</span><span class="sxs-lookup"><span data-stu-id="7494a-114">Select the entire `<VersionOverrides>` node (including open and close tags) and replace it with the following XML.</span></span>

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
> <span data-ttu-id="7494a-115">若要了解有关 Outlook 外接程序的清单的详细信息，请参阅 [outlook 外接程序清单](manifests.md)。</span><span class="sxs-lookup"><span data-stu-id="7494a-115">To learn more about manifests for Outlook add-ins, see [Outlook add-in manifests](manifests.md).</span></span>

## <a name="implement-append-on-send-handling"></a><span data-ttu-id="7494a-116">实现附加发送前处理</span><span class="sxs-lookup"><span data-stu-id="7494a-116">Implement append-on-send handling</span></span>

<span data-ttu-id="7494a-117">接下来，实现在 send 事件上追加。</span><span class="sxs-lookup"><span data-stu-id="7494a-117">Next, implement appending on the send event.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="7494a-118">如果您的外接程序还实现了[使用 `ItemSend` 中的发送事件处理](outlook-on-send-addins.md)，则在 `AppendOnSendAsync` 发送时处理程序中调用将返回错误，因为这种情况不受支持。</span><span class="sxs-lookup"><span data-stu-id="7494a-118">If your add-in also implements [on-send event handling using `ItemSend`](outlook-on-send-addins.md), calling `AppendOnSendAsync` in the on-send handler returns an error as this scenario isn't supported.</span></span>

<span data-ttu-id="7494a-119">在这种情况下，您将实现在用户发送时向项目追加免责声明。</span><span class="sxs-lookup"><span data-stu-id="7494a-119">For this scenario, you'll implement appending a disclaimer to the item when the user sends.</span></span>

1. <span data-ttu-id="7494a-120">在同一 "快速启动" 项目中，在代码编辑器中打开 **/src/commands/commands.js** 。</span><span class="sxs-lookup"><span data-stu-id="7494a-120">From the same quick start project, open the file **./src/commands/commands.js** in your code editor.</span></span>

1. <span data-ttu-id="7494a-121">在 `action` 函数后面，插入以下 JavaScript 函数。</span><span class="sxs-lookup"><span data-stu-id="7494a-121">After the `action` function, insert the following JavaScript function.</span></span>

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

1. <span data-ttu-id="7494a-122">在文件末尾，添加以下语句。</span><span class="sxs-lookup"><span data-stu-id="7494a-122">At the end of the file, add the following statement.</span></span>

    ```js
    g.appendDisclaimerOnSend = appendDisclaimerOnSend;
    ```

## <a name="try-it-out"></a><span data-ttu-id="7494a-123">试用</span><span class="sxs-lookup"><span data-stu-id="7494a-123">Try it out</span></span>

1. <span data-ttu-id="7494a-124">在项目的根目录中运行以下命令。</span><span class="sxs-lookup"><span data-stu-id="7494a-124">Run the following command in the root directory of your project.</span></span> <span data-ttu-id="7494a-125">运行此命令时，本地 web 服务器将启动（如果它尚未运行）。</span><span class="sxs-lookup"><span data-stu-id="7494a-125">When you run this command, the local web server will start if it's not already running.</span></span>

    ```command&nbsp;line
    npm run dev-server
    ```

1. <span data-ttu-id="7494a-126">按照 [旁加载 Outlook 外接程序](sideload-outlook-add-ins-for-testing.md)中的说明进行操作，以进行测试。</span><span class="sxs-lookup"><span data-stu-id="7494a-126">Follow the instructions in [Sideload Outlook add-ins for testing](sideload-outlook-add-ins-for-testing.md).</span></span>

1. <span data-ttu-id="7494a-127">创建新邮件，并将自己添加到 " **to** " 行。</span><span class="sxs-lookup"><span data-stu-id="7494a-127">Create a new message, and add yourself to the **To** line.</span></span>

1. <span data-ttu-id="7494a-128">从 "功能区" 或 "溢出" 菜单中，选择 " **执行操作**"。</span><span class="sxs-lookup"><span data-stu-id="7494a-128">From the ribbon or overflow menu, choose **Perform an action**.</span></span>

1. <span data-ttu-id="7494a-129">发送邮件，然后从 **"收件箱" 或 "** **已发送邮件** " 文件夹中打开它以查看追加的免责声明。</span><span class="sxs-lookup"><span data-stu-id="7494a-129">Send the message, then open it from your **Inbox** or **Sent Items** folder to view the appended disclaimer.</span></span>

    ![在 Outlook 网页版上追加的包含免责声明的示例邮件的屏幕截图。](../images/outlook-web-append-disclaimer.png)

## <a name="see-also"></a><span data-ttu-id="7494a-131">另请参阅</span><span class="sxs-lookup"><span data-stu-id="7494a-131">See also</span></span>

[<span data-ttu-id="7494a-132">Outlook 加载项清单</span><span class="sxs-lookup"><span data-stu-id="7494a-132">Outlook add-in manifests</span></span>](manifests.md)
