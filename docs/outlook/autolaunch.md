---
title: '为 Outlook 外接程序配置基于事件的激活 (预览) '
description: 了解如何配置 Outlook 外接程序进行基于事件的激活。
ms.topic: article
ms.date: 02/03/2021
localization_priority: Normal
ms.openlocfilehash: d9108b4debea5e59503f3c935a537e5fafde00c8
ms.sourcegitcommit: fefc279b85e37463413b6b0e84c880d9ed5d7ac3
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/12/2021
ms.locfileid: "50234273"
---
# <a name="configure-your-outlook-add-in-for-event-based-activation-preview"></a>为 Outlook 外接程序配置基于事件的激活 (预览) 

如果没有基于事件的激活功能，用户必须显式启动加载项才能完成其任务。 利用此功能，加载项能够基于特定事件运行任务，尤其是适用于每个项目的操作。 还可以与任务窗格和无 UI 功能集成。 目前，支持以下事件。

- `OnNewMessageCompose`：撰写新邮件时 (包括答复、全部答复和转发) 
- `OnNewAppointmentOrganizer`：创建新约会时

  > [!IMPORTANT]
  > 在编辑 **项目** （例如草稿或现有约会）时，此功能不会激活。

在此演练结束时，您将拥有一个在新建邮件时运行的外接程序。

> [!IMPORTANT]
> 此功能仅在 Outlook [网页版](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) 和具有 Microsoft 365 订阅的 Windows 中受支持预览。 有关详细信息 [，请参阅](#how-to-preview-the-event-based-activation-feature) 本文中如何预览基于事件的激活功能。
>
> 由于预览功能可能会随时更改，恕不另行通知，因此不应将其用于生产外接程序。

## <a name="how-to-preview-the-event-based-activation-feature"></a>如何预览基于事件的激活功能

我们邀请你试用基于事件的激活功能！ 请告诉我们你的方案以及如何通过 GitHub 提供反馈来改进 (请参阅此页面末尾的"反馈"部分) 。 

预览此功能：

- 对于 Outlook 网页：
  - [在 Microsoft 365 租户上配置定向版本](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)。
  - 在 **CDN** 服务器上引用 beta https://appsforoffice.microsoft.com/lib/beta/hosted/office.js) (。 TypeScript [编译和键入](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts) IntelliSense在 CDN 和 [DefinitelyTyped 上找到](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts)。 可以使用 .安装这些类型 `npm install --save-dev @types/office-js-preview` 。
- 对于 Windows 版 Outlook：最低要求版本为 16.0.13729.20000。 加入 [Office 预览体验计划](https://insider.office.com) 以访问 Office beta 版本。

## <a name="set-up-your-environment"></a>设置环境

使用 [适用于 Office](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) 加载项的 Yeoman 生成器完成创建外接程序项目的 Outlook 快速入门。

## <a name="configure-the-manifest"></a>配置清单

若要启用加载项的基于事件的激活，必须在清单节点中配置 [Runtimes](../reference/manifest/runtimes.md) 元素和 [LaunchEvent](../reference/manifest/extensionpoint.md#launchevent-preview) `VersionOverridesV1_1` 扩展点。 目前， `DesktopFormFactor` 是唯一受支持的外形类型。

1. 在代码编辑器中，打开快速启动项目。

1. 打开 **manifest.xml** 根目录下的文件。

1. 选择整个 `<VersionOverrides>` 节点 (包括打开和关闭标记) 并将其替换为以下 XML。

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

Windows 上的 Outlook 使用 JavaScript 文件，而 Web 上的 Outlook 使用可引用同一 JavaScript 文件的 HTML 文件。 您必须在清单节点中提供对这两个文件的引用，因为 Outlook 平台最终确定是使用基于 Outlook 客户端的 HTML 还是 `Resources` JavaScript。 因此，若要配置事件处理，请提供 HTML 在元素中的位置，然后在其子元素中提供 HTML 内附或引用 `Runtime` `Override` 的 JavaScript 文件的位置。

> [!TIP]
> 若要了解有关 Outlook 外接程序清单的更多信息，请参阅 [Outlook 外接程序清单](manifests.md)。

## <a name="implement-event-handling"></a>实现事件处理

您必须对所选事件实现处理。

在此方案中，您将添加用于撰写新项的处理。

1. 从同一快速启动项目中，在代码编辑器中commands.js **./src/commands/commands.js** 文件。

1. 在函数 `action` 后插入以下 JavaScript 函数。

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

1. 若要使用由 Office 加载项的 Yeoman 生成器生成的此项目在 **Outlook** 网页 Outlook 中运行的函数，在文件末尾添加以下语句。

    ```js
    g.onMessageComposeHandler = onMessageComposeHandler;
    g.onAppointmentComposeHandler = onAppointmentComposeHandler;
    ```

1. 若要使函数在 Windows 上的 **Outlook 中运行**，在文件末尾添加以下 JavaScript 代码。

    ```js
    if (Office.actions) {
      // 1st parameter: FunctionName of LaunchEvent in the manifest; 2nd parameter: Its implementation in this .js file.
      Office.actions.associate("onMessageComposeHandler", onMessageComposeHandler);
      Office.actions.associate("onAppointmentComposeHandler", onAppointmentComposeHandler);
    }
    ```

    **注意**：检查 `Office.actions` 以确保 Web 上的 Outlook 忽略这些语句。

## <a name="try-it-out"></a>试用

1. 在项目的根目录中运行以下命令。 运行此命令时，如果本地 Web (尚未运行，) 将旁加载您的外接程序。

    ```command&nbsp;line
    npm start
    ```

1. 在 Outlook 网页版中，创建新邮件。

    ![Outlook 网页邮件窗口的屏幕截图，撰写时主题已设置](../images/outlook-web-autolaunch-1.png)

1. 在 Windows 上的 Outlook 中，创建新邮件。

    ![Windows 上的 Outlook 中邮件窗口的屏幕截图，撰写时主题已设置](../images/outlook-win-autolaunch.png)

## <a name="event-based-activation-behavior-and-limitations"></a>基于事件的激活行为和限制

基于事件激活的外接程序应尽可能短运行、轻型和非高空。 若要指示加载项已完成对启动事件的处理，建议让加载项调用 `event.completed` 该方法。 如果未进行该调用，加载项将在大约 300 秒（运行基于事件的加载项所允许的最大时间长度）内退出。当用户关闭撰写窗口时，加载项也会结束。

如果用户有多个订阅同一事件的加载项，则 Outlook 平台将启动外接程序，而没有任何特定顺序。 目前，只能主动运行五个基于事件的加载项。 任何其他加载项将推送到队列，然后随着之前处于活动状态的加载项完成或停用而运行。

用户可以切换或导航离开加载项开始运行的当前邮件项目。 启动的加载项将在后台完成其操作。

某些Office.js更改或更改 UI 的 API 不允许来自基于事件的加载项。以下是阻止的 API：

- 在 `Office.context.mailbox` ：
  - `displayAppointmentForm`
  - `displayMessageForm`
  - `displayNewAppointmentForm`
  - `displayNewMessageForm`
- 在 `Office.context.ui` ：
  - `displayDialogAsync`
  - `messageParent`
- 在 `Office.context.auth` ：
  - `getAccessToken`
  - `getAccessTokenAsync`

## <a name="see-also"></a>另请参阅

[Outlook 加载项清单](manifests.md)