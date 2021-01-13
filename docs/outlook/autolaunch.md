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
# <a name="configure-your-outlook-add-in-for-event-based-activation-preview"></a>为 Outlook 外接程序配置基于事件的激活和 (预览) 

如果没有基于事件的激活功能，用户必须显式启动加载项才能完成其任务。 利用此功能，您的外接程序可以基于特定事件运行任务，尤其是适用于每个项目的操作。 还可以与任务窗格和无 UI 功能集成。 目前，支持以下事件。

- `OnNewMessageCompose`：撰写新邮件时 (包括答复、全部答复和转发) 
- `OnNewAppointmentOrganizer`：创建新约会时

  > [!IMPORTANT]
  > 在编辑 **项目** （例如草稿或现有约会）时，此功能不会激活。

在此演练结束时，您将具有一个在新建邮件时运行的外接程序。

> [!IMPORTANT]
> 此功能仅在具有 Microsoft [](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) 365 订阅的 Outlook 网页版中受预览支持。 请参阅 [本文中](#how-to-preview-the-event-based-activation-feature) 如何预览基于事件的激活功能，了解更多详细信息。
>
> 由于预览功能可能会随时更改，恕不另行通知，因此不应在生产外接程序中使用。

## <a name="how-to-preview-the-event-based-activation-feature"></a>如何预览基于事件的激活功能

我们邀请你试用基于事件的激活功能！ 通过 GitHub 提供反馈，让我们了解你的方案以及如何改进 (请参阅此页面末尾的"反馈"部分) 。 

预览此功能：

- 引用 **CDN** 版本上的 beta https://appsforoffice.microsoft.com/lib/beta/hosted/office.js) (。 TypeScript [编译和](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts) 编译的类型IntelliSense CDN 和 [DefinitelyTyped 找到](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts)。 可以安装这些类型 `npm install --save-dev @types/office-js-preview` 。
- [在 Microsoft 365 租户](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)上配置定向发布。

## <a name="set-up-your-environment"></a>设置环境

使用 [适用于 Office](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) 加载项的 Yeoman 生成器完成创建外接程序项目的 Outlook 快速入门。

## <a name="configure-the-manifest"></a>配置清单

若要启用加载项的基于事件的激活，必须在清单中配置 [Runtimes](../reference/manifest/runtimes.md) 元素和 [LaunchEvent](../reference/manifest/extensionpoint.md#launchevent-preview) 扩展点。 目前， `DesktopFormFactor` 是唯一受支持的外形类型。

1. 在代码编辑器中，打开快速启动项目。

1. 打开 **manifest.xml** 根目录下的文件。

1. 选择整个 `<VersionOverrides>` 节点 (包括打开和关闭) 并将其替换为以下 XML。

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

Windows 上的 Outlook 使用 JavaScript 文件，而 Web 上的 Outlook 使用引用同一 JavaScript 文件的 HTML 文件。 您必须在清单中提供对这两个文件的引用，因为 Outlook 平台最终确定是使用基于 Outlook 客户端的 HTML 还是 JavaScript。 因此，若要配置事件处理，请提供 HTML 在元素中的位置，然后在其子元素中提供 HTML 内附或引用 `Runtime` `Override` 的 JavaScript 文件的位置。

> [!TIP]
> 若要了解有关 Outlook 外接程序清单的更多信息，请参阅 [Outlook 外接程序清单](manifests.md)。

## <a name="implement-event-handling"></a>实现事件处理

您必须对选定的事件实现处理。

在此方案中，将添加用于撰写新项的处理。

1. 从同一快速启动项目中，在代码编辑器中打开commands.js **./src/commands/commands.js** 文件。

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

1. 在文件末尾，添加以下语句。

    ```js
    g.onMessageComposeHandler = onMessageComposeHandler;
    g.onAppointmentComposeHandler = onAppointmentComposeHandler;
    ```

## <a name="try-it-out"></a>试用

1. 在项目的根目录中运行以下命令。 运行此命令时，本地 Web 服务器将启动（如果尚未运行）。

    ```command&nbsp;line
    npm run dev-server
    ```

1. 按照[旁加载 Outlook 加载项以供测试](sideload-outlook-add-ins-for-testing.md)中的说明操作，旁加载加载项。

1. 在 Outlook 网页版中，创建新邮件。

    ![Outlook 网页中的邮件窗口的屏幕截图，主题在撰写时设置。](../images/outlook-web-autolaunch.png)

## <a name="event-based-activation-behavior-and-limitations"></a>基于事件的激活行为和限制

基于事件激活的外接程序设计为短时间运行，最多 330 秒。 我们建议你让加载项调用该方法，以表明它 `event.completed` 已完成处理启动事件。 当用户关闭撰写窗口时，外接程序也会结束。

如果用户有多个订阅同一事件的加载项，Outlook 平台将按特定顺序启动外接程序。 目前，只能主动运行五个基于事件的加载项。 任何其他加载项将推送到队列，然后随着之前处于活动状态的加载项完成或停用而运行。

用户可以切换或导航离开外接程序开始运行的当前邮件项目。 启动的加载项将在后台完成其操作。

某些Office.js更改或更改 UI 的 API 不允许来自基于事件的加载项。下面是阻止的 API。

- Under `Office.context.mailbox` ：
  - `displayAppointmentForm`
  - `displayMessageForm`
  - `displayNewAppointmentForm`
  - `displayNewMessageForm`
- Under `Office.context.ui` ：
  - `displayDialogAsync`
  - `messageParent`
- Under `Office.context.auth` ：
  - `getAccessToken`
  - `getAccessTokenAsync`

## <a name="see-also"></a>另请参阅

[Outlook 加载项清单](manifests.md)
