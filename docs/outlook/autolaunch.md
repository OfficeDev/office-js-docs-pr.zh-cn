---
title: '为 Outlook 外接程序配置基于事件的激活 (预览) '
description: 了解如何配置 Outlook 外接程序进行基于事件的激活。
ms.topic: article
ms.date: 03/30/2021
localization_priority: Normal
ms.openlocfilehash: b9a4460b05af14f57eecb1bf4181c706843537b2
ms.sourcegitcommit: 074526a6dca8381dbdabf2705474c5ae6753b829
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/02/2021
ms.locfileid: "51506145"
---
# <a name="configure-your-outlook-add-in-for-event-based-activation-preview"></a>为 Outlook 外接程序配置基于事件的激活 (预览) 

如果没有基于事件的激活功能，用户必须显式启动外接程序才能完成其任务。 此功能使加载项能够运行基于特定事件的任务，尤其是适用于每个项目的操作。 还可以与任务窗格和无 UI 功能集成。 目前，支持以下事件。

- `OnNewMessageCompose`：撰写新邮件时 (包括答复、全部答复和转发) 
- `OnNewAppointmentOrganizer`：在创建新约会时

  > [!IMPORTANT]
  > 在编辑 **项目** （例如草稿或现有约会）时，此功能不会激活。

在此演练结束时，您将具有一个在新建邮件时运行的外接程序。

> [!IMPORTANT]
> 此功能仅在 Outlook [网页版](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) 和具有 Microsoft 365 订阅的 Windows 上受预览支持。 有关详细信息 [，请参阅本文](#how-to-preview-the-event-based-activation-feature) 中的如何预览基于事件的激活功能。
>
> 由于预览功能可能会随时更改，恕不另行通知，因此不应在生产外接程序中使用。

## <a name="how-to-preview-the-event-based-activation-feature"></a>如何预览基于事件的激活功能

我们邀请你试用基于事件的激活功能！ 请告诉我们你的方案，以及我们如何通过 GitHub 提供反馈来改进 (请参阅此页面末尾的反馈部分) 。 

预览此功能：

- 对于 Outlook 网页：
  - [在 Microsoft 365 租户上配置定向发布](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)。
  - 在 **CDN 服务器上** 引用 beta https://appsforoffice.microsoft.com/lib/beta/hosted/office.js) (。 TypeScript [编译和](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts) 定义的类型定义IntelliSense CDN 和 [DefinitelyTyped 找到](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts)。 可以使用 安装这些类型 `npm install --save-dev @types/office-js-preview` 。
- 对于 Windows 版 Outlook：最低要求版本为 16.0.13729.20000。 加入 [Office 预览体验计划](https://insider.office.com) 以访问 Office beta 版本。

## <a name="set-up-your-environment"></a>设置环境

完成 [Outlook 快速入门](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) ，使用适用于 Office 外接程序的 Yeoman 生成器创建外接程序项目。

## <a name="configure-the-manifest"></a>配置清单

若要启用加载项的基于事件的激活，必须在清单节点中配置 [Runtimes](../reference/manifest/runtimes.md) 元素和 [LaunchEvent](../reference/manifest/extensionpoint.md#launchevent-preview) `VersionOverridesV1_1` 扩展点。 目前， `DesktopFormFactor` 是唯一受支持的外形类型。

1. 在代码编辑器中，打开快速启动项目。

1. 打开 **manifest.xml** 根目录下的文件。

1. 选择整个节点 (包括打开和关闭) `<VersionOverrides>` 并将其替换为以下 XML，然后保存更改。

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

Windows 上的 Outlook 使用 JavaScript 文件，而 Web 上的 Outlook 使用可引用同一 JavaScript 文件的 HTML 文件。 您必须在清单的节点中提供对这些文件的引用，因为 Outlook 平台最终确定是使用 HTML 还是基于 Outlook 客户端 `Resources` 的 JavaScript。 因此，若要配置事件处理，请提供 HTML 在 元素中的位置，然后在其子元素中提供 JavaScript 文件内附或 HTML `Runtime` `Override` 引用的位置。

> [!TIP]
> 若要了解有关 Outlook 外接程序的清单的更多信息，请参阅 [Outlook 外接程序清单](manifests.md)。

## <a name="implement-event-handling"></a>实现事件处理

您必须对所选事件实现处理。

在此方案中，您将添加用于撰写新项的处理。

1. 从同一快速启动项目中，在代码编辑器中打开 **commands.js./src/commands/commands.js** 文件。

1. 在 函数 `action` 之后，插入以下 JavaScript 函数。

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

1. 若要使用 Yeoman 生成器为 Office 加载项生成的此项目在 **Outlook** 网页 Outlook 中运行的功能，在文件末尾添加以下语句。

    ```js
    g.onMessageComposeHandler = onMessageComposeHandler;
    g.onAppointmentComposeHandler = onAppointmentComposeHandler;
    ```

1. 若要使函数在 Windows 上的 **Outlook 中** 运行，在文件末尾添加以下 JavaScript 代码。

    ```js
    if (Office.actions) {
      // 1st parameter: FunctionName of LaunchEvent in the manifest; 2nd parameter: Its implementation in this .js file.
      Office.actions.associate("onMessageComposeHandler", onMessageComposeHandler);
      Office.actions.associate("onAppointmentComposeHandler", onAppointmentComposeHandler);
    }
    ```

    **注意**：检查 `Office.actions` 以确保 Web 上的 Outlook 忽略这些语句。

1. 保存所做的更改。

## <a name="try-it-out"></a>试用

1. 在项目的根目录中运行以下命令。 如果运行此命令，本地 Web 服务器将启动（如果尚未运行），并将旁加载加载项。

    ```command&nbsp;line
    npm start
    ```

1. 在 Outlook 网页版中，创建新邮件。

    ![在撰写时设置了主题的 Outlook 网页中的邮件窗口屏幕截图](../images/outlook-web-autolaunch-1.png)

1. 在 Windows 上的 Outlook 中，创建新邮件。

    ![Windows 上的 Outlook 中邮件窗口的屏幕截图，撰写时主题已设置](../images/outlook-win-autolaunch.png)

    > [!NOTE]
    > 如果看到错误"无法从 localhost 打开此外接程序"，则需要启用环回豁免。
    >
    > 1. 关闭 Outlook。
    > 2. 打开 **任务管理器** ， **并确保msoadfs.exe进程** 未运行。
    > 3. 运行以下命令。
    >
    >     ```command&nbsp;line
    >     call %SystemRoot%\System32\CheckNetIsolation.exe LoopbackExempt -a -n=1_http___localhost_300004ACA5EC-D79A-43EA-AB47-E50E47DD96FC
    >     ```
    >
    > 4. 重新启动 Outlook。

## <a name="debug"></a>Debug

当你实现自己的功能时，你可能需要调试代码。 有关如何调试基于事件的外接程序激活的指南，请参阅调试基于事件的 [Outlook 外接程序](debug-autolaunch.md)。

## <a name="event-based-activation-behavior-and-limitations"></a>基于事件的激活行为和限制

基于事件激活的加载项应尽量短运行、轻量且非轻量。 若要表示加载项已完成处理启动事件，建议让加载项调用 `event.completed` 方法。 如果未进行该调用，外接程序将在大约 300 秒（运行基于事件的外接程序所允许的最大时间长度）内退出。当用户关闭撰写窗口时，外接程序也将结束。

如果用户有多个订阅了同一事件的加载项，Outlook 平台将按特定顺序启动加载项。 目前，只能主动运行五个基于事件的加载项。 任何其他外接程序将推送到队列，然后随着之前处于活动状态的外接程序完成或停用而运行。

用户可以切换或导航离开外接程序开始运行的当前邮件项目。 启动的加载项将在后台完成其操作。

某些Office.js更改或更改 UI 的 API 不允许来自基于事件的外接程序。以下是阻止的 API：

- 在 `Office.context.auth` 下：
  - `getAccessToken`
  - `getAccessTokenAsync`
- 在 `Office.context.mailbox` 下：
  - `displayAppointmentForm`
  - `displayMessageForm`
  - `displayNewAppointmentForm`
  - `displayNewMessageForm`
- 在 `Office.context.mailbox.item` 下：
  - `close`
- 在 `Office.context.ui` 下：
  - `displayDialogAsync`
  - `messageParent`

## <a name="see-also"></a>另请参阅

[Outlook 外接程序清单](manifests.md) 
[如何调试基于事件的外接程序](debug-autolaunch.md)
