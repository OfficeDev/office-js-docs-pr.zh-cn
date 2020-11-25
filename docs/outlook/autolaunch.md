---
title: '配置 Outlook 外接程序以进行基于事件的激活 (预览) '
description: 了解如何配置 Outlook 外接程序以进行基于事件的激活。
ms.topic: article
ms.date: 11/24/2020
localization_priority: Normal
ms.openlocfilehash: d7ba4a0fb87ec51db56892f4eb3002ae5b7fa6ec
ms.sourcegitcommit: f4fa1a0187466ea136009d1fe48ec67e4312c934
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/25/2020
ms.locfileid: "49408839"
---
# <a name="configure-your-outlook-add-in-for-event-based-activation-preview"></a>配置 Outlook 外接程序以进行基于事件的激活 (预览) 

如果没有基于事件的激活功能，用户必须显式启动外接程序以完成其任务。 此功能使加载项能够根据特定事件（尤其是适用于每个项目的操作）运行任务。 您还可以与任务窗格和无 UI 功能集成。 目前，支持以下事件。

- `OnNewMessageCompose`：撰写新邮件时 (包括答复、全部答复和转发) 
- `OnNewAppointmentOrganizer`：创建新约会时

  > [!IMPORTANT]
  > 此 **功能不会激活编辑** 项目（例如，草稿或现有约会）。

本演练结束时，您将拥有一个在创建新邮件时运行的外接程序。

> [!IMPORTANT]
> 只有使用 Microsoft 365 订阅的 Outlook 网页版中的 [预览](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) 才支持此功能。 有关更多详细信息，请参阅 [如何预览本文中基于事件的激活功能](#how-to-preview-the-event-based-activation-feature) 。
>
> 由于预览功能可能会发生更改，恕不另行通知，它们不应在生产外接程序中使用。

## <a name="how-to-preview-the-event-based-activation-feature"></a>如何预览基于事件的激活功能

我们邀请你试用基于事件的激活功能！ 请通过 GitHub 向我们提供反馈，告知我们你的方案以及我们如何改进， (请参阅本页结尾处的 **反馈** 部分) 。

若要预览此功能：

- 参考 CDN (上的 **beta** 库 https://appsforoffice.microsoft.com/lib/beta/hosted/office.js) 。 在 CDN 和[jquery.typescript.definitelytyped](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts)中找到 TypeScript 编译和智能感知的[类型定义文件](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts)。 您可以使用安装这些类型 `npm install --save-dev @types/office-js-preview` 。
- [在 Microsoft 365 租户上配置目标版本](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)。

## <a name="set-up-your-environment"></a>设置环境

完成 [Outlook 快速入门](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) ，它将使用 Office 外接程序的 Yeoman 生成器创建外接程序项目。

## <a name="configure-the-manifest"></a>配置清单

若要启用您的外接程序的基于事件的激活，必须在清单中配置 [运行时](../reference/manifest/runtimes.md) 元素和 [LaunchEvent](../reference/manifest/extensionpoint.md#launchevent-preview) 扩展点。 目前， `DesktopFormFactor` 是唯一受支持的板型。

1. 在代码编辑器中，打开 "快速启动" 项目。

1. 打开位于项目根目录中的 **manifest.xml** 文件。

1. 选择整个 `<VersionOverrides>` 节点 (包括 "打开" 和 "关闭" 标记) 并将其替换为以下 XML。

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

Windows 上的 outlook 使用 JavaScript 文件，而 web 上的 Outlook 使用引用相同 JavaScript 文件的 HTML 文件。 由于 Outlook 平台最终决定是使用基于 Outlook 客户端的 HTML 还是 JavaScript，因此您必须在清单中提供对这些文件的引用。 因此，若要配置事件处理，请在元素中提供 HTML 的位置 `Runtime` ，然后在其子 `Override` 元素中提供由 html 内联或引用的 JavaScript 文件的位置。

> [!TIP]
> 若要了解有关 Outlook 外接程序的清单的详细信息，请参阅 [outlook 外接程序清单](manifests.md)。

## <a name="implement-event-handling"></a>实现事件处理

您必须为选定的事件实现处理。

在这种情况下，您将添加用于撰写新项目的处理。

1. 在同一 "快速启动" 项目中，在代码编辑器中打开 **/src/commands/commands.js** 。

1. 在 `action` 函数后面，插入以下 JavaScript 函数。

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

1. 在文件末尾，添加以下语句。

    ```js
    g.onMessageComposeHandler = onMessageComposeHandler;
    g.onAppointmentComposeHandler = onAppointmentComposeHandler;
    ```

## <a name="try-it-out"></a>试用

1. 在项目的根目录中运行以下命令。运行此命令时，本地 Web 服务器将启动（如果尚未运行）。

    ```command&nbsp;line
    npm run dev-server
    ```

1. 按照[旁加载 Outlook 加载项以供测试](sideload-outlook-add-ins-for-testing.md)中的说明操作，旁加载加载项。

1. 在 Outlook 网页版中，创建新邮件。

    ![Outlook 网页版中邮件窗口的屏幕截图，其中的主题设置为撰写。](../images/outlook-web-autolaunch.png)

## <a name="event-based-activation-behavior-and-limitations"></a>基于事件的激活行为和限制

根据事件激活的加载项设计为运行时间较短，最长为330秒。 我们建议您让您的外接程序调用 `event.completed` 方法，以通知其已完成启动事件的处理。 当用户关闭撰写窗口时，外接端也会结束。

如果用户具有多个订阅同一事件的加载项，则 Outlook 平台将以无特定的顺序启动外接程序。 目前，只有五个基于事件的外接程序可以处于活动状态。 任何其他外接程序将被推送到队列中，然后运行之前的活动外接程序已完成或停用。

用户可以从加载项开始运行的当前邮件项目中进行切换或导航。 启动的外接程序将在后台完成其操作。

基于事件的外接程序不允许更改或更改 UI 的一些 Office.js Api。以下是阻止的 Api。

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
