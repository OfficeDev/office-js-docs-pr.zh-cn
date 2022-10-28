---
title: 为基于事件的激活配置 Outlook 加载项
description: 了解如何为基于事件的激活配置 Outlook 加载项。
ms.topic: article
ms.date: 10/24/2022
ms.localizationpriority: medium
ms.openlocfilehash: b5ae744350389ed222284808a67a9b7c30211136
ms.sourcegitcommit: 693e9a9b24bb81288d41508cb89c02b7285c4b08
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/28/2022
ms.locfileid: "68767173"
---
# <a name="configure-your-outlook-add-in-for-event-based-activation"></a>为基于事件的激活配置 Outlook 加载项

如果没有基于事件的激活功能，用户必须显式启动加载项才能完成任务。 此功能使外接程序能够基于特定事件运行任务，尤其是适用于每个项的操作。 还可以与任务窗格和函数命令集成。

在本演练结束时，你将有一个加载项，每当创建新项目并设置主题时，该加载项就会运行。

> [!NOTE]
> 要求 [集 1.10](/javascript/api/requirement-sets/outlook/requirement-set-1.10/outlook-requirement-set-1.10) 中引入了对此功能的支持，后续要求集中现在提供了其他事件。 有关事件的最低要求集以及支持它的客户端和平台的详细信息，请参阅 [支持的事件](#supported-events) 和 [Exchange 服务器和 Outlook 客户端支持的要求集](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets#requirement-sets-supported-by-exchange-servers-and-outlook-clients)。
>
> iOS 或 Android 版 Outlook 不支持基于事件的激活。

> [!IMPORTANT]
> Office 外接程序的 Teams 清单尚不支持基于事件的激活 [， (预览版) ](../develop/json-manifest-overview.md)。 我们正在努力尽快提供这种支持。

## <a name="supported-events"></a>支持的事件

下表列出了当前可用的事件以及每个事件支持的客户端。 引发事件时，处理程序会收到一个 `event` 对象，该对象可能包含特定于事件类型的详细信息。 **“说明”** 列包含指向相关对象的链接（如果适用）。

|事件规范名称</br>和 XML 清单名称|Teams 清单名称|说明|最低要求集和支持的客户端|
|---|---|---|---|
|`OnNewMessageCompose`| newMessageComposeCreated |撰写新邮件 (包括答复、全部答复和转发) 但不在编辑时，例如草稿。|[1.10](/javascript/api/requirement-sets/outlook/requirement-set-1.10/outlook-requirement-set-1.10)<br><br>- Windows<sup>1</sup><br>- Web 浏览器<br>- 新建 Mac UI |
|`OnNewAppointmentOrganizer`|newAppointmentOrganizerCreated|创建新约会时，但不在编辑现有约会时。|[1.10](/javascript/api/requirement-sets/outlook/requirement-set-1.10/outlook-requirement-set-1.10)<br><br>- Windows<sup>1</sup><br>- Web 浏览器<br>- 新建 Mac UI |
|`OnMessageAttachmentsChanged`|messageAttachmentsChanged|在撰写邮件时添加或删除附件。<br><br>特定于事件的数据对象： [AttachmentsChangedEventArgs](/javascript/api/outlook/office.attachmentschangedeventargs?view=outlook-js-1.11&preserve-view=true)|[1.11](/javascript/api/requirement-sets/outlook/requirement-set-1.11/outlook-requirement-set-1.11)<br><br>- Windows<sup>1</sup><br>- Web 浏览器<br>- 新建 Mac UI|
|`OnAppointmentAttachmentsChanged`|appointmentAttachmentsChanged|在撰写约会时添加或删除附件。<br><br>特定于事件的数据对象： [AttachmentsChangedEventArgs](/javascript/api/outlook/office.attachmentschangedeventargs?view=outlook-js-1.11&preserve-view=true)|[1.11](/javascript/api/requirement-sets/outlook/requirement-set-1.11/outlook-requirement-set-1.11)<br><br>- Windows<sup>1</sup><br>- Web 浏览器<br>- 新建 Mac UI|
|`OnMessageRecipientsChanged`|messageRecipientsChanged|在撰写邮件时添加或删除收件人。<br><br>特定于事件的数据对象： [RecipientsChangedEventArgs](/javascript/api/outlook/office.recipientschangedeventargs?view=outlook-js-1.11&preserve-view=true)|[1.11](/javascript/api/requirement-sets/outlook/requirement-set-1.11/outlook-requirement-set-1.11)<br><br>- Windows<sup>1</sup><br>- Web 浏览器<br>- 新建 Mac UI|
|`OnAppointmentAttendeesChanged`|appointmentAttendeesChanged|在撰写约会时添加或删除与会者。<br><br>特定于事件的数据对象： [RecipientsChangedEventArgs](/javascript/api/outlook/office.recipientschangedeventargs?view=outlook-js-1.11&preserve-view=true)|[1.11](/javascript/api/requirement-sets/outlook/requirement-set-1.11/outlook-requirement-set-1.11)<br><br>- Windows<sup>1</sup><br>- Web 浏览器<br>- 新建 Mac UI|
|`OnAppointmentTimeChanged`|appointmentTimeChanged|撰写约会时更改日期/时间。<br><br>特定于事件的数据对象： [AppointmentTimeChangedEventArgs](/javascript/api/outlook/office.appointmenttimechangedeventargs?view=outlook-js-1.11&preserve-view=true)|[1.11](/javascript/api/requirement-sets/outlook/requirement-set-1.11/outlook-requirement-set-1.11)<br><br>- Windows<sup>1</sup><br>- Web 浏览器<br>- 新建 Mac UI|
|`OnAppointmentRecurrenceChanged`|appointmentRecurrenceChanged|在撰写约会时添加、更改或删除定期详细信息。 如果更改了日期/时间， `OnAppointmentTimeChanged` 也会触发事件。<br><br>特定于事件的数据对象： [RecurrenceChangedEventArgs](/javascript/api/outlook/office.recurrencechangedeventargs?view=outlook-js-1.11&preserve-view=true)|[1.11](/javascript/api/requirement-sets/outlook/requirement-set-1.11/outlook-requirement-set-1.11)<br><br>- Windows<sup>1</sup><br>- Web 浏览器<br>- 新建 Mac UI|
|`OnInfoBarDismissClicked`|infoBarDismissClicked|撰写邮件或约会项目时关闭通知时。 只有添加通知的加载项才会收到通知。<br><br>特定于事件的数据对象： [InfobarClickedEventArgs](/javascript/api/outlook/office.infobarclickedeventargs?view=outlook-js-1.11&preserve-view=true)|[1.11](/javascript/api/requirement-sets/outlook/requirement-set-1.11/outlook-requirement-set-1.11)<br><br>- Windows<sup>1</sup><br>- Web 浏览器<br>- 新建 Mac UI|
|`OnMessageSend`|messageSending|发送消息项时。 若要了解详细信息，请参阅 [智能警报演练](smart-alerts-onmessagesend-walkthrough.md)。|[1.12](/javascript/api/requirement-sets/outlook/requirement-set-1.12/outlook-requirement-set-1.12)<br><br>- Windows<sup>1</sup><br>- Web 浏览器<br>- 新建 Mac UI|
|`OnAppointmentSend`|appointmentSending|发送约会项目时。 若要了解详细信息，请参阅 [智能警报演练](smart-alerts-onmessagesend-walkthrough.md)。|[1.12](/javascript/api/requirement-sets/outlook/requirement-set-1.12/outlook-requirement-set-1.12)<br><br>- Windows<sup>1</sup><br>- Web 浏览器<br>- 新建 Mac UI|
|`OnMessageCompose`|messageComposeOpened|撰写新邮件 (包括答复、全部答复和转发) 或编辑草稿。|[1.12](/javascript/api/requirement-sets/outlook/requirement-set-1.12/outlook-requirement-set-1.12)<br><br>- Windows<sup>1</sup><br>- Web 浏览器<br>- 新建 Mac UI|
|`OnAppointmentOrganizer`|appointmentOrganizerOpened|创建新约会或编辑现有约会时。|[1.12](/javascript/api/requirement-sets/outlook/requirement-set-1.12/outlook-requirement-set-1.12)<br><br>- Windows<sup>1</sup><br>- Web 浏览器<br>- 新建 Mac UI|

> [!NOTE]
> <sup>1</sup> Windows 版 Outlook 中基于事件的加载项至少需要Windows 10版本 1903 (内部版本 18362) 或 Windows Server 2019 版本 1903 才能运行。

## <a name="set-up-your-environment"></a>设置环境

完成 [Outlook 快速入门](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) ，该快速入门使用 Office 加载项的 [Yeoman 生成器创建外接程序](../develop/yeoman-generator-overview.md)项目。

## <a name="configure-the-manifest"></a>配置清单

若要配置清单，请选择要使用的清单类型的选项卡。

# <a name="xml-manifest"></a>[XML 清单](#tab/xmlmanifest)

若要启用加载项的基于事件的激活，必须在清单的节点中`VersionOverridesV1_1`配置 [Runtimes](/javascript/api/manifest/runtimes) 元素和 [LaunchEvent](/javascript/api/manifest/extensionpoint#launchevent) 扩展点。 目前， `DesktopFormFactor` 是唯一受支持的外形规格。

1. 在代码编辑器中，打开快速入门项目。

1. 打开位于项目根目录处的 **manifest.xml** 文件。

1. 选择整个 **\<VersionOverrides\>** 节点 (包括打开和关闭标记) ，并将其替换为以下 XML，然后保存所做的更改。

```XML
<VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
    <Requirements>
      <bt:Sets DefaultMinVersion="1.10">
        <bt:Set Name="Mailbox" />
      </bt:Sets>
    </Requirements>
    <Hosts>
      <Host xsi:type="MailHost">
        <!-- Event-based activation happens in a lightweight runtime.-->
        <Runtimes>
          <!-- HTML file including reference to or inline JavaScript event handlers.
               This is used by Outlook on the web and Outlook on the new Mac UI. -->
          <Runtime resid="WebViewRuntime.Url">
            <!-- JavaScript file containing event handlers. This is used by Outlook on Windows. -->
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
              <LaunchEvent Type="OnNewMessageCompose" FunctionName="onNewMessageComposeHandler"/>
              <LaunchEvent Type="OnNewAppointmentOrganizer" FunctionName="onNewAppointmentComposeHandler"/>
              
              <!-- Other available events -->
              <!--
              <LaunchEvent Type="OnMessageAttachmentsChanged" FunctionName="onMessageAttachmentsChangedHandler" />
              <LaunchEvent Type="OnAppointmentAttachmentsChanged" FunctionName="onAppointmentAttachmentsChangedHandler" />
              <LaunchEvent Type="OnMessageRecipientsChanged" FunctionName="onMessageRecipientsChangedHandler" />
              <LaunchEvent Type="OnAppointmentAttendeesChanged" FunctionName="onAppointmentAttendeesChangedHandler" />
              <LaunchEvent Type="OnAppointmentTimeChanged" FunctionName="onAppointmentTimeChangedHandler" />
              <LaunchEvent Type="OnAppointmentRecurrenceChanged" FunctionName="onAppointmentRecurrenceChangedHandler" />
              <LaunchEvent Type="OnInfoBarDismissClicked" FunctionName="onInfobarDismissClickedHandler" />
              <LaunchEvent Type="OnMessageSend" FunctionName="onMessageSendHandler" SendMode="PromptUser" />
              <LaunchEvent Type="OnAppointmentSend" FunctionName="onAppointmentSendHandler" SendMode="PromptUser" />
              <LaunchEvent Type="OnMessageCompose" FunctionName="onMessageComposeHandler" />
              <LaunchEvent Type="OnAppointmentOrganizer" FunctionName="onAppointmentOrganizerHandler" />
              -->
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
        <!-- Entry needed for Outlook on Windows. -->
        <bt:Url id="JSRuntime.Url" DefaultValue="https://localhost:3000/launchevent.js" />
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

Windows 上的 Outlook 使用 JavaScript 文件，而 Outlook 网页版 和 新的 Mac UI 上使用可以引用相同 JavaScript 文件的 HTML 文件。 您必须在清单的 `Resources` 节点中提供对这两个文件的引用，因为 Outlook 平台最终会根据 Outlook 客户端确定是使用 HTML 还是 JavaScript。 因此，若要配置事件处理，请在 元素中 **\<Runtime\>** 提供 HTML 的位置，然后在其 `Override` 子元素中提供 HTML 内联或引用的 JavaScript 文件的位置。

# <a name="teams-manifest-developer-preview"></a>[Teams 清单 (开发人员预览版) ](#tab/jsonmanifest)

> [!IMPORTANT]
> Office 外接程序的 Teams 清单尚不支持基于事件的激活 [ (预览版) ](../develop/json-manifest-overview.md)。 此选项卡供将来使用。

1. 打开 **manifest.json** 文件。

1. 将以下 对象添加到“extensions.runtimes”数组。 关于此标记，请注意以下几点：

   - 邮箱要求集的“minVersion”设置为“1.10”，因为本文前面的表指定这是支持 `OnNewMessageCompose` 和 `OnNewAppointmentCompose` 事件的要求集的最低版本。
   - 运行时的“id”设置为描述性名称“autorun_runtime”。
   - “code”属性具有设置为 HTML 文件的子“page”属性和设置为 JavaScript 文件的子“script”属性。 稍后的步骤将创建或编辑这些文件。 Office 根据平台使用这些值之一。
       - Windows 上的 Office 在仅限 JavaScript 的运行时中执行事件处理程序，该运行时直接加载 JavaScript 文件。
       - Office on Mac 和 Web 在浏览器运行时中执行处理程序，这将加载 HTML 文件。 该文件又包含一个 `<script>` 用于加载 JavaScript 文件的标记。
     有关详细信息，请参阅 [Office 外接程序中的运行时](../testing/runtimes.md)。
   - “lifetime”属性设置为“short”，这意味着运行时在触发其中一个事件时启动，并在处理程序完成时关闭。  (在某些情况下，运行时在处理程序完成之前关闭。 请参阅 [Office Add-ins.) 中的运行时](../testing/runtimes.md)
   - 有两种类型的“操作”可以在运行时中运行。 在后面的步骤中，你将创建与这些操作对应的函数。

    ```json
     {
        "requirements": {
            "capabilities": [
                {
                    "name": "Mailbox",
                    "minVersion": "1.10"
                }
            ]
        },
        "id": "autorun_runtime",
        "type": "general",
        "code": {
            "page": "https://localhost:3000/commands.html",
            "script": "https://localhost:3000/launchevent.js"
        },
        "lifetime": "short",
        "actions": [
            {
                "id": "onNewMessageComposeHandler",
                "type": "executeFunction",
                "displayName": "onNewMessageComposeHandler"
            },
            {
                "id": "onNewAppointmentComposeHandler",
                "type": "executeFunction",
                "displayName": "onNewAppointmentComposeHandler"
            }
        ]
    }
    ```

1. 将以下“autoRunEvents”数组添加为“extensions”数组中的 对象的属性。

    ```json
    "autoRunEvents": [
    
    ]
    ```

1. 将以下 对象添加到“autoRunEvents”数组。 “events”属性将处理程序映射到事件，如本文前面的表中所述。 处理程序名称必须与前面步骤中“actions”数组中对象的“id”属性中使用的名称匹配。

    ```json
      {
          "requirements": {
              "capabilities": [
                  {
                      "name": "Mailbox",
                      "minVersion": "1.10"
                  }
              ],
              "scopes": [
                  "mail"
              ]
          },
          "events": [
              {
                  "type": "newMessageComposeCreated",
                  "actionId": "onNewMessageComposeHandler"
              },
              {
                  "type": "newAppointmentOrganizerCreated",
                  "actionId": "onNewAppointmentComposeHandler"
              }
          ]
      }
    ```

---

> [!TIP]
>
> - 若要了解外接程序中的运行时，请参阅 [Office 外接程序中的运行时](../testing/runtimes.md)。
> - 若要详细了解 Outlook 外接程序清单，请参阅 [Outlook 外接程序清单](manifests.md)。

## <a name="implement-event-handling"></a>实现事件处理

必须对所选事件实现处理。

在此方案中，你将添加用于撰写新项的处理。

1. 在同一个快速入门项目中，在 **./src** 目录下创建名为 **startvent** 的新文件夹。

1. 在 **./src/startvent** 文件夹中，创建一个名为 **launchevent.js** 的新文件。

1. 在代码编辑器中打开文件 **./src/startvent/launchevent.js** 并添加以下 JavaScript 代码。

    ```js
    /*
    * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
    * See LICENSE in the project root for license information.
    */

    function onNewMessageComposeHandler(event) {
      setSubject(event);
    }
    function onNewAppointmentComposeHandler(event) {
      setSubject(event);
    }
    function setSubject(event) {
      Office.context.mailbox.item.subject.setAsync(
        "Set by an event-based add-in!",
        {
          "asyncContext": event
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

    // 1st parameter: FunctionName of LaunchEvent in the manifest; 2nd parameter: Its implementation in this .js file.
    Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
    Office.actions.associate("onNewAppointmentComposeHandler", onNewAppointmentComposeHandler);
    ```

1. 保存所做的更改。

> [!IMPORTANT]
> Windows：目前，在 JavaScript 文件中，你对基于事件的激活实现处理不支持导入。

## <a name="update-the-commands-html-file"></a>更新命令 HTML 文件

1. 在 **./src/commands** 文件夹中，打开 **commands.html**。

1. 紧接在结束 **头** 标记 (`</head>`) 之前，添加一个脚本条目以包含事件处理 JavaScript 代码。

    ```html
    <script type="text/javascript" src="../launchevent/launchevent.js"></script>
    ```

1. 保存所做的更改。

## <a name="update-webpack-config-settings"></a>更新 webpack 配置设置

1. 打开在项目的根目录中找到的 **webpack.config.js** 文件，然后完成以下步骤。

1. `plugins`在 对象中找到数组，`config`并在数组的开头添加此新对象。

    ```js
    new CopyWebpackPlugin({
      patterns: [
        {
          from: "./src/launchevent/launchevent.js",
          to: "launchevent.js",
        },
      ],
    }),
    ```

1. 保存所做的更改。

## <a name="try-it-out"></a>试用

1. 在项目的根目录中运行以下命令。 运行 `npm start`时，如果本地 Web 服务器尚未运行) 并且加载项将被旁加载，则本地 Web 服务器将启动 (。

    ```command&nbsp;line
    npm run build
    ```

    ```command&nbsp;line
    npm start
    ```

    > [!NOTE]
    > 如果加载项未自动旁加载，请按照 [旁加载 Outlook 加载项](../outlook/sideload-outlook-add-ins-for-testing.md#sideload-manually) 中的说明进行测试，在 Outlook 中手动旁加载加载项。

1. 在 Outlook 网页版中，创建新邮件。

    ![Outlook 网页版中的邮件窗口，主题在撰写时设置。](../images/outlook-web-autolaunch-1.png)

1. 在新 Mac UI 上的 Outlook 中，创建新邮件。

    ![Outlook 中新 Mac UI 中的邮件窗口，主题设置为撰写。](../images/outlook-mac-autolaunch.png)

1. 在 Outlook on Windows 中，创建新邮件。

    ![Outlook on Windows 中的邮件窗口，主题设置为撰写。](../images/outlook-win-autolaunch.png)

## <a name="debug"></a>调试

在加载项中对 launch-event 处理进行更改时，应注意：

- 如果更新了清单， [请删除加载项](sideload-outlook-add-ins-for-testing.md#remove-a-sideloaded-add-in)，然后再次旁加载它。 如果使用 Outlook on Windows，请关闭并重新打开 Outlook。
- 如果对清单以外的文件进行了更改，请关闭并重新打开 Outlook on Windows，或刷新运行Outlook 网页版的浏览器选项卡。

实现自己的功能时，可能需要调试代码。 有关如何调试基于事件的外接程序激活的指导，请参阅 [调试基于事件的 Outlook 外接程序](debug-autolaunch.md)。

Windows 上的此功能也提供运行时日志记录。 有关详细信息，请参阅 [使用运行时日志记录调试加载项](../testing/runtime-logging.md#runtime-logging-on-windows)。

[!INCLUDE [Loopback exemption note](../includes/outlook-loopback-exemption.md)]

## <a name="deploy-to-users"></a>部署到用户

可以通过Microsoft 365 管理中心上传清单来部署基于事件的外接程序。 在管理门户中，展开导航窗格中的 **“设置”** 部分，然后选择“ **集成应用**”。 在 **“集成应用** ”页上，选择“ **上传自定义应用”** 操作。

![Microsoft 365 管理中心上的“集成应用”页，其中突出显示了“上传自定义应用”操作。](../images/outlook-deploy-event-based-add-ins.png)

> [!IMPORTANT]
> 基于事件的外接程序仅限于管理员管理的部署。 用户无法从 AppSource 或应用内 Office 应用商店激活基于事件的加载项。 若要了解详细信息，请参阅 [基于事件的 Outlook 加载项的 AppSource 列表选项](autolaunch-store-options.md)。

[!INCLUDE [outlook-smart-alerts-deployment](../includes/outlook-smart-alerts-deployment.md)]

## <a name="event-based-activation-behavior-and-limitations"></a>基于事件的激活行为和限制

外接程序启动事件处理程序应为短运行、轻型且尽可能非侵入性。 激活后，加载项将在大约 300 秒内超时，这是运行基于事件的外接程序所允许的最大时间长度。若要指示加载项已完成处理启动事件，关联的事件处理程序必须调用 `event.completed` 方法。  (请注意，语句后 `event.completed` 包含的代码不保证运行。) 每次触发外接程序句柄的事件时，都会重新激活加载项并运行关联的事件处理程序，并且超时窗口会重置。 加载项在超时后结束，或者用户关闭撰写窗口或发送项目。

如果用户有多个订阅同一事件的加载项，Outlook 平台将不按特定顺序启动加载项。 目前，只有五个基于事件的加载项可以主动运行。

在所有受支持的 Outlook 客户端中，用户必须保留在激活加载项的当前邮件项目上，才能完成运行。 例如，从当前项目 (导航，切换到另一个撰写窗口或选项卡) 将终止加载项操作。 当用户发送他们正在撰写的邮件或约会时，加载项也会停止操作。

在 JavaScript 文件中，你对 Windows 客户端中基于事件的激活实现处理不支持导入。

某些Office.js API 不允许从基于事件的外接程序更改或更改 UI。以下是阻止的 API。

- 在 下 `Office.context.auth`：
  - `getAccessToken`
  - `getAccessTokenAsync`
    > [!NOTE]
    > 支持基于事件的激活和单一登录 (SSO) 的所有 Outlook 版本都支持 [OfficeRuntime.auth](/javascript/api/office-runtime/officeruntime.auth)，而 [Office.auth](/javascript/api/office/office.auth) 仅在某些 Outlook 版本中受支持。 有关详细信息，请参阅 [在使用基于事件的激活的 Outlook 外接程序中启用单一登录 (SSO) ](use-sso-in-event-based-activation.md)。
- 在 下 `Office.context.mailbox`：
  - `displayAppointmentForm`
  - `displayMessageForm`
  - `displayNewAppointmentForm`
  - `displayNewMessageForm`
- 在 下 `Office.context.mailbox.item`：
  - `close`
- 在 下 `Office.context.ui`：
  - `displayDialogAsync`
  - `messageParent`

### <a name="requesting-external-data"></a>请求外部数据

可以使用 [提取](https://developer.mozilla.org/docs/Web/API/Fetch_API) 等 API 或 [XMLHttpRequest (XHR) ](https://developer.mozilla.org/docs/Web/API/XMLHttpRequest)（一种发出 HTTP 请求以与服务器交互的标准 Web API）来请求外部数据。

请注意，在使用 XMLHttpRequest 对象时，必须使用其他安全措施，需要 [同一源策略](https://developer.mozilla.org/docs/Web/Security/Same-origin_policy) 和简单的 [CORS (跨源资源共享) ](https://developer.mozilla.org/docs/Web/HTTP/CORS)。

[简单的 CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS#simple_requests) 实现：

- 无法使用 Cookie。
- 仅支持简单方法，例如 `GET`、 `HEAD`和 `POST`。
- 接受具有字段名称 `Accept`、 `Accept-Language`或 `Content-Language`的简单标头。
- 可以使用 ， `Content-Type`前提是内容类型为 `application/x-www-form-urlencoded`、 `text/plain`或 `multipart/form-data`。
- 不能在 返回 `XMLHttpRequest.upload`的对象上注册事件侦听器。
- 不能在请求中使用 `ReadableStream` 对象。

> [!NOTE]
> 从版本 2201、内部版本 16.0.14813.10000) 开始，Outlook 网页版、Mac 和 Windows (提供完整的 CORS 支持。

## <a name="see-also"></a>另请参阅

- [Outlook 加载项清单](manifests.md)
- [如何调试基于事件的加载项](debug-autolaunch.md)
- [基于事件的 Outlook 加载项的 AppSource 列表选项](autolaunch-store-options.md)
- [智能警报和 OnMessageSend 演练](smart-alerts-onmessagesend-walkthrough.md)
- Office 外接程序代码示例：
  - [使用 Outlook 基于事件的激活来加密附件、处理会议请求与会者并对约会日期/时间更改做出反应](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/outlook-encrypt-attachments)
  - [使用 Outlook 基于事件的激活设置签名](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/outlook-set-signature)
  - [使用 Outlook 基于事件的激活标记外部收件人](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/outlook-tag-external)
  - [使用 Outlook 智能警报](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/outlook-check-item-categories)
