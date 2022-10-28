---
title: 为联机会议提供商创建 Outlook 加载项
description: 讨论如何为联机会议服务提供商设置 Outlook 加载项。
ms.topic: article
ms.date: 10/24/2022
ms.localizationpriority: medium
ms.openlocfilehash: 7c2cdb9f6369fd851a13fe45df132482b0ccdc0e
ms.sourcegitcommit: 693e9a9b24bb81288d41508cb89c02b7285c4b08
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/28/2022
ms.locfileid: "68767180"
---
# <a name="create-an-outlook-add-in-for-an-online-meeting-provider"></a>为联机会议提供商创建 Outlook 加载项

设置联机会议是 Outlook 用户的核心体验，使用 [Outlook 创建 Teams 会议](/microsoftteams/teams-add-in-for-outlook)很容易。 但是，使用非 Microsoft 服务在 Outlook 中创建联机会议可能很麻烦。 通过实现此功能，服务提供商可以简化其 Outlook 加载项用户的联机会议创建和加入体验。

> [!IMPORTANT]
> 具有 Microsoft 365 订阅的 Outlook 网页版、Windows、Mac、Android 和 iOS 支持此功能。

本文介绍如何设置 Outlook 加载项，使用户能够使用联机会议服务组织和加入会议。 在本文中，我们将使用虚构的联机会议服务提供商“Contoso”。

## <a name="set-up-your-environment"></a>设置环境

完成 [Outlook 快速入门](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) ，该快速入门使用 Office 加载项的 Yeoman 生成器创建外接程序项目。

## <a name="configure-the-manifest"></a>配置清单

若要使用户能够使用加载项创建联机会议，必须配置清单。 标记因两个变量而异：

- 目标平台的类型;移动或非移动。
- 清单的类型; [Office 外接程序的 XML 或 Teams 清单 (预览版) ](../develop/json-manifest-overview.md)。

如果外接程序使用 XML 清单，并且加载项仅在 Outlook 网页版、Windows 和 Mac 中受支持，请选择“**Windows、Mac、Web**”选项卡以获取指导。 但是，如果 Android 版和 iOS 版 Outlook 也支持您的外接程序，请选择“ **移动** ”选项卡。

如果外接程序使用 Teams 清单 (预览) ，请选择“ **Teams 清单 (开发人员预览版)** 选项卡。

> [!IMPORTANT]
> Teams 清单 (预览版) 尚不支持联机会议提供商。 我们正在努力尽快提供这种支持。

# <a name="windows-mac-web"></a>[Windows、Mac、Web](#tab/non-mobile)

1. 在代码编辑器中，打开创建的 Outlook 快速入门项目。

1. 打开位于项目根目录处的 **manifest.xml** 文件。

1. 选择整个 **\<VersionOverrides\>** 节点 (包括打开和关闭标记) ，并将其替换为以下 XML。

```xml
<VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
    <Description resid="residDescription"></Description>
    <Requirements>
      <bt:Sets>
        <bt:Set Name="Mailbox" MinVersion="1.3"/>
      </bt:Sets>
    </Requirements>
    <Hosts>
      <Host xsi:type="MailHost">
        <DesktopFormFactor>
          <FunctionFile resid="residFunctionFile"/>
          <ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
            <OfficeTab id="TabDefault">
              <Group id="apptComposeGroup">
                <Label resid="residDescription"/>
                <Control xsi:type="Button" id="insertMeetingButton">
                  <Label resid="residLabel"/>
                  <Supertip>
                    <Title resid="residLabel"/>
                    <Description resid="residTooltip"/>
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="icon-16"/>
                    <bt:Image size="32" resid="icon-32"/>
                    <bt:Image size="64" resid="icon-64"/>
                    <bt:Image size="80" resid="icon-80"/>
                  </Icon>
                  <Action xsi:type="ExecuteFunction">
                    <FunctionName>insertContosoMeeting</FunctionName>
                  </Action>
                </Control>
              </Group>
            </OfficeTab>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>
    </Hosts>
    <Resources>
      <bt:Images>
        <bt:Image id="icon-16" DefaultValue="https://contoso.com/assets/icon-16.png"/>
        <bt:Image id="icon-32" DefaultValue="https://contoso.com/assets/icon-32.png"/>
        <bt:Image id="icon-48" DefaultValue="https://contoso.com/assets/icon-48.png"/>
        <bt:Image id="icon-64" DefaultValue="https://contoso.com/assets/icon-64.png"/>
        <bt:Image id="icon-80" DefaultValue="https://contoso.com/assets/icon-80.png"/>
      </bt:Images>
      <bt:Urls>
        <bt:Url id="residFunctionFile" DefaultValue="https://contoso.com/commands.html"/>
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="residDescription" DefaultValue="Contoso meeting"/>
        <bt:String id="residLabel" DefaultValue="Add a contoso meeting"/>
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="residTooltip" DefaultValue="Add a contoso meeting to this appointment."/>
      </bt:LongStrings>
    </Resources>
  </VersionOverrides>
</VersionOverrides>
```

# <a name="mobile"></a>[移动设备](#tab/mobile)

为了允许用户从其移动设备创建联机会议，在父元素 **\<MobileFormFactor\>** 下的清单中配置 [了 MobileOnlineMeetingCommandSurface 扩展点](/javascript/api/manifest/extensionpoint#mobileonlinemeetingcommandsurface)。 其他外形规格不支持此扩展点。

1. 在代码编辑器中，打开创建的 Outlook 快速入门项目。

1. 打开位于项目根目录处的 **manifest.xml** 文件。

1. 选择整个 **\<VersionOverrides\>** 节点 (包括打开和关闭标记) ，并将其替换为以下 XML。

```xml
<VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
    <Description resid="residDescription"></Description>
    <Requirements>
      <bt:Sets>
        <bt:Set Name="Mailbox" MinVersion="1.3"/>
      </bt:Sets>
    </Requirements>
    <Hosts>
      <Host xsi:type="MailHost">
        <DesktopFormFactor>
          <FunctionFile resid="residFunctionFile"/>
          <ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
            <OfficeTab id="TabDefault">
              <Group id="apptComposeGroup">
                <Label resid="residDescription"/>
                <Control xsi:type="Button" id="insertMeetingButton">
                  <Label resid="residLabel"/>
                  <Supertip>
                    <Title resid="residLabel"/>
                    <Description resid="residTooltip"/>
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="icon-16"/>
                    <bt:Image size="32" resid="icon-32"/>
                    <bt:Image size="64" resid="icon-64"/>
                    <bt:Image size="80" resid="icon-80"/>
                  </Icon>
                  <Action xsi:type="ExecuteFunction">
                    <FunctionName>insertContosoMeeting</FunctionName>
                  </Action>
                </Control>
              </Group>
            </OfficeTab>
          </ExtensionPoint>
        </DesktopFormFactor>

        <MobileFormFactor>
          <FunctionFile resid="residFunctionFile"/>
          <ExtensionPoint xsi:type="MobileOnlineMeetingCommandSurface">
            <Control xsi:type="MobileButton" id="insertMeetingButton">
              <Label resid="residLabel"/>
              <Icon>
                <bt:Image size="25" scale="1" resid="icon-16"/>
                <bt:Image size="25" scale="2" resid="icon-16"/>
                <bt:Image size="25" scale="3" resid="icon-16"/>

                <bt:Image size="32" scale="1" resid="icon-32"/>
                <bt:Image size="32" scale="2" resid="icon-32"/>
                <bt:Image size="32" scale="3" resid="icon-32"/>

                <bt:Image size="48" scale="1" resid="icon-48"/>
                <bt:Image size="48" scale="2" resid="icon-48"/>
                <bt:Image size="48" scale="3" resid="icon-48"/>
              </Icon>
              <Action xsi:type="ExecuteFunction">
                <FunctionName>insertContosoMeeting</FunctionName>
              </Action>
            </Control>
          </ExtensionPoint>
        </MobileFormFactor>
      </Host>
    </Hosts>
    <Resources>
      <bt:Images>
        <bt:Image id="icon-16" DefaultValue="https://contoso.com/assets/icon-16.png"/>
        <bt:Image id="icon-32" DefaultValue="https://contoso.com/assets/icon-32.png"/>
        <bt:Image id="icon-48" DefaultValue="https://contoso.com/assets/icon-48.png"/>
        <bt:Image id="icon-64" DefaultValue="https://contoso.com/assets/icon-64.png"/>
        <bt:Image id="icon-80" DefaultValue="https://contoso.com/assets/icon-80.png"/>
      </bt:Images>
      <bt:Urls>
        <bt:Url id="residFunctionFile" DefaultValue="https://contoso.com/commands.html"/>
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="residDescription" DefaultValue="Contoso meeting"/>
        <bt:String id="residLabel" DefaultValue="Add a contoso meeting"/>
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="residTooltip" DefaultValue="Add a contoso meeting to this appointment."/>
      </bt:LongStrings>
    </Resources>
  </VersionOverrides>
</VersionOverrides>
```

# <a name="teams-manifest-developer-preview"></a>[Teams 清单 (开发人员预览版) ](#tab/jsonmanifest)

> [!IMPORTANT]
> Office 外接程序的 Teams 清单尚不支持联机会议提供商 [， (预览版) ](../develop/json-manifest-overview.md)。 此选项卡供将来使用。

1. 打开 **manifest.json** 文件。

1. 在“authorization.permissions.resourceSpecific”数组中找到 *第一个* 对象，并将其“name”属性设置为“MailboxItem.ReadWrite.User”。 完成后，它应如下所示。

    ```json
    {
        "name": "MailboxItem.ReadWrite.User",
        "type": "Delegated"
    }
    ```

1. 在“validDomains”数组中，将 URL 更改为“”https://contoso.com，这是虚构联机会议提供商的 URL。 完成后，数组应如下所示。

    ```json
    "validDomains": [
        "https://contoso.com"
    ],
    ```

1. 将以下 对象添加到“extensions.runtimes”数组。 对于此代码，请注意以下事项。

   - 邮箱要求集的“minVersion”设置为“1.3”，因此运行时不会在不支持此功能的平台和 Office 版本上启动。
   - 运行时的“id”设置为描述性名称“online_meeting_runtime”。
   - “code.page”属性设置为将加载函数命令的无 UI HTML 文件的 URL。
   - “lifetime”属性设置为“short”，这意味着运行时在选择函数命令按钮时启动，并在函数完成时关闭。  (在某些情况下，运行时在处理程序完成之前关闭。 请参阅 [Office Add-ins.) 中的运行时](../testing/runtimes.md)
   - 有一个操作来运行名为“insertContosoMeeting”的函数。 你将在后面的步骤中创建此函数。

    ```json
    {
        "requirements": {
            "capabilities": [
                {
                    "name": "Mailbox",
                    "minVersion": "1.3"
                }
            ],
            "formFactors": [
                "desktop"
            ]
        },
        "id": "online_meeting_runtime",
        "type": "general",
        "code": {
            "page": "https://contoso.com/commands.html"
        },
        "lifetime": "short",
        "actions": [
            {
                "id": "insertContosoMeeting",
                "type": "executeFunction",
                "displayName": "insertContosoMeeting"
            }
        ]
    }
    ```

1. 将“extensions.ribbons”数组替换为以下内容。 关于此标记，请注意以下几点。

   - 邮箱要求集的“minVersion”设置为“1.3”，因此功能区自定义项不会出现在不支持此功能的平台和 Office 版本上。
   - “contexts”数组指定功能区仅在会议详细信息组织者窗口中可用。
   - 会议详细信息组织者窗口的默认功能区选项卡上 (将有一个自定义控件组，) 标记为 **Contoso 会议**。
   - 该组将有一个标记为 **“添加 Contoso 会议**”的按钮。
   - 按钮的“actionId”已设置为“insertContosoMeeting”，这与在上一步中创建的操作的“id”匹配。

    ```json
    "ribbons": [
      {
        "requirements": {
            "capabilities": [
                {
                    "name": "Mailbox",
                    "minVersion": "1.3"
                }
            ],
            "scopes": [
                "mail"
            ],
            "formFactors": [
                "desktop"
            ]
        },
        "contexts": [
            "meetingDetailsOrganizer"
        ],
        "tabs": [
            {
                "builtInTabId": "TabDefault",
                "groups": [
                    {
                        "id": "apptComposeGroup",
                        "label": "Contoso meeting",
                        "controls": [
                            {
                                "id": "insertMeetingButton",
                                "type": "button",
                                "label": "Add a Contoso meeting",
                                "icons": [
                                    {
                                        "size": 16,
                                        "file": "icon-16.png"
                                    },
                                    {
                                        "size": 32,
                                        "file": "icon-32.png"
                                    },
                                    {
                                        "size": 64,
                                        "file": "icon-64_02.png"
                                    },
                                    {
                                        "size": 80,
                                        "file": "icon-80.png"
                                    }
                                ],
                                "supertip": {
                                    "title": "Add a Contoso meeting",
                                    "description": "Add a Contoso meeting to this appointment."
                                },
                                "actionId": "insertContosoMeeting",
                            }
                        ]
                    }
                ]
            }
        ]
      }
    ]
    ```

---

> [!TIP]
> 若要了解有关 Outlook 外接程序清单的详细信息，请参阅 [Outlook 外接程序清单](manifests.md) 和 [为 Outlook Mobile 添加对外接程序命令的支持](add-mobile-support.md)。

## <a name="implement-adding-online-meeting-details"></a>实现添加联机会议详细信息

在本部分中，了解加载项脚本如何更新用户的会议以包含联机会议详细信息。 以下内容适用于所有受支持的平台。

1. 在同一快速入门项目中，在代码编辑器中打开文件 **./src/commands/commands.js** 。

1. 将 **commands.js** 文件的全部内容替换为以下 JavaScript。

    ```js
    // 1. How to construct online meeting details.
    // Not shown: How to get the meeting organizer's ID and other details from your service.
    const newBody = '<br>' +
        '<a href="https://contoso.com/meeting?id=123456789" target="_blank">Join Contoso meeting</a>' +
        '<br><br>' +
        'Phone Dial-in: +1(123)456-7890' +
        '<br><br>' +
        'Meeting ID: 123 456 789' +
        '<br><br>' +
        'Want to test your video connection?' +
        '<br><br>' +
        '<a href="https://contoso.com/testmeeting" target="_blank">Join test meeting</a>' +
        '<br><br>';

    let mailboxItem;

    // Office is ready.
    Office.onReady(function () {
            mailboxItem = Office.context.mailbox.item;
        }
    );

    // 2. How to define and register a function command named `insertContosoMeeting` (referenced in the manifest)
    //    to update the meeting body with the online meeting details.
    function insertContosoMeeting(event) {
        // Get HTML body from the client.
        mailboxItem.body.getAsync("html",
            { asyncContext: event },
            function (getBodyResult) {
                if (getBodyResult.status === Office.AsyncResultStatus.Succeeded) {
                    updateBody(getBodyResult.asyncContext, getBodyResult.value);
                } else {
                    console.error("Failed to get HTML body.");
                    getBodyResult.asyncContext.completed({ allowEvent: false });
                }
            }
        );
    }
    // Register the function.
    Office.actions.associate("insertContosoMeeting", insertContosoMeeting);

    // 3. How to implement a supporting function `updateBody`
    //    that appends the online meeting details to the current body of the meeting.
    function updateBody(event, existingBody) {
        // Append new body to the existing body.
        mailboxItem.body.setAsync(existingBody + newBody,
            { asyncContext: event, coercionType: "html" },
            function (setBodyResult) {
                if (setBodyResult.status === Office.AsyncResultStatus.Succeeded) {
                    setBodyResult.asyncContext.completed({ allowEvent: true });
                } else {
                    console.error("Failed to set HTML body.");
                    setBodyResult.asyncContext.completed({ allowEvent: false });
                }
            }
        );
    }
    ```

## <a name="testing-and-validation"></a>测试和验证

按照常规指南[测试和验证加载项](testing-and-tips.md)，然后在 Outlook 网页版、Windows 或 Mac 中[旁加载](sideload-outlook-add-ins-for-testing.md)清单。 如果加载项还支持移动设备，请在旁加载后在 Android 或 iOS 设备上重启 Outlook。 旁加载加载项后，创建一个新会议，并验证是否已将 Microsoft Teams 或 Skype 切换开关替换为你自己的。

### <a name="create-meeting-ui"></a>创建会议 UI

作为会议组织者，在创建会议时，应会看到类似于以下三个图像的屏幕。

[![Android 上的“创建会议”屏幕，其中“Contoso”开关处于关闭状态。](../images/outlook-android-create-online-meeting-off.png)](../images/outlook-android-create-online-meeting-off-expanded.png#lightbox) [![Android 上的“创建会议”屏幕，其中包含“正在加载 Contoso”开关。](../images/outlook-android-create-online-meeting-load.png)](../images/outlook-android-create-online-meeting-load-expanded.png#lightbox) [![Android 上的“创建会议”屏幕，其中“Contoso”开关处于打开状态。](../images/outlook-android-create-online-meeting-on.png)](../images/outlook-android-create-online-meeting-on-expanded.png#lightbox)

### <a name="join-meeting-ui"></a>加入会议 UI

作为会议与会者，在查看会议时，应会看到类似于下图的屏幕。

[![Android 上的“加入会议”屏幕。](../images/outlook-android-join-online-meeting-view-1.png)](../images/outlook-android-join-online-meeting-view-1-expanded.png#lightbox)

> [!IMPORTANT]
> “**加入**”按钮仅在 Outlook 网页版、Mac、Android 和 iOS 中受支持。 如果只看到会议链接，但在受支持的客户端中看不到“ **加入** ”按钮，则可能是服务的在线会议模板未在我们的服务器上注册。 有关详细信息，请参阅 [注册联机会议模板](#register-your-online-meeting-template) 部分。

## <a name="register-your-online-meeting-template"></a>注册联机会议模板

注册联机会议加载项是可选的。 仅当你想要在会议中显示“ **加入** ”按钮以及会议链接时，它才适用。 开发联机会议加载项并想要注册它后，请使用以下指南创建 GitHub 问题。 我们将与你联系以协调注册时间线。

> [!IMPORTANT]
> “**加入**”按钮仅在 Outlook 网页版、Mac、Android 和 iOS 中受支持。

1. [创建新的 GitHub 问题](https://github.com/OfficeDev/office-js/issues/new)。
1. 将新问题的 **标题** 设置为“Outlook：为 my-service 注册联机会议模板”，并将 `my-service` 替换为服务名称。
1. 在问题正文中，将现有文本替换为在本文前面[实现添加联机会议详细信息](#implement-adding-online-meeting-details)部分中的 或类似变量中设置`newBody`的字符串。
1. 单击“ **提交新问题**”。

![包含 Contoso 示例内容的新 GitHub 问题屏幕。](../images/outlook-request-to-register-online-meeting-template.png)

## <a name="available-apis"></a>可用 API

以下 API 可用于此功能。

- 约会组织者 API
  - [Office.context.mailbox.item.body](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#outlook-office-appointmentcompose-body-member) ([Body.getAsync](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#outlook-office-body-getasync-member(1))、 [Body.setAsync](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#outlook-office-body-setasync-member(1))) 
  - [Office.context.mailbox.item.end](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#outlook-office-appointmentcompose-end-member) ([Time](/javascript/api/outlook/office.time?view=outlook-js-preview&preserve-view=true)) 
  - [Office.context.mailbox.item.loadCustomPropertiesAsync](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#outlook-office-appointmentcompose-loadcustompropertiesasync-member(1)) ([CustomProperties](/javascript/api/outlook/office.customproperties?view=outlook-js-preview&preserve-view=true)) 
  - [Office.context.mailbox.item.location](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#outlook-office-appointmentcompose-location-member) ([Location](/javascript/api/outlook/office.location?view=outlook-js-preview&preserve-view=true)) 
  - [Office.context.mailbox.item.optionalAttendees](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#outlook-office-appointmentcompose-optionalattendees-member) ([收件人](/javascript/api/outlook/office.recipients?view=outlook-js-preview&preserve-view=true)) 
  - [Office.context.mailbox.item.requiredAttendees](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#outlook-office-appointmentcompose-requiredattendees-member) ([收件人](/javascript/api/outlook/office.recipients?view=outlook-js-preview&preserve-view=true)) 
  - [Office.context.mailbox.item.start](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#outlook-office-appointmentcompose-start-member) ([Time](/javascript/api/outlook/office.time?view=outlook-js-preview&preserve-view=true)) 
  - [Office.context.mailbox.item.subject](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#outlook-office-appointmentcompose-subject-member) ([Subject](/javascript/api/outlook/office.subject?view=outlook-js-preview&preserve-view=true)) 
  - [Office.context.roamingSettings](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context?view=outlook-js-preview&preserve-view=true#roamingsettings-roamingsettings) ([RoamingSettings](/javascript/api/outlook/office.roamingsettings?view=outlook-js-preview&preserve-view=true)) 
- 处理身份验证流
  - [Dialog API](../develop/dialog-api-in-office-add-ins.md)

## <a name="restrictions"></a>限制

存在一些限制。

- 仅适用于联机会议服务提供商。
- 只有管理员安装的加载项会显示在会议撰写屏幕上，并替换默认的 Teams 或 Skype 选项。 用户安装的加载项不会激活。
- 加载项图标应采用灰度，使用十六进制代码 `#919191` 或其等效的其他 [颜色格式](https://convertingcolors.com/hex-color-919191.html)。
- 约会组织者 (撰写) 模式中仅支持一个函数命令。
- 加载项应在一分钟的超时期限内更新约会表单中的会议详细信息。 但是，例如，在打开加载项进行身份验证的对话框中花费的任何时间都排除在超时期限之外。

## <a name="see-also"></a>另请参阅

- [适用于 Outlook Mobile 的加载项](outlook-mobile-addins.md)
- [添加对适用于 Outlook Mobile 的外接程序命令的支持](add-mobile-support.md)
