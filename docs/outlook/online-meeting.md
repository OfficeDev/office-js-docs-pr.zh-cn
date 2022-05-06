---
title: 为联机会议提供商创建Outlook移动加载项
description: 讨论如何为联机会议服务提供商设置Outlook移动加载项。
ms.topic: article
ms.date: 05/01/2022
ms.localizationpriority: medium
ms.openlocfilehash: ada539c323ce6fdf578c26c221eb6090280b5a0a
ms.sourcegitcommit: 5773c76912cdb6f0c07a932ccf07fc97939f6aa1
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/06/2022
ms.locfileid: "65244861"
---
# <a name="create-an-outlook-mobile-add-in-for-an-online-meeting-provider"></a>为联机会议提供商创建Outlook移动加载项

对于Outlook用户来说，设置联机会议是一种核心体验，并且可以轻松[地使用Outlook移动设备创建Teams会议](/microsoftteams/teams-add-in-for-outlook)。 但是，使用非 Microsoft 服务在Outlook中创建联机会议可能很麻烦。 通过实现此功能，服务提供商可以简化其Outlook外接程序用户的联机会议创建体验。

> [!IMPORTANT]
> 此功能仅在具有Microsoft 365订阅的 Android 和 iOS 上受支持。

本文介绍如何设置Outlook移动外接程序，使用户能够使用联机会议服务组织和加入会议。 在本文中，我们将使用虚构的联机会议服务提供商“Contoso”。

## <a name="set-up-your-environment"></a>设置环境

完成[Outlook快速入](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator)门，使用 yeoman 生成器为Office加载项创建加载项项目。

## <a name="configure-the-manifest"></a>配置清单

若要使用户能够使用外接程序创建联机会议，必须在父元素`MobileFormFactor`下的清单中配置 [MobileOnlineMeetingCommandSurface 扩展点](/javascript/api/manifest/extensionpoint#mobileonlinemeetingcommandsurface)。 不支持其他外形因素。

1. 在代码编辑器中，打开快速启动项目。

1. 打开位于项目根 **目录的manifest.xml** 文件。

1. 选择整个 `<VersionOverrides>` 节点 (包括打开和关闭标记) 并将其替换为以下 XML。

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

> [!TIP]
> 若要详细了解Outlook加载项的清单，请[参阅Outlook加载项清单](manifests.md)和[添加对 Outlook Mobile 加载项命令的支持](add-mobile-support.md)。

## <a name="implement-adding-online-meeting-details"></a>实现添加联机会议详细信息

在本部分中，了解外接程序脚本如何更新用户的会议以包含联机会议详细信息。

1. 在同一快速入门项目中，在代码编辑器中打开文件 **./src/commands/commands.js** 。

1. 将 **commands.js** 文件的整个内容替换为以下 JavaScript。

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

    var mailboxItem;

    // Office is ready.
    Office.onReady(function () {
            mailboxItem = Office.context.mailbox.item;
        }
    );

    // 2. How to define and register a UI-less function named `insertContosoMeeting` (referenced in the manifest)
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

按照通常的指南 [测试和验证加载项](testing-and-tips.md)。 在 Outlook 网页版、Windows 或 Mac 中[旁加载](sideload-outlook-add-ins-for-testing.md)后，在 Android 或 iOS 移动设备上重启Outlook。 然后，在新的会议屏幕上，验证Microsoft Teams或Skype切换是否已替换为自己的切换。

### <a name="create-meeting-ui"></a>创建会议 UI

作为会议组织者，在创建会议时，应会看到类似于以下三个图像的屏幕。

[![Android 上的“创建会议”屏幕 - Contoso 关闭。](../images/outlook-android-create-online-meeting-off.png)](../images/outlook-android-create-online-meeting-off-expanded.png#lightbox) [![Android 上的创建会议屏幕 - 加载 Contoso 切换。](../images/outlook-android-create-online-meeting-load.png)](../images/outlook-android-create-online-meeting-load-expanded.png#lightbox) [![Android 上的创建会议屏幕 - Contoso 切换。](../images/outlook-android-create-online-meeting-on.png)](../images/outlook-android-create-online-meeting-on-expanded.png#lightbox)

### <a name="join-meeting-ui"></a>加入会议 UI

作为会议与会者，在查看会议时，应会看到类似于下图的屏幕。

[![Android 上加入会议屏幕的屏幕截图。](../images/outlook-android-join-online-meeting-view-1.png)](../images/outlook-android-join-online-meeting-view-1-expanded.png#lightbox)

> [!IMPORTANT]
> 如果未看到 **“加入** ”链接，则可能是服务的联机会议模板未在我们的服务器上注册。 有关详细信息，请参阅 [“注册联机会议模板”](#register-your-online-meeting-template) 部分。

## <a name="register-your-online-meeting-template"></a>注册联机会议模板

若要为服务注册联机会议模板，可以创建GitHub详细信息问题。 之后，我们将与你联系以协调注册时间线。

1. 转到本文末尾的 **“反馈** ”部分。
1. 按 **此页** 链接。
1. 将新问题的 **标题** 设置为“注册我的服务的联机会议模板”，替换 `my-service` 为服务名称。
1. 在问题正文中，将字符串“[[在此处输入反馈]”替换为在本文前面的“[实现添加联机会议详细信息](#implement-adding-online-meeting-details)”部分中`newBody`设置的字符串或类似变量中设置的字符串。
1. 单击 **“提交新问题**”。

![包含 Contoso 示例内容的新GitHub问题屏幕的屏幕截图。](../images/outlook-request-to-register-online-meeting-template.png)

## <a name="available-apis"></a>可用 API

以下 API 可用于此功能。

- 约会组织者 API
  - [Office.context.mailbox.item.body](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#outlook-office-appointmentcompose-body-member) ([Body.getAsync](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#outlook-office-body-getasync-member(1))、[Body.setAsync](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#outlook-office-body-setasync-member(1))) 
  - [Office.context.mailbox.item.end](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#outlook-office-appointmentcompose-end-member) ([时间](/javascript/api/outlook/office.time?view=outlook-js-preview&preserve-view=true)) 
  - [Office.context.mailbox.item.loadCustomPropertiesAsync](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#outlook-office-appointmentcompose-loadcustompropertiesasync-member(1)) ([CustomProperties](/javascript/api/outlook/office.customproperties?view=outlook-js-preview&preserve-view=true)) 
  - [Office.context.mailbox.item.location](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#outlook-office-appointmentcompose-location-member) ([位置](/javascript/api/outlook/office.location?view=outlook-js-preview&preserve-view=true)) 
  - [Office.context.mailbox.item.optionalAttendees](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#outlook-office-appointmentcompose-optionalattendees-member) ([收件人](/javascript/api/outlook/office.recipients?view=outlook-js-preview&preserve-view=true)) 
  - [Office.context.mailbox.item.requiredAttendees](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#outlook-office-appointmentcompose-requiredattendees-member) ([收件人](/javascript/api/outlook/office.recipients?view=outlook-js-preview&preserve-view=true)) 
  - [Office.context.mailbox.item.start](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#outlook-office-appointmentcompose-start-member) ([时间](/javascript/api/outlook/office.time?view=outlook-js-preview&preserve-view=true)) 
  - [Office.context.mailbox.item.subject](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#outlook-office-appointmentcompose-subject-member) ([主题](/javascript/api/outlook/office.subject?view=outlook-js-preview&preserve-view=true)) 
  - [Office.context.roamingSettings](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context?view=outlook-js-preview&preserve-view=true#roamingsettings-roamingsettings) ([RoamingSettings](/javascript/api/outlook/office.roamingsettings?view=outlook-js-preview&preserve-view=true)) 
- 处理身份验证流
  - [Dialog API](../develop/dialog-api-in-office-add-ins.md)

## <a name="restrictions"></a>限制

有几个限制适用。

- 仅适用于联机会议服务提供商。
- 会议撰写屏幕上将仅显示管理员安装的加载项，替换默认Teams或Skype选项。 用户安装的加载项不会激活。
- 外接程序图标应使用十六进制代码 `#919191` 或 [以其他颜色格式](https://convertingcolors.com/hex-color-919191.html)等效的灰度。
- 约会组织者 (撰写) 模式中仅支持一个无 UI 命令。
- 加载项应在一分钟的超时时间内更新约会表单中的会议详细信息。 但是，为身份验证打开的加载项等在对话框中花费的任何时间都排除在超时时间段之外。

## <a name="see-also"></a>另请参阅

- [适用于 Outlook Mobile 的加载项](outlook-mobile-addins.md)
- [添加对适用于 Outlook Mobile 的外接程序命令的支持](add-mobile-support.md)
