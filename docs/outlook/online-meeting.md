---
title: 为联机会议提供商创建 Outlook 移动外接程序
description: 讨论如何为联机会议服务提供商设置 Outlook 移动外接程序。
ms.topic: article
ms.date: 02/12/2021
localization_priority: Normal
ms.openlocfilehash: b973a0cada4127ecc614d42764a9ecea2a00fa2c
ms.sourcegitcommit: e7009c565b18c607fe0868db2e26e250ad308dce
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/05/2021
ms.locfileid: "50505519"
---
# <a name="create-an-outlook-mobile-add-in-for-an-online-meeting-provider"></a>为联机会议提供商创建 Outlook 移动外接程序

设置联机会议是 Outlook 用户的核心体验，使用 Outlook 移动版创建 [Teams](/microsoftteams/teams-add-in-for-outlook) 会议很容易。 但是，使用非 Microsoft 服务在 Outlook 中创建联机会议可能很麻烦。 通过实现此功能，服务提供商可以简化其 Outlook 外接程序用户的联机会议创建体验。

> [!IMPORTANT]
> 此功能仅在具有 Microsoft 365 订阅的 Android 和 iOS 上受支持。

本文将了解如何设置 Outlook 移动外接程序，使用户能够使用联机会议服务组织和加入会议。 在本文中，我们将使用虚构的在线会议服务提供商"Contoso"。

## <a name="set-up-your-environment"></a>设置环境

使用 [适用于 Office](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) 加载项的 Yeoman 生成器完成创建外接程序项目的 Outlook 快速入门。

## <a name="configure-the-manifest"></a>配置清单

若要使用户能够使用外接程序创建联机会议，您必须在父元素下的清单中配置 [MobileOnlineMeetingCommandSurface](../reference/manifest/extensionpoint.md#mobileonlinemeetingcommandsurface) 扩展点 `MobileFormFactor` 。 不支持其他外形因素。

1. 在代码编辑器中，打开快速启动项目。

1. 打开 **manifest.xml** 根目录下的文件。

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
> 若要了解有关 Outlook 外接程序清单的更多信息，请参阅 [Outlook](manifests.md) 外接程序清单和添加对 [Outlook Mobile](add-mobile-support.md)外接程序命令的支持。

## <a name="implement-adding-online-meeting-details"></a>实现添加联机会议详细信息

在此部分中，了解外接程序脚本如何更新用户的会议以包含联机会议详细信息。

1. 从同一快速启动项目中，在代码编辑器中commands.js **./src/commands/commands.js** 文件。

1. 用以下 JavaScript **commands.js** 文件的全部内容。

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

    // 2. How to define a UI-less function named `insertContosoMeeting` (referenced in the manifest)
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

    function getGlobal() {
      return typeof self !== "undefined"
        ? self
        : typeof window !== "undefined"
        ? window
        : typeof global !== "undefined"
        ? global
        : undefined;
    }

    const g = getGlobal();

    // The add-in command functions need to be available in global scope.
    g.insertContosoMeeting = insertContosoMeeting;
    ```

## <a name="testing-and-validation"></a>测试和验证

按照常规指南 [测试和验证加载项](testing-and-tips.md)。 在 [Outlook 网页版](sideload-outlook-add-ins-for-testing.md) 、Windows 版或 Mac 版 Outlook 中旁加载后，在 Android 移动设备上重新启动 Outlook。  (Android 现在是唯一受支持的客户端。) 然后，在新的会议屏幕上，验证 Microsoft Teams 或 Skype 切换键是否替换为你自己的。

### <a name="create-meeting-ui"></a>创建会议 UI

作为会议组织者，您应该在创建会议时看到类似于以下三个图像的屏幕。

Android 上的创建会议屏幕的屏幕截图[ ![ - Contoso 在](../images/outlook-android-create-online-meeting-off.png)](../images/outlook-android-create-online-meeting-off-expanded.png#lightbox)Android 上关闭创建会议屏幕的屏幕截图 - 在 Android 上加载创建会议屏幕的[ ![ Contoso](../images/outlook-android-create-online-meeting-load.png)](../images/outlook-android-create-online-meeting-load-expanded.png#lightbox)切换屏幕截图 - [ ![ Contoso 切换打开](../images/outlook-android-create-online-meeting-on.png)](../images/outlook-android-create-online-meeting-on-expanded.png#lightbox)

### <a name="join-meeting-ui"></a>加入会议 UI

作为与会者，在查看会议时应该会看到类似于下图的屏幕。

[![Android 上的加入会议屏幕的屏幕截图](../images/outlook-android-join-online-meeting-view-1.png)](../images/outlook-android-join-online-meeting-view-1-expanded.png#lightbox)

> [!IMPORTANT]
> 如果看不到"加入 **"链接，** 可能是你的服务的联机会议模板未在我们的服务器上注册。 有关详细信息 [，请参阅"注册联机会议模板](#register-your-online-meeting-template) "部分。

## <a name="register-your-online-meeting-template"></a>注册联机会议模板

如果要为服务注册联机会议模板，可以创建包含详细信息的 GitHub 问题。 之后，我们将联系你以协调注册时间线。

1. 转到 **本文** 末尾的"反馈"部分。
1. 按 **"此页面"** 链接。
1. 将 **新问题** 的标题设置为"为 my-service 注册联机会议模板"，替换为 `my-service` 你的服务名称。
1. 在问题正文中，将字符串"[在此处输入反馈]"替换为你在本文前面"实现添加联机会议详细信息"部分或类似变量中设置的 `newBody` 字符串。 [](#implement-adding-online-meeting-details)
1. 单击 **"提交新问题"。**

![包含 Contoso 示例内容的新 GitHub 问题屏幕的屏幕截图](../images/outlook-request-to-register-online-meeting-template.png)

## <a name="available-apis"></a>可用的 API

以下 API 可用于此功能。

- 约会组织者 API
  - [Office.context.mailbox.item.body](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#body) ([Body.getAsync](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#getasync-coerciontype--options--callback-)， [Body.setAsync](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#setasync-data--options--callback-)) 
  - [Office.context.mailbox.item.end (](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#end) [Time](/javascript/api/outlook/office.time?view=outlook-js-preview&preserve-view=true)) 
  - [Office.context.mailbox.item.loadCustomPropertiesAsync](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#loadcustompropertiesasync-callback--usercontext-) ([CustomProperties](/javascript/api/outlook/office.customproperties?view=outlook-js-preview&preserve-view=true)) 
  - [Office.context.mailbox.item.location (](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#location) [Location) ](/javascript/api/outlook/office.location?view=outlook-js-preview&preserve-view=true)
  - [Office.context.mailbox.item.optionalAttendees](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#optionalattendees) ([Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-preview&preserve-view=true)) 
  - [Office.context.mailbox.item.requiredAttendees](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#requiredattendees) ([Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-preview&preserve-view=true)) 
  - [Office.context.mailbox.item.start (](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#start) [Time) ](/javascript/api/outlook/office.time?view=outlook-js-preview&preserve-view=true)
  - [Office.context.mailbox.item.subject (](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#subject) [Subject) ](/javascript/api/outlook/office.subject?view=outlook-js-preview&preserve-view=true)
  - [Office.context.roamingSettings](../reference/objectmodel/preview-requirement-set/office.context.md?view=outlook-js-preview&preserve-view=true#roamingsettings-roamingsettings) ([RoamingSettings) ](/javascript/api/outlook/office.roamingsettings?view=outlook-js-preview&preserve-view=true)
- 处理身份验证流
  - [Dialog API](../develop/dialog-api-in-office-add-ins.md)

## <a name="restrictions"></a>限制

有几个限制适用。

- 仅适用于联机会议服务提供商。
- 只有管理员安装的加载项将显示在会议撰写屏幕上，以替换默认的 Teams 或 Skype 选项。 用户安装的加载项不会激活。
- 外接程序图标应为使用十六进制代码或其他颜色格式的等效项 `#919191` 的 [灰度](https://convertingcolors.com/hex-color-919191.html)。
- 在"约会管理器"模式下，只有一个无 UI 命令 (撰写) 模式。

## <a name="see-also"></a>另请参阅

- [适用于 Outlook Mobile 的加载项](outlook-mobile-addins.md)
- [添加对适用于 Outlook Mobile 的外接程序命令的支持](add-mobile-support.md)
