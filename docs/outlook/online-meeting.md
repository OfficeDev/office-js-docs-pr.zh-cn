---
title: 为联机会议提供商创建 Outlook mobile 外接程序
description: 讨论如何为联机会议服务提供商设置 Outlook 移动外接程序。
ms.topic: article
ms.date: 06/25/2020
localization_priority: Normal
ms.openlocfilehash: 9f0b50602ab4941b16c15abe97c3f099a54f5b42
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/10/2020
ms.locfileid: "45093999"
---
# <a name="create-an-outlook-mobile-add-in-for-an-online-meeting-provider"></a>为联机会议提供商创建 Outlook mobile 外接程序

设置联机会议是 Outlook 用户的核心体验，可轻松[创建使用 outlook mobile 的团队会议](/microsoftteams/teams-add-in-for-outlook)。 但是，在 Outlook 中使用非 Microsoft 服务创建联机会议可能很麻烦。 通过实施此功能，服务提供商可以为其 Outlook 外接程序用户简化联机会议创建体验。

> [!IMPORTANT]
> 仅适用于使用 Microsoft 365 订阅的 Android 支持此功能。

在本文中，您将了解如何设置 Outlook 移动外接程序，以使用户能够使用您的联机会议服务来组织和加入会议。 在整篇文章中，我们将使用虚构的联机会议服务提供商 "Contoso"。

## <a name="set-up-your-environment"></a>设置环境

完成[Outlook 快速入门](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator)，它将使用 Office 外接程序的 Yeoman 生成器创建外接程序项目。

## <a name="configure-the-manifest"></a>配置清单

若要使用户能够使用您的外接程序创建联机会议，必须在 `MobileOnlineMeetingCommandSurface` 父元素下的清单中配置扩展点 `MobileFormFactor` 。 不支持其他外观因素。

1. 在代码编辑器中，打开 "快速启动" 项目。

1. 打开位于项目根目录中的**manifest.xml**文件。

1. 选择整个 `<VersionOverrides>` 节点 (包括 "打开" 和 "关闭" 标记) 并将其替换为以下 XML。

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
> 若要了解有关 Outlook 外接程序清单的详细信息，请参阅[outlook 外接程序清单](manifests.md)和[添加对适用于 outlook Mobile 的外接程序命令的支持](add-mobile-support.md)。

## <a name="implement-adding-online-meeting-details"></a>实施添加联机会议详细信息

在本节中，了解外接程序脚本如何更新用户的会议以包含联机会议详细信息。

1. 在同一 "快速启动" 项目中，在代码编辑器中打开 **/src/commands/commands.js** 。

1. 将**commands.js**文件的整个内容替换为以下 JavaScript。

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

按照通常的指导来[测试和验证您的外接程序](testing-and-tips.md)。 在 Outlook 网页版、Windows 版或 Mac 版中进行[旁加载](sideload-outlook-add-ins-for-testing.md)后，在你的 Android 移动设备上重新启动 outlook。  (Android 是目前唯一受支持的客户端。 ) 然后，在新的会议屏幕上，验证 Microsoft 团队或 Skype 切换是否已替换为您自己的。

### <a name="create-meeting-ui"></a>创建会议用户界面

作为会议组织者，在创建会议时，您应看到类似于以下三幅图像的屏幕。

在 android 上创建会议屏幕的[ ![ 屏幕截图-contoso 切换关闭](../images/outlook-android-create-online-meeting-off.png)](../images/outlook-android-create-online-meeting-off-expanded.png#lightbox)在 android 上创建会议屏幕[ ![ 的屏幕截图](../images/outlook-android-create-online-meeting-load.png)](../images/outlook-android-create-online-meeting-load-expanded.png#lightbox)打开在 android 上创建会议屏幕的屏幕截图[ ![ -contoso 切换](../images/outlook-android-create-online-meeting-on.png)](../images/outlook-android-create-online-meeting-on-expanded.png#lightbox)

### <a name="join-meeting-ui"></a>加入会议 UI

作为会议与会者，查看会议时，您应该会看到类似于下图的屏幕。

[![Android 上的加入会议屏幕的屏幕截图](../images/outlook-android-join-online-meeting-view-1.png)](../images/outlook-android-join-online-meeting-view-1-expanded.png#lightbox)

> [!IMPORTANT]
> 如果看不到**Join**链接，则可能是你的服务的联机会议模板未在我们的服务器上注册。 有关详细信息，请参阅[注册联机会议模板](#register-your-online-meeting-template)部分。

## <a name="register-your-online-meeting-template"></a>注册您的联机会议模板

如果您想要为服务注册联机会议模板，则可以使用详细信息创建 GitHub 问题。 之后，我们将与您联系以协调注册日程表。

1. 请转到本文结尾处的 "**反馈**" 部分。
1. 按 "**此页面**" 链接。
1. 将新问题的**标题**设置为 "为我的服务注册联机会议模板"，并将其替换 `my-service` 为您的服务名称。
1. 在问题正文中，将字符串 "[输入反馈此处]" 替换为您在 `newBody` 本文前面的 "[实现添加联机会议详细信息" 部分中的 "实现添加联机会议详细信息](#implement-adding-online-meeting-details)" 部分中设置的字符串。
1. 单击 "**提交新问题**"。

![包含 Contoso 示例内容的新 GitHub 问题屏幕的屏幕截图](../images/outlook-request-to-register-online-meeting-template.png)

## <a name="available-apis"></a>可用 Api

以下 Api 可用于此功能。

- 约会组织者 Api
  - [使用者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#subject) ([主题](/javascript/api/outlook/office.subject?view=outlook-js-preview)) 的主题
  -  (时间) [的开始](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#start)[时间](/javascript/api/outlook/office.time?view=outlook-js-preview)
  -  (时间) [的结束](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#end)[时间](/javascript/api/outlook/office.time?view=outlook-js-preview)
  -  ([位置](/javascript/api/outlook/office.location?view=outlook-js-preview)) [的位置。](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#location)
  - [OptionalAttendees](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#optionalattendees) ([收件人](/javascript/api/outlook/office.recipients?view=outlook-js-preview)) 中的
  - [RequiredAttendees](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#requiredattendees) ([收件人](/javascript/api/outlook/office.recipients?view=outlook-js-preview)) 中的
  -  ([setAsync](/javascript/api/outlook/office.body?view=outlook-js-preview#setasync-data--options--callback-)) 的 " [context.subname](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#body) [" 的 "](/javascript/api/outlook/office.body?view=outlook-js-preview#getasync-coerciontype--options--callback-)正文"。
  - [LoadCustomPropertiesAsync](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#loadcustompropertiesasync-callback--usercontext-) ([CustomProperties](/javascript/api/outlook/office.customproperties?view=outlook-js-preview)) 的
  - [RoamingSettings](../reference/objectmodel/preview-requirement-set/office.context.md?view=outlook-js-preview#roamingsettings-roamingsettings) ([roamingSettings](/javascript/api/outlook/office.roamingsettings?view=outlook-js-preview)) 
- 处理身份验证流
  - [Dialog API](../develop/dialog-api-in-office-add-ins.md)

## <a name="restrictions"></a>条件

应用了多个限制。

- 仅适用于联机会议服务提供商。
- 目前，Android 是唯一受支持的客户端。 即将推出对 iOS 的支持。
- 只有管理员安装的加载项才会显示在会议撰写屏幕上，替换默认团队或 Skype 选项。 无法激活用户安装的外接程序。
- 外接端图标应使用十六进制代码 `#919191` 或以[其他颜色格式](https://convertingcolors.com/hex-color-919191.html)的等效项进行灰度。
- 约会组织者 (撰写) 模式中仅支持一个无 UI 的命令。

## <a name="see-also"></a>另请参阅

- [适用于 Outlook Mobile 的加载项](outlook-mobile-addins.md)
- [添加对适用于 Outlook Mobile 的外接程序命令的支持](add-mobile-support.md)
