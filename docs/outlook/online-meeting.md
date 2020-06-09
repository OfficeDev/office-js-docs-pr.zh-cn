---
title: 为联机会议提供商创建 Outlook mobile 外接程序
description: 讨论如何为联机会议服务提供商设置 Outlook 移动外接程序。
ms.topic: article
ms.date: 05/19/2020
localization_priority: Normal
ms.openlocfilehash: d35aa1ecd2b03b51314b5e88ae08c7fcb8382817
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/08/2020
ms.locfileid: "44609032"
---
# <a name="create-an-outlook-mobile-add-in-for-an-online-meeting-provider"></a>为联机会议提供商创建 Outlook mobile 外接程序

设置联机会议是 Outlook 用户的核心体验，可轻松[创建使用 outlook mobile 的团队会议](/microsoftteams/teams-add-in-for-outlook)。 但是，在 Outlook 中使用非 Microsoft 服务创建联机会议可能很麻烦。 通过实施此功能，服务提供商可以为其 Outlook 外接程序用户简化联机会议创建体验。

> [!IMPORTANT]
> 此功能仅在 Office 365 订阅的 Android 上受支持。

在本文中，您将了解如何设置 Outlook 移动外接程序，以使用户能够使用您的联机会议服务来组织和加入会议。 在整篇文章中，我们将使用虚构的联机会议服务提供商 "Contoso"。

## <a name="configure-the-manifest"></a>配置清单

若要使用户能够使用您的外接程序创建联机会议，必须在 `MobileOnlineMeetingCommandSurface` 父元素下的清单中配置扩展点 `MobileFormFactor` 。 不支持其他外观因素。

下面的示例展示了包含 `MobileFormFactor` 元素和扩展点的清单中的摘录 `MobileOnlineMeetingCommandSurface` 。

> [!TIP]
> 若要了解有关 Outlook 外接程序清单的详细信息，请参阅[outlook 外接程序清单](manifests.md)和[添加对适用于 outlook Mobile 的外接程序命令的支持](add-mobile-support.md)。

```xml
...
<VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
    ...
    <Hosts>
      <Host xsi:type="MailHost">
        <MobileFormFactor>
          <FunctionFile resid="residMobileFuncUrl" />
          <ExtensionPoint xsi:type="MobileOnlineMeetingCommandSurface">
            <!-- Configure selected extension point. -->
            <Control xsi:type="MobileButton" id="onlineMeetingFunctionButton">
              <Label resid="residUILessButton0Name" />
              <Icon>
                <bt:Image resid="UiLessIcon" size="25" scale="1" />
                <bt:Image resid="UiLessIcon" size="25" scale="2" />
                <bt:Image resid="UiLessIcon" size="25" scale="3" />
                <bt:Image resid="UiLessIcon" size="32" scale="1" />
                <bt:Image resid="UiLessIcon" size="32" scale="2" />
                <bt:Image resid="UiLessIcon" size="32" scale="3" />
                <bt:Image resid="UiLessIcon" size="48" scale="1" />
                <bt:Image resid="UiLessIcon" size="48" scale="2" />
                <bt:Image resid="UiLessIcon" size="48" scale="3" />
              </Icon>
              <Action xsi:type="ExecuteFunction">
                <FunctionName>insertContosoMeeting</FunctionName>
              </Action>
            </Control>
          </ExtensionPoint>
        </MobileFormFactor>
      </Host>
    </Hosts>
    ...
  </VersionOverrides>
</VersionOverrides>
...
```

## <a name="implement-adding-online-meeting-details"></a>实施添加联机会议详细信息

在本节中，了解外接程序脚本如何更新用户的会议以包含联机会议详细信息。

下面的示例展示了如何构建联机会议详细信息。 不显示：如何从服务中获取会议组织者的 ID 和其他详细信息。

```js
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
```

下面的示例演示如何定义清单中引用的无 UI 的函数， `insertContosoMeeting` 以使用联机会议详细信息更新会议正文。

```js
var mailboxItem;

// Office is ready.
Office.onReady(function () {
        mailboxItem = Office.context.mailbox.item;
    }
);

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
```

下面的示例演示 `updateBody` 在上一示例中使用的支持函数的实现，该示例将联机会议详细信息追加到会议的当前正文。

```js
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

按照通常的指导来[测试和验证您的外接程序](testing-and-tips.md)。 在 Outlook 网页版、Windows 版或 Mac 版中进行[旁加载](sideload-outlook-add-ins-for-testing.md)后，在你的 android 移动设备上重新启动 outlook （现在，android 是唯一受支持的客户端）。 然后，在新的会议屏幕上，验证 Microsoft 团队或 Skype 切换是否已替换为你自己的。

### <a name="create-meeting-ui"></a>创建会议用户界面

作为会议组织者，在创建会议时，您应看到类似于以下三幅图像的屏幕。

在 android 上创建会议屏幕的[ ![ 屏幕截图-contoso 切换关闭](../images/outlook-android-create-online-meeting-off.png)](../images/outlook-android-create-online-meeting-off-expanded.png#lightbox)在 android 上创建会议屏幕[ ![ 的屏幕截图](../images/outlook-android-create-online-meeting-load.png)](../images/outlook-android-create-online-meeting-load-expanded.png#lightbox)打开在 android 上创建会议屏幕的屏幕截图[ ![ -contoso 切换](../images/outlook-android-create-online-meeting-on.png)](../images/outlook-android-create-online-meeting-on-expanded.png#lightbox)

### <a name="join-meeting-ui"></a>加入会议 UI

作为会议与会者，查看会议时，您应该会看到类似于下图的屏幕。

[![Android 上的加入会议屏幕的屏幕截图](../images/outlook-android-join-online-meeting-view-1.png)](../images/outlook-android-join-online-meeting-view-1-expanded.png#lightbox)

## <a name="available-apis"></a>可用 Api

以下 Api 可用于此功能。

- 约会组织者 Api
  - " [Context.subname](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#subject) " （[subject](/javascript/api/outlook/office.subject?view=outlook-js-preview)）
  - "Context.subname" （Time）。[开始](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#start)（[Time](/javascript/api/outlook/office.time?view=outlook-js-preview)）
  - （Time）[结尾](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#end)（[Time](/javascript/api/outlook/office.time?view=outlook-js-preview)）
  - " [Context.subname](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#location) " （"[位置](/javascript/api/outlook/office.location?view=outlook-js-preview)"）
  - [OptionalAttendees](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#optionalattendees) （[收件人](/javascript/api/outlook/office.recipients?view=outlook-js-preview)）的内容
  - [RequiredAttendees](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#requiredattendees) （[收件人](/javascript/api/outlook/office.recipients?view=outlook-js-preview)）的内容
  - "[GetAsync](/javascript/api/outlook/office.body?view=outlook-js-preview#getasync-coerciontype--options--callback-) [" （"."](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#body) 、"setAsync"、"Body"、" [Body.setAsync](/javascript/api/outlook/office.body?view=outlook-js-preview#setasync-data--options--callback-)"）
  - [LoadCustomPropertiesAsync](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#loadcustompropertiesasync-callback--usercontext-) （[CustomProperties](/javascript/api/outlook/office.customproperties?view=outlook-js-preview)）的内容
  - [RoamingSettings](../reference/objectmodel/preview-requirement-set/office.context.md?view=outlook-js-preview#roamingsettings-roamingsettings) （[roamingSettings](/javascript/api/outlook/office.roamingsettings?view=outlook-js-preview)）
- 处理身份验证流
  - [Dialog API](../develop/dialog-api-in-office-add-ins.md)

## <a name="restrictions"></a>条件

应用了多个限制。

- 仅适用于联机会议服务提供商。
- 目前，Android 是唯一受支持的客户端。 即将推出对 iOS 的支持。
- 只有管理员安装的加载项才会显示在会议撰写屏幕上，替换默认团队或 Skype 选项。 无法激活用户安装的外接程序。
- 外接端图标应使用十六进制代码 `#919191` 或以[其他颜色格式](https://convertingcolors.com/hex-color-919191.html)的等效项进行灰度。
- 在约会组织者（撰写）模式下仅支持一个无 UI 的命令。

## <a name="see-also"></a>另请参阅

- [适用于 Outlook Mobile 的加载项](outlook-mobile-addins.md)
- [添加对适用于 Outlook Mobile 的外接程序命令的支持](add-mobile-support.md)
