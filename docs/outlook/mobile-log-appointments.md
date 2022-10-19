---
title: 在 Outlook 移动外接程序中将约会说明记录到外部应用程序
description: 了解如何设置 Outlook 移动外接程序以将约会笔记和其他详细信息记录到外部应用程序。
ms.topic: article
ms.date: 10/17/2022
ms.localizationpriority: medium
ms.openlocfilehash: a980b68c603154c42112f525ec6285b740ce38a5
ms.sourcegitcommit: eca6c16d0bb74bed2d35a21723dd98c6b41ef507
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/18/2022
ms.locfileid: "68607581"
---
# <a name="log-appointment-notes-to-an-external-application-in-outlook-mobile-add-ins"></a>在 Outlook 移动外接程序中将约会说明记录到外部应用程序

将约会笔记和其他详细信息保存到客户关系管理 (CRM) 或记笔记应用程序可以帮助你跟踪已参加的会议。

本文介绍如何设置 Outlook 移动外接程序，使用户能够将笔记和其他有关其约会的详细信息记录到 CRM 或记笔记应用程序。 在本文中，我们将使用名为“Contoso”的虚构 CRM 服务提供程序。

> [!IMPORTANT]
> 此功能仅在具有 Microsoft 365 订阅的 Android 上受支持。

## <a name="set-up-your-environment"></a>设置环境

完成 [Outlook 快速入](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) 门，使用 Office 外接程序的 Yeoman 生成器创建加载项项目。

## <a name="capture-and-view-appointment-notes"></a>捕获和查看约会说明

可以选择实现函数命令或任务窗格。 若要更新加载项，请选择函数命令或任务窗格的选项卡，然后按照说明操作。

# <a name="function-command"></a>[函数命令](#tab/noui)

此选项将使用户能够在从功能区中选择函数命令时记录和查看其笔记及其约会的其他详细信息。

### <a name="configure-the-manifest"></a>配置清单

若要使用户能够使用外接程序记录约会笔记，必须在父元素`MobileFormFactor`下的清单中配置 [MobileLogEventAppointmentAttendee 扩展点](/javascript/api/manifest/extensionpoint#mobilelogeventappointmentattendee)。 不支持其他外形因素。

[!INCLUDE [Teams manifest not supported on mobile devices](../includes/no-mobile-with-json-note.md)]

1. 在代码编辑器中，打开快速启动项目。

1. 打开位于项目根 **目录的manifest.xml** 文件。

1. 选择整个 `<VersionOverrides>` 节点 (包括打开和关闭标记) 并将其替换为以下 XML。 请确保将 **对 Contoso** 的所有引用替换为公司信息。

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
              <ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">
                <OfficeTab id="TabDefault">
                  <Group id="apptReadGroup">
                    <Label resid="residDescription"/>
                    <Control xsi:type="Button" id="apptReadOpenPaneButton">
                      <Label resid="residLabel"/>
                      <Supertip>
                        <Title resid="residLabel"/>
                        <Description resid="residTooltip"/>
                      </Supertip>
                      <Icon>
                        <bt:Image size="16" resid="icon-16"/>
                        <bt:Image size="32" resid="icon-32"/>
                        <bt:Image size="80" resid="icon-80"/>
                      </Icon>
                      <Action xsi:type="ExecuteFunction">
                        <FunctionName>logCRMEvent</FunctionName>
                      </Action>
                    </Control>
                  </Group>
                </OfficeTab>
              </ExtensionPoint>
            </DesktopFormFactor>
            <MobileFormFactor>
              <FunctionFile resid="residFunctionFile"/>
              <ExtensionPoint xsi:type="MobileLogEventAppointmentAttendee">
                <Control xsi:type="MobileButton" id="appointmentReadFunctionButton">
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
                    <FunctionName>logCRMEvent</FunctionName>
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
            <bt:Image id="icon-80" DefaultValue="https://contoso.com/assets/icon-80.png"/>
          </bt:Images>
          <bt:Urls>
            <bt:Url id="residFunctionFile" DefaultValue="https://contoso.com/commands.html"/>
          </bt:Urls>
          <bt:ShortStrings>
            <bt:String id="residDescription" DefaultValue="Log appointment notes and other details to Contoso CRM."/>
            <bt:String id="residLabel" DefaultValue="Log to Contoso CRM"/>
          </bt:ShortStrings>
          <bt:LongStrings>
            <bt:String id="residTooltip" DefaultValue="Log notes to Contoso CRM for this appointment."/>
          </bt:LongStrings>
        </Resources>
      </VersionOverrides>
    </VersionOverrides>
    ```

> [!TIP]
> 若要详细了解 Outlook 外接程序的清单，请参阅 [Outlook 外接程序清单](manifests.md) 和 [添加对 Outlook Mobile 外接程序命令的支持](add-mobile-support.md)。

### <a name="capture-appointment-notes"></a>捕获约会说明

在本部分中，了解当用户选择“ **日志** ”按钮时，外接程序如何提取约会详细信息。

1. 在同一快速入门项目中，在代码编辑器中打开文件 **./src/commands/commands.js** 。

1. 将 **commands.js** 文件的整个内容替换为以下 JavaScript。

    ```js
    var event;

    Office.initialize = function (reason) {
      // Add any initialization code here.
    };

    function logCRMEvent(appointmentEvent) {
      event = appointmentEvent;
      console.log(`Subject: ${Office.context.mailbox.item.subject}`);
      Office.context.mailbox.item.body.getAsync(
        "html",
        { asyncContext: "This is passed to the callback" },
        function callback(result) {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            event.completed({ allowEvent: true });
          } else {
            console.error("Failed to get body.");
            event.completed({ allowEvent: false });
          }
        }
      );
    }

    // Register the function.
    Office.actions.associate("logCRMEvent", logCRMEvent);
    ```

接下来，更新 **commands.html** 文件以引用 **commands.js**。

1. 在同一快速入门项目中，在代码编辑器中打开 **./src/commands/commands.html** 文件。

1. 查找并替换 `<script type="text/javascript" src="https://appsforoffice.microsoft.com/lib/1.1/hosted/office.js"></script>` 为以下内容：

    ```html
    <script type="text/javascript" src="https://appsforoffice.microsoft.com/lib/1.1/hosted/office.js"></script>
    <script type="text/javascript" src="commands.js"></script>
    ```

### <a name="view-appointment-notes"></a>查看约会说明

通过设置为此目的保留的 **EventLogged** 自定义属性，可以切换 **“日志**”按钮标签以显示 **视图**。 当用户选择 **“视图** ”按钮时，他们可以查看其为此约会记录的笔记。

外接程序定义日志查看体验。 例如，当用户选择 **“视图** ”按钮时，可以在对话框中显示记录的约会笔记。 有关使用对话框的详细信息，请参阅 [Office 加载项中的“使用 Office”对话框 API](../develop/dialog-api-in-office-add-ins.md)。

将以下函数添加到 **./src/commands/commands.js**。 此函数设置当前约会项上的 **EventLogged** 自定义属性。

```js
function updateCustomProperties() {
  Office.context.mailbox.item.loadCustomPropertiesAsync(
    function callback(customPropertiesResult) {
      if (customPropertiesResult.status === Office.AsyncResultStatus.Succeeded) {
        let customProperties = customPropertiesResult.value;
        customProperties.set("EventLogged", true);
        customProperties.saveAsync(
          function callback(setSaveAsyncResult) {
            if (setSaveAsyncResult.status === Office.AsyncResultStatus.Succeeded) {
              console.log("EventLogged custom property saved successfully.");
              event.completed({ allowEvent: true });
              event = undefined;
            }
          }
        );
      }
    }
  );
}
```

然后在加载项成功记录约会说明后调用它。 例如，可以从 **logCRMEvent** 调用它，如以下函数所示。

```js
function logCRMEvent(appointmentEvent) {
  event = appointmentEvent;
  console.log(`Subject: ${Office.context.mailbox.item.subject}`);
  Office.context.mailbox.item.body.getAsync(
    "html",
    { asyncContext: "This is passed to the callback" },
    function callback(result) {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        // Replace `event.completed({ allowEvent: true });` with the following statement.
        updateCustomProperties();
      } else {
        console.error("Failed to get body.");
        event.completed({ allowEvent: false });
      }
    }
  );
}
```

### <a name="delete-the-appointment-log"></a>删除约会日志

如果希望让用户撤消日志记录或删除记录的约会笔记，以便保存替换日志，则有两个选项。

1. 当用户选择功能区中的相应按钮时，使用 Microsoft Graph [清除自定义属性对象](/graph/api/resources/extended-properties-overview?view=graph-rest-1.0&preserve-view=true) 。
1. 将以下函数添加到 **./src/commands/commands.js** ，以清除当前约会项上的 **EventLogged** 自定义属性。

    ```js
    function clearCustomProperties() {
      Office.context.mailbox.item.loadCustomPropertiesAsync(
        function callback(customPropertiesResult) {
          if (customPropertiesResult.status === Office.AsyncResultStatus.Succeeded) {
            var customProperties = customPropertiesResult.value;
            customProperties.remove("EventLogged");
            customProperties.saveAsync(
              function callback(removeSaveAsyncResult) {
                if (removeSaveAsyncResult.status === Office.AsyncResultStatus.Succeeded) {
                  console.log("Custom properties cleared");
                  event.completed({ allowEvent: true });
                  event = undefined;
                }
              }
            );
          }
        }
      );
    }
    ```

然后在想要清除自定义属性时调用它。 例如，如果设置日志以某种方式失败，则可以从 **logCRMEvent** 调用它，如以下函数所示。

  ```js
  function logCRMEvent(appointmentEvent) {
    event = appointmentEvent;
    console.log(`Subject: ${Office.context.mailbox.item.subject}`);
    Office.context.mailbox.item.body.getAsync(
      "html",
      { asyncContext: "This is passed to the callback" },
      function callback(result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          updateCustomProperties();
        } else {
          console.error("Failed to get body.");
          // Replace `event.completed({ allowEvent: false });` with the following statement.
          clearCustomProperties();
        }
      }
    );
  }
  ```

# <a name="task-pane"></a>[任务窗格](#tab/taskpane)

此选项将使用户能够从任务窗格记录和查看其笔记及其约会的其他详细信息。

### <a name="configure-the-manifest"></a>配置清单

若要使用户能够使用外接程序记录约会笔记，必须在父元素`MobileFormFactor`下的清单中配置 [MobileLogEventAppointmentAttendee 扩展点](/javascript/api/manifest/extensionpoint#mobilelogeventappointmentattendee)。 不支持其他外形因素。

[!INCLUDE [Teams manifest not supported on mobile devices](../includes/no-mobile-with-json-note.md)]

1. 在代码编辑器中，打开快速启动项目。

1. 打开位于项目根 **目录的manifest.xml** 文件。

1. 选择整个 `<VersionOverrides>` 节点 (包括打开和关闭标记) 并将其替换为以下 XML。 请确保将 **对 Contoso** 的所有引用替换为公司信息。

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
                <ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">
                  <OfficeTab id="TabDefault">
                    <Group id="apptReadGroup">
                      <Label resid="residDescription"/>
                      <Control xsi:type="Button" id="apptReadOpenPaneButton">
                        <Label resid="residLabel"/>
                        <Supertip>
                          <Title resid="residLabel"/>
                          <Description resid="residTooltip"/>
                        </Supertip>
                        <Icon>
                          <bt:Image size="16" resid="icon-16"/>
                          <bt:Image size="32" resid="icon-32"/>
                          <bt:Image size="80" resid="icon-80"/>
                        </Icon>
                        <Action xsi:type="ShowTaskpane">
                          <SourceLocation resid="Taskpane.Url"/>
                        </Action>
                      </Control>
                    </Group>
                  </OfficeTab>
                </ExtensionPoint>
              </DesktopFormFactor>
              <MobileFormFactor>
                <ExtensionPoint xsi:type="MobileLogEventAppointmentAttendee">
                  <Control xsi:type="MobileButton" id="appointmentReadFunctionButton">
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
                    <Action xsi:type="ShowTaskpane">
                      <SourceLocation resid="Taskpane.Url"/>
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
              <bt:Image id="icon-80" DefaultValue="https://contoso.com/assets/icon-80.png"/>
            </bt:Images>
            <bt:Urls>
              <bt:Url id="residFunctionFile" DefaultValue="https://contoso.com/commands.html"/>
              <bt:Url id="Taskpane.Url" DefaultValue="https://contoso.com/taskpane.html"/>
            </bt:Urls>
            <bt:ShortStrings>
              <bt:String id="residDescription" DefaultValue="Log appointment notes and other details to Contoso CRM."/>
              <bt:String id="residLabel" DefaultValue="Log to Contoso CRM"/>
            </bt:ShortStrings>
            <bt:LongStrings>
              <bt:String id="residTooltip" DefaultValue="Log notes to Contoso CRM for this appointment."/>
            </bt:LongStrings>
          </Resources>
        </VersionOverrides>
    </VersionOverrides>
    ```

> [!TIP]
> 若要详细了解 Outlook 外接程序的清单，请参阅 [Outlook 外接程序清单](manifests.md) 和 [添加对 Outlook Mobile 外接程序命令的支持](add-mobile-support.md)。

### <a name="capture-appointment-notes"></a>捕获约会说明

在本部分中，了解当用户选择“ **日志** ”按钮时，如何在任务窗格中显示记录的约会笔记和其他详细信息。

1. 在同一快速入门项目中，在代码编辑器中打开文件 **./src/taskpane/taskpane.js** 。

1. 将 **taskpane.js** 文件的整个内容替换为以下 JavaScript。

    ```js
    // Office is ready.
    Office.onReady(function () {
        getEventData();
      }
    );

    function getEventData() {
      console.log(`Subject: ${Office.context.mailbox.item.subject}`);
      Office.context.mailbox.item.body.getAsync(
        "html",
        function callback(result) {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            console.log("event logged successfully");
          } else {
            console.error("Failed to get body.");
          }
        }
      );
    }
    ```

接下来，更新 **taskpane.html** 文件以引用 **taskpane.js**。

1. 在同一快速入门项目中，在代码编辑器中打开 **./src/taskpane/taskpane.html** 文件。

1. 查找并替换 `<script type="text/javascript" src="https://appsforoffice.microsoft.com/lib/1.1/hosted/office.js"></script>` 为以下内容：

    ```html
    <script type="text/javascript" src="https://appsforoffice.microsoft.com/lib/1.1/hosted/office.js"></script>
    <script type="text/javascript" src="taskpane.js"></script>
    ```

### <a name="view-appointment-notes"></a>查看约会说明

通过设置为此目的保留的 **EventLogged** 自定义属性，可以切换 **“日志**”按钮标签以显示 **视图**。 当用户选择 **“视图** ”按钮时，他们可以查看其为此约会记录的笔记。 外接程序定义日志查看体验。

将以下函数添加到 **./src/taskpane/taskpane.js**。 此函数设置当前约会项上的 **EventLogged** 自定义属性。

```js
function updateCustomProperties() {
  Office.context.mailbox.item.loadCustomPropertiesAsync(
    function callback(customPropertiesResult) {
      if (customPropertiesResult.status === Office.AsyncResultStatus.Succeeded) {
        let customProperties = customPropertiesResult.value;
        customProperties.set("EventLogged", true);
        customProperties.saveAsync(
          function callback(setSaveAsyncResult) {
            if (setSaveAsyncResult.status === Office.AsyncResultStatus.Succeeded) {
              console.log("EventLogged custom property saved successfully.");
            }
          }
        );
      }
    }
  );
}
```

然后在加载项成功记录约会说明后调用它。 例如，可以从 **getEventData** 调用它，如以下函数所示。

```js
function getEventData() {
  console.log(`Subject: ${Office.context.mailbox.item.subject}`);
  Office.context.mailbox.item.body.getAsync(
    "html",
    function callback(result) {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        console.log("event logged successfully");
        updateCustomProperties();
      } else {
        console.error("Failed to get body.");
      }
    }
  );
}
```

### <a name="delete-the-appointment-log"></a>删除约会日志

如果希望让用户撤消日志记录或删除记录的约会笔记，以便保存替换日志，则有两个选项。

1. 当用户选择任务窗格中的相应按钮时，使用 Microsoft Graph [清除自定义属性对象](/graph/api/resources/extended-properties-overview?view=graph-rest-1.0&preserve-view=true) 。
1. 将以下函数添加到 **./src/taskpane/taskpane.js** 以清除当前约会项上的 **EventLogged** 自定义属性。

    ```js
    function clearCustomProperties() {
      Office.context.mailbox.item.loadCustomPropertiesAsync(
        function callback(customPropertiesResult) {
          if (customPropertiesResult.status === Office.AsyncResultStatus.Succeeded) {
            var customProperties = customPropertiesResult.value;
            customProperties.remove("EventLogged");
            customProperties.saveAsync(
              function callback(removeSaveAsyncResult) {
                if (removeSaveAsyncResult.status === Office.AsyncResultStatus.Succeeded) {
                  console.log("Custom properties cleared");
                }
              }
            );
          }
        }
      );
    }
    ```

然后在想要清除自定义属性时调用它。 例如，如果设置日志失败，则可以从 **getEventData** 调用它，如以下函数所示。

  ```js
  function getEventData() {
    console.log(`Subject: ${Office.context.mailbox.item.subject}`);
    Office.context.mailbox.item.body.getAsync(
      "html",
      function callback(result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          console.log("event logged successfully");
          updateCustomProperties();
        } else {
          console.error("Failed to get body.");
          clearCustomProperties();
        }
      }
    );
  }
  ```

---

## <a name="test-and-validate"></a>测试和验证

1. 按照通常的指南 [测试和验证加载项](testing-and-tips.md)。
1. 在 Outlook 网页版、Windows 或 Mac 中[旁加载](sideload-outlook-add-ins-for-testing.md)外接程序后，在 Android 移动设备上重启 Outlook。
1. 以与会者身份打开约会，然后验证 **会议见解** 卡下是否有一张新卡片，其中加载项的名称旁边是 **“日志”** 按钮。

### <a name="ui-log-the-appointment-notes"></a>UI：记录约会说明

作为会议与会者，在打开会议时，应会看到类似于下图的屏幕。

![显示 Android 上约会屏幕上的“日志”按钮的屏幕截图。](../images/outlook-android-log-appointment-details.jpg)

### <a name="ui-view-the-appointment-log"></a>UI：查看约会日志

成功记录约会笔记后，该按钮现在应标记为 **“视图** ”而不是 **“日志**”。 应会看到类似于下图的屏幕。

![显示 Android 上约会屏幕上的“查看”按钮的屏幕截图。](../images/outlook-android-view-appointment-log.jpg)

## <a name="available-apis"></a>可用 API

以下 API 可用于此功能。

- [Dialog API](../develop/dialog-api-in-office-add-ins.md)
- [Office.AddinCommands.Event](/javascript/api/office/office.addincommands.event?view=outlook-js-preview&preserve-view=true)
- [Office.CustomProperties](/javascript/api/outlook/office.customproperties?view=outlook-js-preview&preserve-view=true)
- [Office.RoamingSettings](/javascript/api/outlook/office.roamingsettings?view=outlook-js-preview&preserve-view=true)
- [约会读取 (与会者) API](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true)**，但以下内容除外**：
  - [Office.context.mailbox.item.categories](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true#categories)
  - [Office.context.mailbox.item.enhancedLocation](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true#enhancedLocation)
  - [Office.context.mailbox.item.isAllDayEvent](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true#isAllDayEvent)
  - [Office.context.mailbox.item.recurrence](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true#recurrence)
  - [Office.context.mailbox.item.sensitivity](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true#sensitivity)
  - [Office.context.mailbox.item.seriesId](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true#seriesId)

## <a name="restrictions"></a>限制

有几个限制适用。

- 无法更改 **日志** 按钮名称。 但是，可以通过在约会项上设置自定义属性来显示其他标签。 有关更多详细信息，请参阅视 **图约会说明** 部分，了解 [函数命令](?tabs=noui#view-appointment-notes) 或 [任务窗格](?tabs=taskpane#view-appointment-notes-1) 。
- 如果要将 **“日志**”按钮的标签切换到 **“查看**”和“返回”，则必须使用 **EventLogged** 自定义属性。
- 外接程序图标应使用十六进制代码 `#919191` 或 [以其他颜色格式](https://convertingcolors.com/hex-color-919191.html)等效的灰度。
- 加载项应在一分钟的超时时间内从约会表单中提取会议详细信息。 但是，在对话框中花费的任何时间（例如，为身份验证打开的外接程序）都从超时期中排除。

## <a name="see-also"></a>另请参阅

- [适用于 Outlook Mobile 的加载项](outlook-mobile-addins.md)
- [添加对适用于 Outlook Mobile 的外接程序命令的支持](add-mobile-support.md)
