---
title: '在外接程序预览版中Outlook智能警报 (OnMessageSend) '
description: 了解如何使用基于事件的激活在 Outlook外接程序中处理发送邮件事件。
ms.topic: article
ms.date: 03/07/2022
ms.localizationpriority: medium
ms.openlocfilehash: 2a9d44844c7fff3d5305de53f57c2950ae1909fb
ms.sourcegitcommit: b66ba72aee8ccb2916cd6012e66316df2130f640
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/26/2022
ms.locfileid: "64484502"
---
# <a name="use-smart-alerts-and-the-onmessagesend-event-in-your-outlook-add-in-preview"></a>在外接程序预览版中Outlook智能警报 (OnMessageSend) 

事件`OnMessageSend`利用智能警报，允许用户在用户选择其邮件中的"发送"后Outlook逻辑。 事件处理程序允许你为用户提供在发送电子邮件之前改进其电子邮件的机会。 事件 `OnAppointmentSend` 相似，但适用于约会。

在此演练结束时，您将拥有一个外接程序，该外接程序在邮件发送时运行，并检查用户是否忘记添加电子邮件中提到的文档或图片。

> [!IMPORTANT]
> `OnAppointmentSend`和 `OnMessageSend` 事件仅在预览版中提供，Microsoft 365订阅位于 Outlook Windows。 有关详细信息，请参阅 [如何预览](autolaunch.md#how-to-preview)。 预览事件不应在生产外接程序中使用。

## <a name="prerequisites"></a>先决条件

该事件 `OnMessageSend` 通过基于事件的激活功能提供。 若要了解如何将加载项配置为使用此功能、可用事件、如何预览此事件、调试、功能限制等，请参阅配置 Outlook 加载项进行基于事件的[激活](autolaunch.md)。

## <a name="set-up-your-environment"></a>设置环境

完成[Outlook](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator)使用适用于加载项的 Yeoman 生成器创建加载项Office快速入门。

## <a name="configure-the-manifest"></a>配置清单

1. 在代码编辑器中，打开快速启动项目。

1. 打开 **manifest.xml** 根目录下的文件。

1. Select the entire **VersionOverrides** node (including open and close tags) and replace it with the following XML， then save your changes.

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
               This is used by Outlook on the web and Outlook on the new Mac UI preview. -->
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

          <!-- Enable launching the add-in on the included event. -->
          <ExtensionPoint xsi:type="LaunchEvent">
            <LaunchEvents>
              <LaunchEvent Type="OnMessageSend" FunctionName="onMessageSendHandler" SendMode="PromptUser" />
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

> [!TIP]
>
> - 有关 **事件提供的 SendMode** 选项 `OnMessageSend` ，请参阅 [可用 SendMode 选项](/javascript/api/manifest/launchevent#available-sendmode-options-preview)。
> - 若要了解有关加载项清单Outlook，请参阅Outlook[加载项清单](manifests.md)。

## <a name="implement-event-handling"></a>实现事件处理

您必须对所选事件实现处理。

在此方案中，您将添加用于发送邮件的处理。 外接程序将检查邮件中的某些关键字。 如果找到其中任何关键字，它将检查是否有附件。 如果没有附件，外接程序将建议用户添加可能缺少的附件。

1. 从同一快速启动项目中，在 **./src** 目录下新建一个名为 **launchevent** 的文件夹。

1. 在 **"./src/launchevent** "文件夹中，新建一个名为"launchevent.js **"**。

1. 在代码编辑器 **中打开文件 ./src/launchevent/launchevent.js** 并添加以下 JavaScript 代码。

    ```js
    /*
    * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
    * See LICENSE in the project root for license information.
    */

    function onMessageSendHandler(event) {
      Office.context.mailbox.item.body.getAsync(
        "text",
        { asyncContext: event },
        getBodyCallback
      );
    }

    function getBodyCallback(asyncResult){
      let event = asyncResult.asyncContext;
      let body = "";
      if (asyncResult.status !== Office.AsyncResultStatus.Failed && asyncResult.value !== undefined) {
        body = asyncResult.value;
      } else {
        let message = "Failed to get body text";
        console.error(message);
        event.completed({ allowEvent: false, errorMessage: message });
        return;
      }

      let matches = hasMatches(body);
      if (matches) {
        Office.context.mailbox.item.getAttachmentsAsync(
          { asyncContext: event },
          getAttachmentsCallback);
      } else {
        event.completed({ allowEvent: true });
      }
    }

    function hasMatches(body) {
      if (body == null || body == "") {
        return false;
      }

      const arrayOfTerms = ["send", "picture", "document", "attachment"];
      for (let index = 0; index < arrayOfTerms.length; index++) {
        const term = arrayOfTerms[index].trim();
        const regex = RegExp(term, 'i');
        if (regex.test(body)) {
          return true;
        }
      }

      return false;
    }

    function getAttachmentsCallback(asyncResult) {
      let event = asyncResult.asyncContext;
      if (asyncResult.value.length > 0) {
        for (let i = 0; i < asyncResult.value.length; i++) {
          if (asyncResult.value[i].isInline == false) {
            event.completed({ allowEvent: true });
            return;
          }
        }

        event.completed({ allowEvent: false, errorMessage: "Looks like you forgot to include an attachment?" });
      } else {
        event.completed({ allowEvent: false, errorMessage: "Looks like you're forgetting to include an attachment?" });
      }
    }

    // 1st parameter: FunctionName of LaunchEvent in the manifest; 2nd parameter: Its implementation in this .js file.
    Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
    ```

1. 保存所做的更改。

## <a name="update-webpack-config-settings"></a>更新 webpack 配置设置

1. 打开 **webpack.config.js** 根目录中找到的目录文件，并完成以下步骤。

1. 在 对象 `plugins` 内找到 数组 `config` ，将此新对象添加到数组的开头。

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

1. 在项目的根目录中运行以下命令。 运行 时 `npm start`，如果本地 Web 服务器尚未运行 (将启动) 外接程序将旁加载。

    ```command&nbsp;line
    npm run build
    ```
    ```command&nbsp;line
    npm start
    ```

    > [!NOTE]
    > 如果加载项未自动旁加载，请按照旁加载 [Outlook](../outlook/sideload-outlook-add-ins-for-testing.md#sideload-manually) 加载项进行测试中的说明，在加载项中手动旁加载Outlook。

1. 在Outlook中Windows新建邮件并设置主题。 在正文中，添加类似"你好，查看我的 dog 的此图片！"这样的文本。
1. 发送消息. 应弹出一个对话框，建议你添加附件。

    !["使用对话框打开Outlook Windows窗口的屏幕截图。](../images/outlook-win-smart-alert.png)

1. 添加附件，然后再次发送邮件。 此时应该没有警报。

## <a name="see-also"></a>另请参阅

- [Outlook 加载项清单](manifests.md)
- [配置Outlook加载项进行基于事件的激活](autolaunch.md)
- [如何调试基于事件的外接程序](debug-autolaunch.md)
- [基于事件的加载项的 AppSource Outlook选项](autolaunch-store-options.md)
