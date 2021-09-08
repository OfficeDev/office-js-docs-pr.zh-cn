---
title: 在加载项中实现Outlook On-send
description: 了解如何在加载项中实现附加的Outlook功能。
ms.topic: article
ms.date: 02/01/2021
localization_priority: Normal
ms.openlocfilehash: 8b69fbbaef1d0f060f0675fe5c4948a70d935b7a
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/08/2021
ms.locfileid: "58939255"
---
# <a name="implement-append-on-send-in-your-outlook-add-in"></a>在加载项中实现Outlook On-send

在此演练结束时，您将拥有一个Outlook外接程序，该外接程序可以在邮件发送时插入免责声明。

> [!NOTE]
> 要求集 1.9 中引入了对此功能的支持。 请查看支持此要求集的[客户端和平台](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients)。

## <a name="set-up-your-environment"></a>设置环境

完成[Outlook](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator)使用适用于加载项的 Yeoman 生成器创建加载项Office快速入门。

## <a name="configure-the-manifest"></a>配置清单

若要在加载项中启用附加 Onss 功能，必须在 `AppendOnSend` [ExtendedPermissions](../reference/manifest/extendedpermissions.md)集合中包括 权限。

对于此方案，你将运行 函数，而不是在选择"操作"按钮时运行 `action`  `appendOnSend` 函数。

1. 在代码编辑器中，打开快速启动项目。

1. 打开 **manifest.xml** 根目录下的文件。

1. Select the entire `<VersionOverrides>` node (including open and close tags) and replace it with the following XML.

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
            <DesktopFormFactor>
              <FunctionFile resid="Commands.Url" />
              <ExtensionPoint xsi:type="MessageComposeCommandSurface">
                <OfficeTab id="TabDefault">
                  <Group id="msgComposeGroup">
                    <Label resid="GroupLabel" />
                    <Control xsi:type="Button" id="msgComposeOpenPaneButton">
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
                        <FunctionName>appendDisclaimerOnSend</FunctionName>
                      </Action>
                    </Control>
                  </Group>
                </OfficeTab>
              </ExtensionPoint>

              <!-- Configure AppointmentOrganizerCommandSurface extension point to support
              append on sending a new appointment. -->

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
        <ExtendedPermissions>
          <ExtendedPermission>AppendOnSend</ExtendedPermission>
        </ExtendedPermissions>
      </VersionOverrides>
    </VersionOverrides>
    ```

> [!TIP]
> 若要了解有关加载项清单Outlook，请参阅Outlook[加载项清单。](manifests.md)

## <a name="implement-append-on-send-handling"></a>实现附加 On-send 处理

接下来，在发送事件上实现追加。

> [!IMPORTANT]
> 如果外接程序还使用 实现[Onss `ItemSend` ons](outlook-on-send-addins.md)发送事件处理，则 Onss ons handler 中的调用将返回错误，因为不支持此 `AppendOnSendAsync` 方案。

对于此方案，你将在用户发送时实现向项目附加免责声明。

1. 从同一快速启动项目中，在代码编辑器中打开文件 **./src/commands/commands.js** 文件。

1. 在 函数 `action` 后插入以下 JavaScript 函数。

    ```js
    function appendDisclaimerOnSend(event) {
      var appendText =
        '<p style = "color:blue"> <i>This and subsequent emails on the same topic are for discussion and information purposes only. Only those matters set out in a fully executed agreement are legally binding. This email may contain confidential information and should not be shared with any third party without the prior written agreement of Contoso. If you are not the intended recipient, take no action and contact the sender immediately.<br><br>Contoso Limited (company number 01624297) is a company registered in England and Wales whose registered office is at Contoso Campus, Thames Valley Park, Reading RG6 1WG</i></p>';  
      /**
        *************************************************************
         Ideal Usage - Call the getBodyType API. Use the coercionType
         it returns as the parameter value below.
        *************************************************************
      */
      Office.context.mailbox.item.body.appendOnSendAsync(
        appendText,
        {
          coercionType: Office.CoercionType.Html
        },
        function(asyncResult) {
          console.log(asyncResult);
        }
      );

      event.completed();
    }
    ```

1. 在文件末尾，添加以下语句。

    ```js
    g.appendDisclaimerOnSend = appendDisclaimerOnSend;
    ```

## <a name="try-it-out"></a>试用

1. 在项目的根目录中运行以下命令。 运行此命令时，如果本地 Web 服务器尚未运行，它将启动，并且外接程序将旁加载。 

    ```command&nbsp;line
    npm start
    ```

1. 创建新邮件，并将自己添加到 **"收件人"** 行。

1. 从功能区或溢出菜单中，选择 **"执行操作"。**

1. 发送邮件，然后从"收件箱"或"已发送邮件"文件夹中 **打开** 邮件，以查看附加的免责声明。

    ![在发送时附加免责声明的示例邮件屏幕截图Outlook 网页版。](../images/outlook-web-append-disclaimer.png)

## <a name="see-also"></a>另请参阅

[Outlook 加载项清单](manifests.md)
