---
title: 在 Outlook 加载项中实现 append-on-send
description: 了解如何在 Outlook 加载项中实现“发送时追加”功能。
ms.topic: article
ms.date: 10/24/2022
ms.localizationpriority: medium
ms.openlocfilehash: c8239634b6c9ca281255caf89276fb1b454efc84
ms.sourcegitcommit: 693e9a9b24bb81288d41508cb89c02b7285c4b08
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/28/2022
ms.locfileid: "68767159"
---
# <a name="implement-append-on-send-in-your-outlook-add-in"></a>在 Outlook 加载项中实现 append-on-send

在本演练结束时，你将拥有一个 Outlook 加载项，可在发送邮件时插入免责声明。

> [!NOTE]
> 要求集 1.9 中引入了对此功能的支持。 请查看支持此要求集的[客户端和平台](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets#requirement-sets-supported-by-exchange-servers-and-outlook-clients)。

## <a name="set-up-your-environment"></a>设置环境

完成 [Outlook 快速入门](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) ，该快速入门使用 Office 加载项的 Yeoman 生成器创建外接程序项目。

## <a name="configure-the-manifest"></a>配置清单

若要配置清单，请打开要使用的清单类型的选项卡。

# <a name="xml-manifest"></a>[XML 清单](#tab/xmlmanifest)

若要在外接程序中启用“发送时追加”功能，必须在 `AppendOnSend` [ExtendedPermissions](/javascript/api/manifest/extendedpermissions) 的集合中包含 权限。

对于此方案，你将运行函数 `action` ，而不是在选择“ **执行操作** ”按钮时运行 `appendOnSend` 函数。

1. 在代码编辑器中，打开快速入门项目。

1. 打开位于项目根目录处的 **manifest.xml** 文件。

1. 选择整个 **\<VersionOverrides\>** 节点 (包括打开和关闭标记) ，并将其替换为以下 XML。

    ```XML
    <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
      <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
        <Requirements>
          <bt:Sets DefaultMinVersion="1.9">
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

# <a name="teams-manifest-developer-preview"></a>[Teams 清单 (开发人员预览版) ](#tab/jsonmanifest)

> [!IMPORTANT]
> Office 外接程序的 Teams 清单尚不支持在发送时追加 [ (预览版) ](../develop/json-manifest-overview.md)。 此选项卡供将来使用。

1. 打开 manifest.json 文件。

1. 将以下 对象添加到“extensions.runtimes”数组。 对于此代码，请注意以下事项。

   - 邮箱要求集的“minVersion”设置为“1.9”，因此无法在不支持此功能的平台和 Office 版本上安装加载项。 
   - 运行时的“id”设置为描述性名称“function_command_runtime”。
   - “code.page”属性设置为将加载函数命令的无 UI HTML 文件的 URL。
   - “lifetime”属性设置为“short”，这意味着运行时在选择函数命令按钮时启动，并在函数完成时关闭。  (在某些情况下，运行时在处理程序完成之前关闭。 请参阅 [Office Add-ins.) 中的运行时](../testing/runtimes.md)
   - 有一个操作可以运行名为“appendDisclaimerOnSend”的函数。 你将在后面的步骤中创建此函数。

    ```json
    {
        "requirements": {
            "capabilities": [
                {
                    "name": "Mailbox",
                    "minVersion": "1.9"
                }
            ],
            "formFactors": [
                "desktop"
            ]
        },
        "id": "function_command_runtime",
        "type": "general",
        "code": {
            "page": "https://localhost:3000/commands.html"
        },
        "lifetime": "short",
        "actions": [
            {
                "id": "appendDisclaimerOnSend",
                "type": "executeFunction",
                "displayName": "appendDisclaimerOnSend"
            }
        ]
    }
    ```

1. 在“authorization.permissions.resourceSpecific”数组中，添加以下 对象。 请确保它与数组中的其他对象之间用逗号分隔。

    ```json
    {
      "name": "Mailbox.AppendOnSend.User",
      "type": "Delegated"
    }
    ```

---

> [!TIP]
> 若要详细了解 Outlook 外接程序清单，请参阅 [Outlook 外接程序清单](manifests.md)。

## <a name="implement-append-on-send-handling"></a>实现发送时追加处理

接下来，在发送事件上实现追加。

> [!IMPORTANT]
> 如果外接程序还[使用 `ItemSend`实现 on-send 事件处理](outlook-on-send-addins.md)，则发送时处理程序中的调用`AppendOnSendAsync`将返回错误，因为不支持此方案。

对于此方案，你将实现在用户发送时向项追加免责声明。

1. 在同一快速入门项目中，在代码编辑器中打开文件 **./src/commands/commands.js** 。

1. 在 `action` 函数后面插入以下 JavaScript 函数。

    ```js
    function appendDisclaimerOnSend(event) {
      const appendText =
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

1. 紧靠在函数下方添加以下行来注册函数。

    ```js
    Office.actions.associate("appendDisclaimerOnSend", appendDisclaimerOnSend);
    ```

## <a name="try-it-out"></a>试用

1. 在项目的根目录中运行以下命令。 运行此命令时，如果本地 Web 服务器尚未运行，则本地 Web 服务器将启动，并且加载项将被旁加载。

    ```command&nbsp;line
    npm start
    ```

1. 创建新邮件，并将自己添加到 **“To** ”行。

1. 在功能区或溢出菜单中，选择“ **执行操作**”。

1. 发送邮件，然后从 **“收件箱”** 或“ **已发送邮件”** 文件夹打开邮件，以查看追加的免责声明。

    ![在Outlook 网页版发送时追加了免责声明的示例消息。](../images/outlook-web-append-disclaimer.png)

## <a name="see-also"></a>另请参阅

[Outlook 加载项清单](manifests.md)
