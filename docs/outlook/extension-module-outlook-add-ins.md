---
title: 模块扩展 Outlook 加载项
description: 可以创建在 Outlook 中运行的应用程序，以便用户无需退出 Outlook 即可轻松地访问业务信息和工作效率工具。
ms.date: 08/30/2022
ms.localizationpriority: medium
ms.openlocfilehash: d234f4e1aad77b3cc30d0e9bc9450ec79af958aa
ms.sourcegitcommit: eef2064d7966db91f8401372dd255a32d76168c2
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/31/2022
ms.locfileid: "67464802"
---
# <a name="module-extension-outlook-add-ins"></a>模块扩展 Outlook 加载项

模块扩展加载项与邮件、任务和日历一起显示在 Outlook 导航栏中。 模块扩展不限于使用邮件和约会信息。 可以创建在 Outlook 中运行的应用程序，以便用户无需退出 Outlook 即可轻松地访问业务信息和工作效率工具。

> [!TIP]
> Teams 清单中不支持模块扩展 [ (预览) ](../develop/json-manifest-overview.md)，但可以通过在 [Outlook 中打开个人选项卡](/microsoftteams/platform/m365-apps/extend-m365-teams-personal-tab)为用户创建非常相似的体验。 在 Outlook 外接程序中 Teams 清单的早期预览期内，无法在同一清单中合并 Outlook 外接程序和个人选项卡，并将其安装为单元。 我们正在努力解决此问题，但与此同时，必须为加载项和个人选项卡创建单独的应用。它们都可以在同一域上使用文件。

> [!NOTE]
> 仅 Windows 上的 Outlook 2016 或更高版本支持模块扩展。  

## <a name="open-a-module-extension"></a>打开模块扩展

要打开模块扩展，用户单击 Outlook 导航栏中的模块的名称或图标即可。如果用户选择了紧凑型导航，导航栏有一个显示已加载扩展的图标。

![当模块扩展在 Outlook 中加载时，显示紧凑型导航栏。](../images/outlook-module-navigationbar-compact.png)

如果用户没有使用紧凑型导航，则导航栏有两种外观。 加载一个扩展后，它将显示加载项的名称。

![当一个模块扩展在 Outlook 中加载时，显示展开的导航栏。](../images/outlook-module-navigationbar-one.png)

在加载多个加载项时，会显示 **加载项** 一词。单击其中任何一个即可打开扩展的用户界面。

![当多个模块扩展在 Outlook 中加载时，显示展开的导航栏。](../images/outlook-module-navigationbar-more.png)

在单击扩展时，Outlook 会将内置模块替换为自定义模块，以便用户可以与该加载项进行交互。 可以在加载项中使用 Outlook JavaScript API 的一些功能。 逻辑上假定特定 Outlook 项（如消息或约会）的 API 在模块扩展中不起作用。 该模块还可以在 Outlook 功能区中包含与加载项页面交互的函数命令。 为此，函数命令调用 [Office.onReady 或 Office.initialize 方法](../develop/initialize-add-in.md) 和 [Event.completed](/javascript/api/office/office.addincommands.event#office-office-addincommands-event-completed-member(1)) 方法。 若要演练如何配置模块扩展 Outlook 外接程序，请参阅 [Outlook 模块扩展计费小时示例](https://github.com/OfficeDev/Outlook-Add-in-JavaScript-ModuleExtension)。

以下屏幕截图显示了一个外接程序，该外接程序集成在 Outlook 导航栏中，并具有用于更新外接程序页面的功能区命令。

![显示模块扩展的用户界面。](../images/outlook-module-extension.png)

## <a name="example"></a>示例

下面是定义模块扩展的清单文件部分。

```xml
<!-- Add Outlook module extension point -->
<VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides"
                  xsi:type="VersionOverridesV1_0">
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1"
                    xsi:type="VersionOverridesV1_1">

    <!-- Begin override of existing elements -->
    <Description resid="residVersionOverrideDesc" />

    <Requirements>
      <bt:Sets DefaultMinVersion="1.3">
        <bt:Set Name="Mailbox" />
      </bt:Sets>
    </Requirements>
    <!-- End override of existing elements -->

    <Hosts>
      <Host xsi:type="MailHost">
        <DesktopFormFactor>
          <!-- Set the URL of the file that contains the
                JavaScript function that controls the extension -->
          <FunctionFile resid="residFunctionFileUrl" />

          <!--New Extension Point - Module for a ModuleApp -->
          <ExtensionPoint xsi:type="Module">
            <SourceLocation resid="residExtensionPointUrl" />
            <Label resid="residExtensionPointLabel" />

            <CommandSurface>
              <CustomTab id="idTab">
                <Group id="idGroup">
                  <Label resid="residGroupLabel" />

                  <Control xsi:type="Button" id="group.changeToAssociate">
                    <Label resid="residChangeToAssociateLabel" />
                    <Supertip>
                      <Title resid="residChangeToAssociateLabel" />
                      <Description resid="residChangeToAssociateDesc" />
                    </Supertip>
                    <Icon>
                      <bt:Image size="16" resid="residAssociateIcon16" />
                      <bt:Image size="32" resid="residAssociateIcon32" />
                      <bt:Image size="80" resid="residAssociateIcon80" />
                    </Icon>
                    <Action xsi:type="ExecuteFunction">
                      <FunctionName>changeToAssociateRate</FunctionName>
                    </Action>
                  </Control>
                  
              </Group>
                <Label resid="residCustomTabLabel" />
              </CustomTab>
            </CommandSurface>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>
    </Hosts>

    <Resources>
      <bt:Images>
        <bt:Image id="residAddinIcon16" 
                  DefaultValue="https://localhost:8080/Executive-16.png" />
        <bt:Image id="residAddinIcon32" 
                  DefaultValue="https://localhost:8080/Executive-32.png" />
        <bt:Image id="residAddinIcon80" 
                  DefaultValue="https://localhost:8080/Executive-80.png" />
      
        <bt:Image id="residAssociateIcon16" 
                  DefaultValue="https://localhost:8080/Associate-16.png" />
        <bt:Image id="residAssociateIcon32" 
                  DefaultValue="https://localhost:8080/Associate-32.png" />
        <bt:Image id="residAssociateIcon80" 
                  DefaultValue="https://localhost:8080/Associate-80.png" />
      </bt:Images>

      <bt:Urls>
        <bt:Url id="residFunctionFileUrl" 
                DefaultValue="https://localhost:8080/" />
        <bt:Url id="residExtensionPointUrl" 
                DefaultValue="https://localhost:8080/" />
      </bt:Urls>

      <!--Short strings must be less than 30 characters long -->
      <bt:ShortStrings>
        <bt:String id="residExtensionPointLabel" 
                    DefaultValue="Billable Hours" />
        <bt:String id="residGroupLabel" 
                    DefaultValue="Change billing rate" />
        <bt:String id="residCustomTabLabel" 
                    DefaultValue="Billable hours" />

        <bt:String id="residChangeToAssociateLabel" 
                    DefaultValue="Associate" />
      </bt:ShortStrings>

      <bt:LongStrings>
        <bt:String id="residVersionOverrideDesc" 
                    DefaultValue="Version override description" />

        <bt:String id="residChangeToAssociateDesc" 
                    DefaultValue="Change to the associate billing rate: $127/hr" />
      </bt:LongStrings>
    </Resources>
  </VersionOverrides>
</VersionOverrides>
```

## <a name="see-also"></a>另请参阅

- [Outlook 加载项清单](manifests.md)
- [适用于 Outlook 的外接程序命令](add-in-commands-for-outlook.md)
- [Outlook 模块扩展计酬时间示例](https://github.com/OfficeDev/Outlook-Add-in-JavaScript-ModuleExtension)
