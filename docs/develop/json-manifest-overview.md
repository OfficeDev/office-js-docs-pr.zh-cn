---
title: 适用于 Office 加载项的 Teams 清单（预览版）
description: 获取预览版 JSON 清单的概述。
ms.date: 05/24/2022
ms.localizationpriority: high
ms.openlocfilehash: 8a40f28674892545dee00e5a3138b55400d04352
ms.sourcegitcommit: 35e7646c5ad0d728b1b158c24654423d999e0775
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/02/2022
ms.locfileid: "65833897"
---
# <a name="teams-manifest-for-office-add-ins-preview"></a>适用于 Office 加载项的 Teams 清单（预览版）

Microsoft 正在对 Microsoft 365 开发人员平台进行大量改进。 这些改进在开发、部署、安装和管理所有类型的 Microsoft 365 扩展（包括 Office 加载项）方面提供了更多一致性。这些更改与现有加载项兼容。 

我们正在努力的一个重要改进是能够基于当前 JSON 格式的 Teams 清单，使用相同的清单格式和架构为所有 Microsoft 365 扩展创建单个分发单元。

我们已针对这些目标执行了重要的第一步，让你能够使用 Teams JSON 清单的版本创建仅在 Windows 上运行的 Outlook 加载项。

> [!NOTE]
> 新清单提供预览版，可能会根据反馈进行更改。 我们鼓励经验丰富的加载项开发人员进行试验。 预览清单不应用于生产加载项。 

在早期预览期间，以下限制适用。

- Teams 清单的预览版仅支持 Outlook 加载项，并且仅支持订阅 Office for Windows。 我们正在努力扩展对 Excel、PowerPoint 和 Word 的支持。
- 尚无法将加载项与 Teams 应用（如 Teams 个人选项卡或其他Microsoft 365扩展类型）合并和旁加载。 在接下来的几个月中，我们将继续扩展预览版以支持这些方案，并提供其他工具来将清单更新为预览格式。

> [!TIP]
> 准备好开始使用预览版 Teams 清单了吗？ 开始 [使用 Teams 清单（预览版）生成 Outlook 加载项](../quickstarts/outlook-quickstart-json-manifest.md)。

## <a name="overview-of-the-json-manifest"></a>JSON 清单概述

### <a name="schemas-and-general-points"></a>架构和常规点

[预览版 JSON 清单](/microsoftteams/platform/resources/dev-preview/developer-preview-intro) 只有一个架构，而当前 XML 清单总共有七个 [架构](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8)。  

### <a name="conceptual-mapping-of-the-preview-json-and-current-xml-manifests"></a>预览版 JSON 和当前 XML 清单的概念映射

本部分介绍适用于熟悉当前 XML 清单的读者的预览 JSON 清单。 需要记住的一些要点： 

- JSON 不会像 XML 那样区分属性值和元素值。 通常，映射到 XML 元素的 JSON 使元素值和每个属性都成为子属性。 下面的示例演示了一些 XML 标记及其等效的 JSON。
  
  ```xml
  <MyThing color="blue">Some text</MyThing>
  ```

  ```json
  "myThing" : {
      "color": "blue",
      "text": "Some text"
  }
  ```
- 在当前 XML 清单中有许多位置，具有复数名称的元素具有同名的单一版本的子级。 例如，用于配置自定义菜单的标记包括一个 **Items** 元素，该元素可以具有多个 **Item** 元素子级。 这些复数元素的 JSON 等效项是具有数组作为其值的属性。 数组的成员是 *匿名* 对象，而不是名为“item”或“item1”、“item2”等的属性。下面为一个示例。

  ```json
  "items": [
      {
          -- markup for a menu item is here --
      },
      {
          -- markup for another menu item is here --
      }
  ]
  ```

#### <a name="top-level-structure"></a>顶级结构

预览版 JSON 清单的根级别（大致对应于当前 XML 清单中的 **OfficeApp** 元素）是匿名对象。 

**OfficeApp** 的子级通常分为两个非传统类别。 **VersionOverrides** 元素是一个类别。 另一个包含 **OfficeApp** 的所有其他子级，它们统称为基本清单。 因此，预览版 JSON 清单具有类似的除法。 有一个顶级的“extension”属性，该属性在其用途和子属性中大致对应于 **VersionOverrides** 元素。 预览版 JSON 清单还具有 10 多个其他顶级属性，它们共同为 XML 清单的基本清单提供相同的用途。 可以将这些其他属性统称为 JSON 清单的基本清单。 

> [!NOTE]
> 如果可以在单个清单中将外接程序与其他 Microsoft 365 扩展类型组合在一起，则还有其他顶级属性不符合基本清单的概念。 通常，每种 Microsoft 365 扩展类型都有顶级属性，例如 "configurableTabs"、"bots" 和 "connectors"。 有关示例，请参阅 [Teams 清单文档](/microsoftteams/platform/resources/schema/manifest-schema)。 此结构明确指出，“扩展”属性将 Office 加载项表示为一种 Microsoft 365 扩展。

#### <a name="base-manifest"></a>基本清单

基本清单属性指定 *任何类型的扩展* Microsoft 365 应具有的外接程序的特征。 这包括 Teams 选项卡和消息扩展，而不仅仅是 Office 加载项。这些特征包括公共名称和唯一 ID。 下表显示了预览版 JSON 清单中某些关键顶级属性到当前清单中的 XML 元素的映射，其中映射原则是标记的 *用途*。

|JSON 属性|用途|XML 元素|备注|
|:-----|:-----|:-----|:-----|
|"$schema"| 标识清单架构。 | **OfficeApp** 和 **VersionOverrides** 的属性 | |
|"id"| 外接程序的 GUID。 | **Id**| |
|"version"| 加载项的版本。 | **版本** | |
|"manifestVersion"| 清单架构的版本。 |  **OfficeApp** 的属性 | |
|"name"| 加载项的公共名称。 | **DisplayName** | |
|"description"| 加载项的公共说明。  | **说明** | |
|"accentColor"||| 此属性在当前 XML 清单中没有等效项，并且不用于 JSON 清单的预览版。 但它必须存在。 |
|“developer”| 标识加载项的开发人员。 | **ProviderName** | |
|"localizationInfo"| 配置默认区域设置和其他受支持的区域设置。 | **DefaultLocale** 和 **Override** | |
|"webApplicationInfo"| 标识加载项的 Web 应用，因为它在 Azure Active Directory 中是已知的。 | **WebApplicationInfo** | 在当前 XML 清单中， **WebApplicationInfo** 元素位于 **VersionOverrides** 内部，而不是基本清单中。 |
|"authorization"| 标识加载项所需的任何 Microsoft Graph 权限。 | **WebApplicationInfo** | 在当前 XML 清单中， **WebApplicationInfo** 元素位于 **VersionOverrides** 内部，而不是基本清单中。 |

**Hosts**、**Requirements** 和 **ExtendedOverrides** 元素是当前 XML 清单中基本清单的一部分。 但与这些元素关联的概念和目的在预览版 JSON 清单的“扩展”属性中进行配置。 

#### <a name="extension-property"></a>"extension" 属性

预览版 JSON 清单中的“extension”属性主要表示与其他类型的 Microsoft 365 扩展无关的外接程序的特征。 例如，外接程序扩展的 Office 应用程序（如 Excel、PowerPoint、Word 和 Outlook）在 "extension" 属性中指定，Office 应用程序功能区的自定义项也指定。 "extension" 属性的配置目的与当前 XML 清单中 **VersionOverrides** 元素的配置目的非常匹配。

> [!NOTE]
> 当前 XML 清单的 **VersionOverrides** 部分对许多字符串资源具有“双跳转”系统。 在 **VersionOverrides** 的 **Resources** 子级中指定字符串（包括 URL）并为其分配 ID。 需要字符串的元素具有与 **Resources** 元素中字符串 ID 匹配的`resid`属性。 预览版 JSON 清单的 "extension" 属性通过将字符串直接定义为属性值来简化操作。 JSON 清单中没有任何等效于 **Resources** 的元素。

下表显示了预览版 JSON 清单中 "extension" 属性的某些高级子属性与当前清单中的 XML 元素的映射。 点表示法用于引用子属性。

|JSON 属性|用途|XML 元素|备注|
|:-----|:-----|:-----|:-----|
| "requirements.capabilities" | 标识加载项需要安装的要求集。 | **要求** 和 **集** | |
| "requirements.scopes" | 标识可在其中安装加载项的 Office 应用程序。 | **Hosts** |  |
| "ribbons" | 加载项自定义的功能区。 | **Hosts**、**ExtensionPoints** 和各种 **\*FormFactor** 元素 | "ribbons" 属性是匿名对象的数组，每个对象合并这三个元素的目的。 请参阅 [“功能区”表](#ribbons-table)。|
| "alternatives" | 指定与等效的 COM 加载项、XLL 或两者的向后兼容性。 | **EquivalentAddins** | 请参阅 [EquivalentAddins - 另请参阅](/javascript/api/manifest/equivalentaddins#see-also) 背景信息。 |
| "runtimes"  | 配置各种“无 UI”加载项，例如，自定义函数和函数直接从自定义功能区按钮运行。 | **运行时**。 **FunctionFile** 和 **ExtensionPoint** （CustomFunctions 类型） |  |
| "autoRunEvents" | 配置指定事件的事件处理程序。 | **事件** 和 **扩展点**（事件类型） |  |

##### <a name="ribbons-table"></a>"ribbons" 表

下表将 "ribbons" 数组中匿名子对象的子属性映射到当前清单中的 XML 元素。 

|JSON 属性|用途|XML 元素|备注|
|:-----|:-----|:-----|:-----|
| "contexts" | 指定加载项自定义的命令图面。 | 各种 **\*CommandSurface** 元素，如 **PrimaryCommandSurface** 和 **MessageReadCommandSurface** |  |
| "tabs" | 配置自定义功能区选项卡。 | **CustomTab** | "tabs" 的后代属性的名称和层次结构与 **CustomTab** 的后代非常匹配。  |

## <a name="sample-preview-json-manifest"></a>示例预览 JSON 清单

下面是加载项的预览 JSON 清单示例。

```json
{
  "$schema": "https://raw.githubusercontent.com/OfficeDev/microsoft-teams-app-schema/op/extensions/MicrosoftTeams.schema.json",
  "id": "00000000-0000-0000-0000-000000000000",
  "version": "1.0.0",
  "manifestVersion": "devPreview",
  "name": {
    "short": "Name of your app (<=30 chars)",
    "full": "Full name of app, if longer than 30 characters (<=100 chars)"
  },
  "description": {
    "short": "Short description of your app (<= 80 chars)",
    "full": "Full description of your app (<= 4000 chars)"
  },
  "icons": {
    "outline": "outline.png",
    "color": "color.png"
  },
  "accentColor": "#230201",
  "developer": {
    "name": "Contoso",
    "websiteUrl": "https://www.contoso.com",
    "privacyUrl": "https://www.contoso.com/privacy",
    "termsOfUseUrl": "https://www.contoso.com/servicesagreement"
  },
  "localizationInfo": {
    "defaultLanguageTag": "en-us",
    "additionalLanguages": [
      {
        "languageTag": "es-es",
        "file": "es-es.json"
      }
    ]
  },
  "webApplicationInfo": {
    "id": "00000000-0000-0000-0000-000000000000",
    "resource": "api://www.contoso.com/prodapp"
  },
  "authorization": {
    "permissions": {
      "resourceSpecific": [
        {
          "name": "Mailbox.ReadWrite.User",
          "type": "Delegated"
        }
      ]
    }
  },
  "extensions": [
    {
      "requirements": {
        "scopes": [ "mail" ],
        "capabilities": [
          {
            "name": "Mailbox", "minVersion": "1.1"
          }
        ]
      },
      "runtimes": [
        {
          "requirements": {
            "capabilities": [
              {
                "name": "MailBox", "minVersion": "1.10"
              }
            ]
          },
          "id": "eventsRuntime",
          "type": "general",
          "code": {
            "page": "https://contoso.com/events.html",
            "script": "https://contoso.com/events.js"
          },
          "lifetime": "short",
          "actions": [
            {
              "id": "onMessageSending",
              "type": "executeFunction"
            },
            {
              "id": "onNewMessageComposeCreated",
              "type": "executeFunction"
            }
          ]
        },
        {
          "requirements": {
            "capabilities": [
              {
                "name": "MailBox", "minVersion": "1.1"
              }
            ]
          },
          "id": "commandsRuntime",
          "type": "general",
          "code": {
            "page": "https://contoso.com/commands.html",
            "script": "https://contoso.com/commands.js"
          },
          "lifetime": "short",
          "actions": [
            {
              "id": "action1",
              "type": "executeFunction"
            },
            {
              "id": "action2",
              "type": "executeFunction"
            },
            {
              "id": "action3",
              "type": "executeFunction"
            }
          ]
        }
      ],
      "ribbons": [
        {
          "contexts": [
            "mailCompose"
          ],
          "tabs": [
            {
              "builtInTabId": "TabDefault",
              "groups": [
                {
                  "id": "dashboard",
                  "label": "Controls",
                  "controls": [
                    {
                      "id": "control1",
                      "type": "button",
                      "label": "Action 1",
                      "icons": [
                        {
                          "size": 16,
                          "file": "test_16.png"
                        },
                        {
                          "size": 32,
                          "file": "test_32.png"
                        },
                        {
                          "size": 80,
                          "file": "test_80.png"
                        }
                      ],
                      "supertip": {
                        "title": "Action 1 Title",
                        "description": "Action 1 Description"
                      },
                      "actionId": "action1"
                    },
                    {
                      "id": "menu1",
                      "type": "menu",
                      "label": "My Menu",
                      "icons": [
                        {
                          "size": 16,
                          "file": "test_16.png"
                        },
                        {
                          "size": 32,
                          "file": "test_32.png"
                        },
                        {
                          "size": 80,
                          "file": "test_80.png"
                        }
                      ],
                      "supertip": {
                        "title": "My Menu",
                        "description": "Menu with 2 actions"
                      },
                      "items": [
                        {
                          "id": "menuItem1",
                          "type": "menuItem",
                          "label": "Action 2",
                          "supertip": {
                            "title": "Action 2 Title",
                            "description": "Action 2 Description"
                          },
                          "actionId": "action2"
                        },
                        {
                          "id": "menuItem2",
                          "type": "menuItem",
                          "label": "Action 3",
                          "icons": [
                            {
                              "size": 16,
                              "file": "test_16.png"
                            },
                            {
                              "size": 32,
                              "file": "test_32.png"
                            },
                            {
                              "size": 80,
                              "file": "test_80.png"
                            }
                          ],
                          "supertip": {
                            "title": "Action 3 Title",
                            "description": "Action 3 Description"
                          },
                          "actionId": "action3"
                        }
                      ]
                    }
                  ]
                }
              ]
            }
          ]
        },
        {
          "contexts": [ "mailRead" ],
          "tabs": [
            {
              "builtInTabId": "TabDefault",
              "groups": [
                {
                  "id": "dashboard",
                  "label": "Controls",
                  "controls": [
                    {
                      "id": "control1",
                      "type": "button",
                      "label": "Action 1",
                      "icons": [
                        {
                          "size": 16,
                          "file": "test_16.png"
                        },
                        {
                          "size": 32,
                          "file": "test_32.png"
                        },
                        {
                          "size": 80,
                          "file": "test_80.png"
                        }
                      ],
                      "supertip": {
                        "title": "Action 1 Title",
                        "description": "Action 1 Description"
                      },
                      "actionId": "action1"
                    }
                  ]
                }
              ]
            }
          ]
        }
      ],
      "autoRunEvents": [
        {
          "requirements": {
            "capabilities": [
              {
                "name": "MailBox", "minVersion": "1.10"
              }
            ]
          },
          "events": [
            {
              "type": "newMessageComposeCreated",
              "actionId": "onNewMessageComposeCreated"
            },
            {
              "type": "messageSending",
              "actionId": "onMessageSending",
              "options": {
                "sendMode": "promptUser"
              }
            }
          ]
        }
      ],
      "alternates": [
        {
          "requirements": {
            "scopes": [ "mail" ]
          },
          "prefer": {
            "comAddin": {
              "progId": "ContosoExtension"
            }
          },
          "hide": {
            "storeOfficeAddin": {
              "officeAddinId": "00000000-0000-0000-0000-000000000000",
              "assetId": "WA000000000"
            }
          }
        }
      ]
    }
  ]
}
```

## <a name="next-steps"></a>后续步骤

- [使用 Teams 清单（预览版）生成 Outlook 加载项](../quickstarts/outlook-quickstart-json-manifest.md)。