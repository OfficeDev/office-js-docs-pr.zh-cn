---
title: 上下文 Outlook 加载项
description: 无需离开邮件本身即可启用与邮件相关的任务，以此带来更便捷、更丰富的用户体验。
ms.date: 04/09/2020
ms.localizationpriority: medium
ms.openlocfilehash: 83ed12c2bb8c61ba25db93c321406a1d43de4594
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/12/2021
ms.locfileid: "59152419"
---
# <a name="contextual-outlook-add-ins"></a>上下文 Outlook 加载项

上下文加载项是一些根据邮件或约会中的文本进行激活的 Outlook 外接程序。通过使用上下文加载项，用户无需离开邮件本身即可启动与邮件相关的任务，这会带来更便捷、更丰富的用户体验。

下面是上下文外接程序的示例。

- 选择地址以打开位置地图。
- 选择会打开会议建议加载项的字符串。
- 选择要添加到你的联系人的电话号码。


> [!NOTE]
> 上下文加载项暂不适用于 Android 版和 iOS 版 Outlook。 今后将推出此功能。
>
> 要求集1.6 中引入了对此功能的支持。 请查看支持此要求集的[客户端和平台](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients)。

## <a name="how-to-make-a-contextual-add-in"></a>如何生成上下文加载项

上下文外接程序的清单必须包含将 `xsi:type` 属性设置为 `DetectedEntity` 的 [ExtensionPoint](../reference/manifest/extensionpoint.md#detectedentity) 元素。 在 **ExtensionPoint** 元素中，该外接程序指定可以激活它的实体或正则表达式。 如果指定实体，则该实体可以是 [Entities](/javascript/api/outlook/office.entities) 对象中的任何属性。

因此，外接程序清单必须包含类型为 **ItemHasKnownEntity** 或 **ItemHasRegularExpressionMatch** 的规则。 以下示例演示如何指定外接程序应在检测到的实体为电话号码的邮件上激活。

```XML
<ExtensionPoint xsi:type="DetectedEntity">
  <Label resid="contextLabel" />
  <!--If you opt to include RequestedHeight, it must be between 140px to 450px, inclusive.-->
  <!--<RequestedHeight>360</RequestedHeight>-->
  <SourceLocation resid="detectedEntityURL" />
  <Rule xsi:type="RuleCollection" Mode="And">
    <Rule xsi:type="ItemIs" ItemType="Message" />
    <Rule xsi:type="ItemHasKnownEntity" EntityType="PhoneNumber" Highlight="all" />
  </Rule>
</ExtensionPoint>
```

在上下文加载项与帐户关联后，当用户单击突出显示的实体或正则表达式时，加载项会自动启动。 若要详细了解 Outlook 加载项正则表达式，请参阅[使用正则表达式激活规则显示 Outlook 加载项](use-regular-expressions-to-show-an-outlook-add-in.md)。

上下文加载项有一些限制：

- 上下文外接程序可以仅存在于阅读加载项中（而不是撰写加载项中）。
- 不能指定突出显示的实体颜色。
- 未突出显示的实体无法启动卡片中的上下文外接程序。

由于未突出显示的实体或正则表达式无法启动上下文外接程序，因此外接程序至少必须包含一个将 `Highlight` 属性设置为 `all` 的 `Rule` 元素。

> [!NOTE]
> `EmailAddress` 和 `Url` 实体类型不支持突出显示，因此它们不能用于启动上下文外接程序。但是，它们也可以组合在 `RuleCollection` 规则类型中作为其他激活条件。

## <a name="how-to-launch-a-contextual-add-in"></a>如何启动上下文外接程序

用户通过文本（可以是已知实体或开发人员的正则表达式）启动上下文外接程序。用户通常标识某个上下文外接程序的原因是该实体突出显示。以下示例说明如何使邮件中的内容突出显示。这里的实体（地址）是蓝色的，并带有蓝线虚线下划线。用户通过单击突出显示实体启动上下文外接程序。 

**含有突出显示实体（地址）的文本示例**

![在电子邮件中显示突出显示的实体。](../images/outlook-detected-entity-highlight.png)
    
当邮件中含有多个实体或上下文外接程序时，用户交互规则如下所示：

- 如果有多个实体，用户必须单击不同的实体才能启动对应的外接程序。
- 如果一个实体激活多个外接程序，则每个外接程序会打开一个新选项卡。用户可在选项卡之间切换，以在外接程序之间更改。例如，名称和地址可以触发电话外接程序和地图。
- 如果单个字符串中包含激活多个外接程序的多个实体，则整个字符串将突出显示，单击字符串可在单独的选项卡上显示与此字符串相关的所有外接程序。例如，表达建议在餐厅集会的字符串将激活"建议的会议"外接程序和餐厅评级外接程序。

## <a name="how-a-contextual-add-in-displays"></a>上下文外接程序的显示方式

激活的上下文外接程序显示在卡片中，该卡片是靠近实体的单独窗口。该卡片通常会出现在实体下方，并尽可能地以实体为中心。如果实体下方没有足够的空间，则将卡片置于实体上方。以下屏幕截图显示了突出显示实体，并在其下方显示了卡片中激活的外接程序（必应地图）。

**显示在卡片中的外接程序示例**

![在卡片中显示上下文相关应用。](../images/outlook-detected-entity-card.png)

若要关闭卡片并结束该外接程序，用户可单击该卡片外的任意位置。

## <a name="current-contextual-add-ins"></a>当前上下文外接程序

默认情况下，会为具有外接程序的用户安装Outlook加载项。

- 必应地图
- 建议的会议

## <a name="see-also"></a>另请参阅

- [Outlook 加载项：Contoso 订单编号](https://github.com/OfficeDev/Outlook-Add-In-Contextual-Regex)（根据正则表达式匹配项激活的示例上下文加载项）
- [编写第一个 Outlook 加载项](../quickstarts/outlook-quickstart.md)
- [使用正则表达式激活规则显示 Outlook 外接程序](use-regular-expressions-to-show-an-outlook-add-in.md)
- [Entities 对象](/javascript/api/outlook/office.entities)
