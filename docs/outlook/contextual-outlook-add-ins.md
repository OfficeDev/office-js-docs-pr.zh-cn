---
title: 上下文 Outlook 加载项
description: 无需离开邮件本身即可启用与邮件相关的任务，以此带来更便捷、更丰富的用户体验。
ms.date: 04/09/2020
localization_priority: Normal
ms.openlocfilehash: 7898f836e431ad4446952a0f34a24d3771e51d01
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/09/2021
ms.locfileid: "53348984"
---
# <a name="contextual-outlook-add-ins"></a><span data-ttu-id="c1e3e-103">上下文 Outlook 加载项</span><span class="sxs-lookup"><span data-stu-id="c1e3e-103">Contextual Outlook add-ins</span></span>

<span data-ttu-id="c1e3e-p101">上下文加载项是一些根据邮件或约会中的文本进行激活的 Outlook 外接程序。通过使用上下文加载项，用户无需离开邮件本身即可启动与邮件相关的任务，这会带来更便捷、更丰富的用户体验。</span><span class="sxs-lookup"><span data-stu-id="c1e3e-p101">Contextual add-ins are Outlook add-ins that activate based on text in a message or appointment. By using contextual add-ins, a user can initiate tasks related to a message without leaving the message itself, which results in an easier and richer user experience.</span></span>

<span data-ttu-id="c1e3e-106">下面是上下文外接程序的示例。</span><span class="sxs-lookup"><span data-stu-id="c1e3e-106">The following are examples of contextual add-ins.</span></span>

- <span data-ttu-id="c1e3e-107">选择地址以打开位置地图。</span><span class="sxs-lookup"><span data-stu-id="c1e3e-107">Choosing an address to open a map of the location.</span></span>
- <span data-ttu-id="c1e3e-108">选择会打开会议建议加载项的字符串。</span><span class="sxs-lookup"><span data-stu-id="c1e3e-108">Choosing a string that opens a meeting suggestion add-in.</span></span>
- <span data-ttu-id="c1e3e-109">选择要添加到你的联系人的电话号码。</span><span class="sxs-lookup"><span data-stu-id="c1e3e-109">Choosing a phone number to add to your contacts.</span></span>


> [!NOTE]
> <span data-ttu-id="c1e3e-110">上下文加载项暂不适用于 Android 版和 iOS 版 Outlook。</span><span class="sxs-lookup"><span data-stu-id="c1e3e-110">Contextual add-ins are not currently available in Outlook on Android and iOS.</span></span> <span data-ttu-id="c1e3e-111">今后将推出此功能。</span><span class="sxs-lookup"><span data-stu-id="c1e3e-111">This functionality will be made available in the future.</span></span>
>
> <span data-ttu-id="c1e3e-112">要求集1.6 中引入了对此功能的支持。</span><span class="sxs-lookup"><span data-stu-id="c1e3e-112">Support for this feature was introduced in requirement set 1.6.</span></span> <span data-ttu-id="c1e3e-113">请查看支持此要求集的[客户端和平台](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients)。</span><span class="sxs-lookup"><span data-stu-id="c1e3e-113">See [clients and platforms](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) that support this requirement set.</span></span>

## <a name="how-to-make-a-contextual-add-in"></a><span data-ttu-id="c1e3e-114">如何生成上下文加载项</span><span class="sxs-lookup"><span data-stu-id="c1e3e-114">How to make a contextual add-in</span></span>

<span data-ttu-id="c1e3e-115">上下文外接程序的清单必须包含将 `xsi:type` 属性设置为 `DetectedEntity` 的 [ExtensionPoint](../reference/manifest/extensionpoint.md#detectedentity) 元素。</span><span class="sxs-lookup"><span data-stu-id="c1e3e-115">A contextual add-in's manifest must include an [ExtensionPoint](../reference/manifest/extensionpoint.md#detectedentity) element with an `xsi:type` attribute set to `DetectedEntity`.</span></span> <span data-ttu-id="c1e3e-116">在 **ExtensionPoint** 元素中，该外接程序指定可以激活它的实体或正则表达式。</span><span class="sxs-lookup"><span data-stu-id="c1e3e-116">Within the **ExtensionPoint** element, the add-in specifies the entities or regular expression that can activate it.</span></span> <span data-ttu-id="c1e3e-117">如果指定实体，则该实体可以是 [Entities](/javascript/api/outlook/office.entities) 对象中的任何属性。</span><span class="sxs-lookup"><span data-stu-id="c1e3e-117">If an entity is specified, the entity can be any of the properties in the [Entities](/javascript/api/outlook/office.entities) object.</span></span>

<span data-ttu-id="c1e3e-118">因此，外接程序清单必须包含类型为 **ItemHasKnownEntity** 或 **ItemHasRegularExpressionMatch** 的规则。</span><span class="sxs-lookup"><span data-stu-id="c1e3e-118">Thus, the add-in manifest must contain a rule of type **ItemHasKnownEntity** or **ItemHasRegularExpressionMatch**.</span></span> <span data-ttu-id="c1e3e-119">以下示例演示如何指定外接程序应在检测到的实体为电话号码的邮件上激活。</span><span class="sxs-lookup"><span data-stu-id="c1e3e-119">The following example shows how to specify that an add-in should activate on messages with a detected entity that is a phone number.</span></span>

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

<span data-ttu-id="c1e3e-120">在上下文加载项与帐户关联后，当用户单击突出显示的实体或正则表达式时，加载项会自动启动。</span><span class="sxs-lookup"><span data-stu-id="c1e3e-120">After a contextual add-in is associated with an account, it will automatically start when the user clicks a highlighted entity or regular expression.</span></span> <span data-ttu-id="c1e3e-121">若要详细了解 Outlook 加载项正则表达式，请参阅[使用正则表达式激活规则显示 Outlook 加载项](use-regular-expressions-to-show-an-outlook-add-in.md)。</span><span class="sxs-lookup"><span data-stu-id="c1e3e-121">For more information about regular expressions for Outlook add-ins, see [Use regular expression activation rules to show an Outlook add-in](use-regular-expressions-to-show-an-outlook-add-in.md).</span></span>

<span data-ttu-id="c1e3e-122">上下文加载项有一些限制：</span><span class="sxs-lookup"><span data-stu-id="c1e3e-122">There are several restrictions on contextual add-ins:</span></span>

- <span data-ttu-id="c1e3e-123">上下文外接程序可以仅存在于阅读加载项中（而不是撰写加载项中）。</span><span class="sxs-lookup"><span data-stu-id="c1e3e-123">A contextual add-in can only exist in read add-ins (not compose add-ins).</span></span>
- <span data-ttu-id="c1e3e-124">不能指定突出显示的实体颜色。</span><span class="sxs-lookup"><span data-stu-id="c1e3e-124">You cannot specify the color of the highlighted entity.</span></span>
- <span data-ttu-id="c1e3e-125">未突出显示的实体无法启动卡片中的上下文外接程序。</span><span class="sxs-lookup"><span data-stu-id="c1e3e-125">An entity that is not highlighted will not launch a contextual add-in in a card.</span></span>

<span data-ttu-id="c1e3e-126">由于未突出显示的实体或正则表达式无法启动上下文外接程序，因此外接程序至少必须包含一个将 `Highlight` 属性设置为 `all` 的 `Rule` 元素。</span><span class="sxs-lookup"><span data-stu-id="c1e3e-126">Because an entity or regular expression that is not highlighted will not launch a contextual add-in, add-ins must include at least one `Rule` element with the `Highlight` attribute set to `all`.</span></span>

> [!NOTE]
> <span data-ttu-id="c1e3e-p107">`EmailAddress` 和 `Url` 实体类型不支持突出显示，因此它们不能用于启动上下文外接程序。但是，它们也可以组合在 `RuleCollection` 规则类型中作为其他激活条件。</span><span class="sxs-lookup"><span data-stu-id="c1e3e-p107">The `EmailAddress` and `Url` entity types do not support highlighting, so they cannot be used to launch a contextual add-in. They can however be combined in a `RuleCollection` rule type as an additional activation criteria.</span></span>

## <a name="how-to-launch-a-contextual-add-in"></a><span data-ttu-id="c1e3e-129">如何启动上下文外接程序</span><span class="sxs-lookup"><span data-stu-id="c1e3e-129">How to launch a contextual add-in</span></span>

<span data-ttu-id="c1e3e-p108">用户通过文本（可以是已知实体或开发人员的正则表达式）启动上下文外接程序。用户通常标识某个上下文外接程序的原因是该实体突出显示。以下示例说明如何使邮件中的内容突出显示。这里的实体（地址）是蓝色的，并带有蓝线虚线下划线。用户通过单击突出显示实体启动上下文外接程序。</span><span class="sxs-lookup"><span data-stu-id="c1e3e-p108">A user launches a contextual add-in through text, either a known entity or a developer's regular expression. Typically, a user identifies a contextual add-in because the entity is highlighted. The following example shows how highlighting appears in a message. Here the entity (an address) is colored blue and underlined with a dotted blue line. A user launches the contextual add-in by clicking the highlighted entity.</span></span> 

<span data-ttu-id="c1e3e-135">**含有突出显示实体（地址）的文本示例**</span><span class="sxs-lookup"><span data-stu-id="c1e3e-135">**Example of text with highlighted entity (an address)**</span></span>

![在电子邮件中显示突出显示的实体。](../images/outlook-detected-entity-highlight.png)
    
<span data-ttu-id="c1e3e-137">当邮件中含有多个实体或上下文外接程序时，用户交互规则如下所示：</span><span class="sxs-lookup"><span data-stu-id="c1e3e-137">When there are multiple entities or contextual add-ins in a message, there are a few user interaction rules:</span></span>

- <span data-ttu-id="c1e3e-138">如果有多个实体，用户必须单击不同的实体才能启动对应的外接程序。</span><span class="sxs-lookup"><span data-stu-id="c1e3e-138">If there are multiple entities, the user has to click a different entity to launch the add-in for it.</span></span>
- <span data-ttu-id="c1e3e-139">如果一个实体激活多个外接程序，则每个外接程序会打开一个新选项卡。用户可在选项卡之间切换，以在外接程序之间更改。例如，名称和地址可以触发电话外接程序和地图。</span><span class="sxs-lookup"><span data-stu-id="c1e3e-139">If an entity activates multiple add-ins, each add-in opens a new tab. The user switches between tabs to change between add-ins. For example, a name and address might trigger a phone add-in and a map.</span></span>
- <span data-ttu-id="c1e3e-p109">如果单个字符串中包含激活多个外接程序的多个实体，则整个字符串将突出显示，单击字符串可在单独的选项卡上显示与此字符串相关的所有外接程序。例如，表达建议在餐厅集会的字符串将激活"建议的会议"外接程序和餐厅评级外接程序。</span><span class="sxs-lookup"><span data-stu-id="c1e3e-p109">If a single string contains multiple entities that activate multiple add-ins, the entire string is highlighted, and clicking the string shows all add-ins relevant to the string on separate tabs. For example, a string that describes a proposed meeting at a restaurant might activate the Suggested Meeting add-in and a restaurant rating add-in.</span></span>

## <a name="how-a-contextual-add-in-displays"></a><span data-ttu-id="c1e3e-142">上下文外接程序的显示方式</span><span class="sxs-lookup"><span data-stu-id="c1e3e-142">How a contextual add-in displays</span></span>

<span data-ttu-id="c1e3e-p110">激活的上下文外接程序显示在卡片中，该卡片是靠近实体的单独窗口。该卡片通常会出现在实体下方，并尽可能地以实体为中心。如果实体下方没有足够的空间，则将卡片置于实体上方。以下屏幕截图显示了突出显示实体，并在其下方显示了卡片中激活的外接程序（必应地图）。</span><span class="sxs-lookup"><span data-stu-id="c1e3e-p110">An activated contextual add-in appears in a card, which is a separate window near the entity. The card will normally appear below the entity and centered with respect to the entity as much as possible. If there is not enough room below the entity, the card is placed above it. The following screenshot shows the highlighted entity, and below it, an activated add-in (Bing Maps) in a card.</span></span>

<span data-ttu-id="c1e3e-147">**显示在卡片中的外接程序示例**</span><span class="sxs-lookup"><span data-stu-id="c1e3e-147">**Example of an add-in displayed in a card**</span></span>

![在卡片中显示上下文相关应用。](../images/outlook-detected-entity-card.png)

<span data-ttu-id="c1e3e-149">若要关闭卡片并结束该外接程序，用户可单击该卡片外的任意位置。</span><span class="sxs-lookup"><span data-stu-id="c1e3e-149">To close the card and the add-in, a user clicks anywhere outside of the card.</span></span>

## <a name="current-contextual-add-ins"></a><span data-ttu-id="c1e3e-150">当前上下文外接程序</span><span class="sxs-lookup"><span data-stu-id="c1e3e-150">Current contextual add-ins</span></span>

<span data-ttu-id="c1e3e-151">默认情况下，会为使用加载项的用户安装Outlook加载项。</span><span class="sxs-lookup"><span data-stu-id="c1e3e-151">The following contextual add-ins are installed by default for users with Outlook add-ins.</span></span>

- <span data-ttu-id="c1e3e-152">必应地图</span><span class="sxs-lookup"><span data-stu-id="c1e3e-152">Bing Maps</span></span>
- <span data-ttu-id="c1e3e-153">建议的会议</span><span class="sxs-lookup"><span data-stu-id="c1e3e-153">Suggested Meetings</span></span>

## <a name="see-also"></a><span data-ttu-id="c1e3e-154">另请参阅</span><span class="sxs-lookup"><span data-stu-id="c1e3e-154">See also</span></span>

- <span data-ttu-id="c1e3e-155">[Outlook 加载项：Contoso 订单编号](https://github.com/OfficeDev/Outlook-Add-In-Contextual-Regex)（根据正则表达式匹配项激活的示例上下文加载项）</span><span class="sxs-lookup"><span data-stu-id="c1e3e-155">[Outlook add-in: Contoso Order Number](https://github.com/OfficeDev/Outlook-Add-In-Contextual-Regex) (sample contextual add-in that activates based on a regular expression match)</span></span>
- [<span data-ttu-id="c1e3e-156">编写第一个 Outlook 加载项</span><span class="sxs-lookup"><span data-stu-id="c1e3e-156">Write your first Outlook add-in</span></span>](../quickstarts/outlook-quickstart.md)
- [<span data-ttu-id="c1e3e-157">使用正则表达式激活规则显示 Outlook 外接程序</span><span class="sxs-lookup"><span data-stu-id="c1e3e-157">Use regular expression activation rules to show an Outlook add-in</span></span>](use-regular-expressions-to-show-an-outlook-add-in.md)
- [<span data-ttu-id="c1e3e-158">Entities 对象</span><span class="sxs-lookup"><span data-stu-id="c1e3e-158">Entities object</span></span>](/javascript/api/outlook/office.entities)
