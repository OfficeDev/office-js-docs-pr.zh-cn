# <a name="supertip"></a><span data-ttu-id="47a7f-101">Supertip</span><span class="sxs-lookup"><span data-stu-id="47a7f-101">Supertip</span></span>

<span data-ttu-id="47a7f-p101">定义丰富的工具提示（标题和说明）。它由[按钮](control.md#button-control)或[菜单](control.md#menu-dropdown-button-controls)控件使用。</span><span class="sxs-lookup"><span data-stu-id="47a7f-p101">Defines a rich tooltip (both Title and Description). It is used by both [Button](control.md#button-control) or [Menu](control.md#menu-dropdown-button-controls)  controls.</span></span>

## <a name="child-elements"></a><span data-ttu-id="47a7f-104">子元素</span><span class="sxs-lookup"><span data-stu-id="47a7f-104">Child elements</span></span>

|  <span data-ttu-id="47a7f-105">元素</span><span class="sxs-lookup"><span data-stu-id="47a7f-105">Element</span></span> |  <span data-ttu-id="47a7f-106">必需</span><span class="sxs-lookup"><span data-stu-id="47a7f-106">Required</span></span>  |  <span data-ttu-id="47a7f-107">描述</span><span class="sxs-lookup"><span data-stu-id="47a7f-107">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="47a7f-108">标题</span><span class="sxs-lookup"><span data-stu-id="47a7f-108">Title</span></span>](#title)        | <span data-ttu-id="47a7f-109">是</span><span class="sxs-lookup"><span data-stu-id="47a7f-109">Yes</span></span> |   <span data-ttu-id="47a7f-110">Supertip 的文本。</span><span class="sxs-lookup"><span data-stu-id="47a7f-110">The text for the supertip.</span></span>         |
|  [<span data-ttu-id="47a7f-111">描述</span><span class="sxs-lookup"><span data-stu-id="47a7f-111">Description</span></span>](#description)  | <span data-ttu-id="47a7f-112">是</span><span class="sxs-lookup"><span data-stu-id="47a7f-112">Yes</span></span> |  <span data-ttu-id="47a7f-113">Supertip 的说明。</span><span class="sxs-lookup"><span data-stu-id="47a7f-113">The description for the supertip.</span></span>    |

### <a name="title"></a><span data-ttu-id="47a7f-114">标题</span><span class="sxs-lookup"><span data-stu-id="47a7f-114">Title</span></span>

<span data-ttu-id="47a7f-p102">必需。SuperTip 的文本。**resid** 属性必须设置为 **ShortStrings** 元素（位于 **Resources** 元素）中 **String** 元素的 [id](resources.md) 属性的值。</span><span class="sxs-lookup"><span data-stu-id="47a7f-p102">Required. The text for the supertip. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

### <a name="description"></a><span data-ttu-id="47a7f-118">描述</span><span class="sxs-lookup"><span data-stu-id="47a7f-118">Description</span></span>

<span data-ttu-id="47a7f-p103">必需。SuperTip 的描述。**resid** 属性必须设置为 **LongStrings** 元素（位于 **Resources** 元素）中 **String** 元素的 [id](resources.md) 属性的值。</span><span class="sxs-lookup"><span data-stu-id="47a7f-p103">Required. The description for the supertip. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **LongStrings** element in the [Resources](resources.md) element.</span></span>

## <a name="example"></a><span data-ttu-id="47a7f-122">示例</span><span class="sxs-lookup"><span data-stu-id="47a7f-122">Example</span></span>

```xml
<Supertip>
    <Title resid="funcReadSuperTipTitle" />
    <Description resid="funcReadSuperTipDescription" />
</Supertip>
```
