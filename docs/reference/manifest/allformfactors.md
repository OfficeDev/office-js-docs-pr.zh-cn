# <a name="allformfactors-element"></a><span data-ttu-id="ce514-101">AllFormFactors 元素</span><span class="sxs-lookup"><span data-stu-id="ce514-101">AllFormFactors element</span></span>

<span data-ttu-id="ce514-102">指定加载项的所有外观设置。</span><span class="sxs-lookup"><span data-stu-id="ce514-102">Specifies the settings for an add-in for all form factors.</span></span> <span data-ttu-id="ce514-103">目前，使用  **AllFormFactors** 的唯一功能是自定义函数。</span><span class="sxs-lookup"><span data-stu-id="ce514-103">Currently, the only feature using AllFormFactors is custom functions.</span></span> <span data-ttu-id="ce514-104">**AllFormFactors** 是使用自定义函数时的必要元素。</span><span class="sxs-lookup"><span data-stu-id="ce514-104">AllFormFactors is a required element when using custom functions.</span></span>

## <a name="child-elements"></a><span data-ttu-id="ce514-105">子元素</span><span class="sxs-lookup"><span data-stu-id="ce514-105">Child elements</span></span>

|  <span data-ttu-id="ce514-106">元素</span><span class="sxs-lookup"><span data-stu-id="ce514-106">Element</span></span> |  <span data-ttu-id="ce514-107">必需</span><span class="sxs-lookup"><span data-stu-id="ce514-107">Required</span></span>  |  <span data-ttu-id="ce514-108">说明</span><span class="sxs-lookup"><span data-stu-id="ce514-108">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="ce514-109">ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="ce514-109">ExtensionPoint</span></span>](extensionpoint.md) |  <span data-ttu-id="ce514-110">是</span><span class="sxs-lookup"><span data-stu-id="ce514-110">Yes</span></span> |  <span data-ttu-id="ce514-111">定义外接程序公开功能的位置。</span><span class="sxs-lookup"><span data-stu-id="ce514-111">Defines where an add-in exposes functionality.</span></span> |

## <a name="allformfactors-example"></a><span data-ttu-id="ce514-112">AllFormFactors 示例</span><span class="sxs-lookup"><span data-stu-id="ce514-112">AllFormFactors example</span></span>

```xml
<Hosts>
    <Host xsi:type="Workbook">
        <AllFormFactors>
            <ExtensionPoint xsi:type="CustomFunctions">
                    <!-- Information on this extension point -->
            </ExtensionPoint>
        </AllFormFactors>
    </Host>
</Hosts>
```
