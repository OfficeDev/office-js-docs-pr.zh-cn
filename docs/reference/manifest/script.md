# <a name="script-element"></a><span data-ttu-id="5df1e-101">Script 元素</span><span class="sxs-lookup"><span data-stu-id="5df1e-101">Script element</span></span>

<span data-ttu-id="5df1e-102">定义 Excel 中的自定义函数所使用的脚本设置。</span><span class="sxs-lookup"><span data-stu-id="5df1e-102">Defines script settings used by a custom function in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="5df1e-103">属性</span><span class="sxs-lookup"><span data-stu-id="5df1e-103">Attributes</span></span>

<span data-ttu-id="5df1e-104">None</span><span class="sxs-lookup"><span data-stu-id="5df1e-104">None</span></span>

## <a name="child-elements"></a><span data-ttu-id="5df1e-105">子元素</span><span class="sxs-lookup"><span data-stu-id="5df1e-105">Child elements</span></span>

|<span data-ttu-id="5df1e-106">元素</span><span class="sxs-lookup"><span data-stu-id="5df1e-106">Elements</span></span>  |  <span data-ttu-id="5df1e-107">必需</span><span class="sxs-lookup"><span data-stu-id="5df1e-107">Required</span></span>  |  <span data-ttu-id="5df1e-108">说明</span><span class="sxs-lookup"><span data-stu-id="5df1e-108">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="5df1e-109">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="5df1e-109">SourceLocation</span></span>](customfunctionssourcelocation.md)  |  <span data-ttu-id="5df1e-110">是</span><span class="sxs-lookup"><span data-stu-id="5df1e-110">Yes</span></span>  | <span data-ttu-id="5df1e-111">自定义函数所使用的 JavaScript 文件的资源 id 的字符串。</span><span class="sxs-lookup"><span data-stu-id="5df1e-111">String with resource id of the JavaScript file used by custom functions.</span></span>|

## <a name="example"></a><span data-ttu-id="5df1e-112">示例</span><span class="sxs-lookup"><span data-stu-id="5df1e-112">Example</span></span>

```xml
<Script>
    <SourceLocation resid="scriptURL" />
    <!-- The Script element is not used in the Developer Preview. -->
</Script>
```
