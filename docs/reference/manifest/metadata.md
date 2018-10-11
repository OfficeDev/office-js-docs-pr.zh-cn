# <a name="metadata-element"></a><span data-ttu-id="f9adc-101">Metadata 元素</span><span class="sxs-lookup"><span data-stu-id="f9adc-101">MetaData element</span></span>

<span data-ttu-id="f9adc-102">定义 Excel 中的自定义函数所使用的元数据设置。</span><span class="sxs-lookup"><span data-stu-id="f9adc-102">Defines script settings used by a custom function in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="f9adc-103">属性</span><span class="sxs-lookup"><span data-stu-id="f9adc-103">Attributes</span></span>

<span data-ttu-id="f9adc-104">无</span><span class="sxs-lookup"><span data-stu-id="f9adc-104">None</span></span>

## <a name="child-elements"></a><span data-ttu-id="f9adc-105">子元素</span><span class="sxs-lookup"><span data-stu-id="f9adc-105">Child elements</span></span>

|  <span data-ttu-id="f9adc-106">元素</span><span class="sxs-lookup"><span data-stu-id="f9adc-106">Element</span></span>  |  <span data-ttu-id="f9adc-107">必需</span><span class="sxs-lookup"><span data-stu-id="f9adc-107">Required</span></span>  |  <span data-ttu-id="f9adc-108">说明</span><span class="sxs-lookup"><span data-stu-id="f9adc-108">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="f9adc-109">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="f9adc-109">SourceLocation</span></span>](customfunctionssourcelocation.md)  |  <span data-ttu-id="f9adc-110">是</span><span class="sxs-lookup"><span data-stu-id="f9adc-110">Yes</span></span>  | <span data-ttu-id="f9adc-111">包含自定义函数所使用的 JSON 文件的资源 ID 的字符串。</span><span class="sxs-lookup"><span data-stu-id="f9adc-111">String with the resource id of the HTML file used by custom functions.</span></span> |

## <a name="example"></a><span data-ttu-id="f9adc-112">示例</span><span class="sxs-lookup"><span data-stu-id="f9adc-112">Example</span></span>

```xml
<Metadata>
    <SourceLocation resid="JSON-URL" />
</Metadata>
```
