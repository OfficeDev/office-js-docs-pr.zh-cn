# <a name="event-element"></a><span data-ttu-id="fc8f1-101">Event 元素</span><span class="sxs-lookup"><span data-stu-id="fc8f1-101">Event element</span></span>

<span data-ttu-id="fc8f1-102">定义加载项中的事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="fc8f1-102">Defines an event handler in an add-in.</span></span>

> [!NOTE] 
> <span data-ttu-id="fc8f1-103">目前，Office 365 中的  Outlook 网页版仅支持 `Event` 元素。</span><span class="sxs-lookup"><span data-stu-id="fc8f1-103">Note: The `Event` element is currently only supported by Outlook on the web in Office 365.</span></span>

## <a name="attributes"></a><span data-ttu-id="fc8f1-104">属性</span><span class="sxs-lookup"><span data-stu-id="fc8f1-104">Attributes</span></span>

|  <span data-ttu-id="fc8f1-105">属性</span><span class="sxs-lookup"><span data-stu-id="fc8f1-105">Attribute</span></span>  |  <span data-ttu-id="fc8f1-106">必需</span><span class="sxs-lookup"><span data-stu-id="fc8f1-106">Required</span></span>  |  <span data-ttu-id="fc8f1-107">说明</span><span class="sxs-lookup"><span data-stu-id="fc8f1-107">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="fc8f1-108">类型</span><span class="sxs-lookup"><span data-stu-id="fc8f1-108">Type</span></span>](#type-attribute)  |  <span data-ttu-id="fc8f1-109">是</span><span class="sxs-lookup"><span data-stu-id="fc8f1-109">Yes</span></span>  | <span data-ttu-id="fc8f1-110">指定要处理的事件。</span><span class="sxs-lookup"><span data-stu-id="fc8f1-110">Specifies the event to handle.</span></span> |
|  [<span data-ttu-id="fc8f1-111">FunctionExecution</span><span class="sxs-lookup"><span data-stu-id="fc8f1-111">FunctionExecution</span></span>](#functionexecution-attribute)  |  <span data-ttu-id="fc8f1-112">是</span><span class="sxs-lookup"><span data-stu-id="fc8f1-112">Yes</span></span>  | <span data-ttu-id="fc8f1-p101">指定事件处理程序的执行风格、异步或同步。目前仅支持同步事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="fc8f1-p101">Specifies the execution style for the event handler, asynchronous or synchronous. Currently only synchronous event handlers are supported.</span></span> |
|  [<span data-ttu-id="fc8f1-115">FunctionName</span><span class="sxs-lookup"><span data-stu-id="fc8f1-115">FunctionName</span></span>](#functionname-attribute)  |  <span data-ttu-id="fc8f1-116">是</span><span class="sxs-lookup"><span data-stu-id="fc8f1-116">Yes</span></span>  | <span data-ttu-id="fc8f1-117">指定事件处理程序的函数名称。</span><span class="sxs-lookup"><span data-stu-id="fc8f1-117">Specifies the function name for the event handler.</span></span> |

### <a name="type-attribute"></a><span data-ttu-id="fc8f1-118">类型属性</span><span class="sxs-lookup"><span data-stu-id="fc8f1-118">Type attribute</span></span>

<span data-ttu-id="fc8f1-p102">必需。指定哪个事件将调用该事件处理程序。此属性的可能值在下表中指定。</span><span class="sxs-lookup"><span data-stu-id="fc8f1-p102">Required. Specifies which event will invoke the event handler. The possible values for this attribute are specified in the following table.</span></span>

|  <span data-ttu-id="fc8f1-122">事件类型</span><span class="sxs-lookup"><span data-stu-id="fc8f1-122">Event type</span></span>  |  <span data-ttu-id="fc8f1-123">说明</span><span class="sxs-lookup"><span data-stu-id="fc8f1-123">Description</span></span>  |
|:-----|:-----|
|  `ItemSend`  |  <span data-ttu-id="fc8f1-124">当用户发送邮件或会议邀请时将调用该事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="fc8f1-124">The event handler will be invoked when the user sends a message or meeting invitation.</span></span>  |

### <a name="functionexecution-attribute"></a><span data-ttu-id="fc8f1-125">FunctionExecution 属性</span><span class="sxs-lookup"><span data-stu-id="fc8f1-125">FunctionExecution attribute</span></span>

<span data-ttu-id="fc8f1-126">必需。</span><span class="sxs-lookup"><span data-stu-id="fc8f1-126">Required.</span></span> <span data-ttu-id="fc8f1-127">必须设置为 `synchronous`。</span><span class="sxs-lookup"><span data-stu-id="fc8f1-127">MUST be set to `synchronous`.</span></span>

### <a name="functionname-attribute"></a><span data-ttu-id="fc8f1-128">FunctionName 属性</span><span class="sxs-lookup"><span data-stu-id="fc8f1-128">FunctionName attribute</span></span>

<span data-ttu-id="fc8f1-p104">必需。指定事件处理程序的函数名称。这个值必须与加载项的[函数文件](functionfile.md)中的函数名称相匹配。</span><span class="sxs-lookup"><span data-stu-id="fc8f1-p104">Required. Specifies the function name of the event handler. This value must match a function name in the add-in's [function file](functionfile.md).</span></span>

```xml
<Event Type="ItemSend" FunctionExecution="synchronous" FunctionName="itemSendHandler" /> 
```