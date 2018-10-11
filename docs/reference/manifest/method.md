# <a name="method-element"></a><span data-ttu-id="cf293-101">Method 元素</span><span class="sxs-lookup"><span data-stu-id="cf293-101">Method element</span></span>

<span data-ttu-id="cf293-102">指定来自适用于 Office 的 JavaScript API 的单个方法，Office 加载项需要该方法才能激活。</span><span class="sxs-lookup"><span data-stu-id="cf293-102">Specifies an individual method from the JavaScript API for Office that your Office Add-in requires in order to activate.</span></span>

<span data-ttu-id="cf293-103">**加载项类型：** 内容、任务窗格</span><span class="sxs-lookup"><span data-stu-id="cf293-103">**Add-in type:** Content, Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="cf293-104">语法</span><span class="sxs-lookup"><span data-stu-id="cf293-104">Syntax</span></span>

```XML
<Method Name="string"/>
```

## <a name="contained-in"></a><span data-ttu-id="cf293-105">包含在</span><span class="sxs-lookup"><span data-stu-id="cf293-105">Contained in:</span></span>

[<span data-ttu-id="cf293-106">方法</span><span class="sxs-lookup"><span data-stu-id="cf293-106">Methods</span></span>](methods.md)

## <a name="attributes"></a><span data-ttu-id="cf293-107">属性</span><span class="sxs-lookup"><span data-stu-id="cf293-107">Attributes</span></span>

|<span data-ttu-id="cf293-108">**属性**</span><span class="sxs-lookup"><span data-stu-id="cf293-108">**Attribute**</span></span>|<span data-ttu-id="cf293-109">**类型**</span><span class="sxs-lookup"><span data-stu-id="cf293-109">**Type**</span></span>|<span data-ttu-id="cf293-110">**必需**</span><span class="sxs-lookup"><span data-stu-id="cf293-110">**Required**</span></span>|<span data-ttu-id="cf293-111">**说明**</span><span class="sxs-lookup"><span data-stu-id="cf293-111">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="cf293-112">名称</span><span class="sxs-lookup"><span data-stu-id="cf293-112">Name</span></span>|<span data-ttu-id="cf293-113">字符串</span><span class="sxs-lookup"><span data-stu-id="cf293-113">string</span></span>|<span data-ttu-id="cf293-114">必需</span><span class="sxs-lookup"><span data-stu-id="cf293-114">required</span></span>|<span data-ttu-id="cf293-p101">指定由其父对象限定的所需方法的名称。例如，要指定 **getSelectedDataAsync** 方法，必须指定 `"Document.getSelectedDataAsync"`。</span><span class="sxs-lookup"><span data-stu-id="cf293-p101">Specifies the name of the required method qualified with its parent object. For example, to specify the  **getSelectedDataAsync** method, you must specify `"Document.getSelectedDataAsync"`.</span></span>|

## <a name="remarks"></a><span data-ttu-id="cf293-117">备注</span><span class="sxs-lookup"><span data-stu-id="cf293-117">Remarks</span></span>

<span data-ttu-id="cf293-118">**Methods**  和 **Method** 元素不受邮件加载项的支持。有关要求集的详细信息，请参阅 [Office 版本和要求集](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets) 。</span><span class="sxs-lookup"><span data-stu-id="cf293-118">The  Methods and Method elements aren't supported by mail add-ins. For more information about requirement sets, see Specify Office hosts and API requirements.</span></span>

> [!IMPORTANT] 
> <span data-ttu-id="cf293-119">因为无法指定单个方法的最低版本要求，所以为了确保在运行时提供可用的方法，当你在加载项的脚本中调用该方法时，还应该使用 **if** 语句。</span><span class="sxs-lookup"><span data-stu-id="cf293-119">Important  Because there is no way to specify the minimum version requirement for individual methods, to make sure that a method is available at runtime, you should also use an  **if** statement when calling that method in the script of your add-in. For more information about how to do this, see Understanding the JavaScript API for Office.</span></span> <span data-ttu-id="cf293-120">有关如何执行此操作的详细信息，请参阅[了解适用于 Office 的 JavaScript API](https://docs.microsoft.com/office/dev/add-ins/develop/understanding-the-javascript-api-for-office)。</span><span class="sxs-lookup"><span data-stu-id="cf293-120">For more information about how to do this, see [Understanding the JavaScript API for Office](https://docs.microsoft.com/office/dev/add-ins/develop/understanding-the-javascript-api-for-office).</span></span>

