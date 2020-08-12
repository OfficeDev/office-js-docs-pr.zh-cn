---
title: 清单文件中的 Method 元素
description: Method 元素指定 Office JavaScript API 中的单个方法，Office 外接程序需要这些方法才能激活。
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 0e3e74a73a3422a7789e82d6f0e7a516bd795ca8
ms.sourcegitcommit: cc6886b47c84ac37a3c957ff85dd0ed526ca5e43
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/12/2020
ms.locfileid: "46641323"
---
# <a name="method-element"></a><span data-ttu-id="09097-103">Method 元素</span><span class="sxs-lookup"><span data-stu-id="09097-103">Method element</span></span>

<span data-ttu-id="09097-104">指定 office JavaScript API 中的单个方法，Office 外接程序需要这些方法才能激活。</span><span class="sxs-lookup"><span data-stu-id="09097-104">Specifies an individual method from the Office JavaScript API that your Office Add-in requires in order to activate.</span></span>

<span data-ttu-id="09097-105">**外接程序类型：** 内容、任务窗格</span><span class="sxs-lookup"><span data-stu-id="09097-105">**Add-in type:** Content, Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="09097-106">语法</span><span class="sxs-lookup"><span data-stu-id="09097-106">Syntax</span></span>

```XML
<Method Name="string"/>
```

## <a name="contained-in"></a><span data-ttu-id="09097-107">包含于</span><span class="sxs-lookup"><span data-stu-id="09097-107">Contained in</span></span>

[<span data-ttu-id="09097-108">Methods</span><span class="sxs-lookup"><span data-stu-id="09097-108">Methods</span></span>](methods.md)

## <a name="attributes"></a><span data-ttu-id="09097-109">属性</span><span class="sxs-lookup"><span data-stu-id="09097-109">Attributes</span></span>

|<span data-ttu-id="09097-110">属性</span><span class="sxs-lookup"><span data-stu-id="09097-110">Attribute</span></span>|<span data-ttu-id="09097-111">类型</span><span class="sxs-lookup"><span data-stu-id="09097-111">Type</span></span>|<span data-ttu-id="09097-112">必需</span><span class="sxs-lookup"><span data-stu-id="09097-112">Required</span></span>|<span data-ttu-id="09097-113">说明</span><span class="sxs-lookup"><span data-stu-id="09097-113">Description</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="09097-114">Name</span><span class="sxs-lookup"><span data-stu-id="09097-114">Name</span></span>|<span data-ttu-id="09097-115">字符串</span><span class="sxs-lookup"><span data-stu-id="09097-115">string</span></span>|<span data-ttu-id="09097-116">必需</span><span class="sxs-lookup"><span data-stu-id="09097-116">required</span></span>|<span data-ttu-id="09097-117">指定由其父对象限定的所需方法的名称。</span><span class="sxs-lookup"><span data-stu-id="09097-117">Specifies the name of the required method qualified with its parent object.</span></span> <span data-ttu-id="09097-118">例如，若要指定 `getSelectedDataAsync` 方法，必须指定 `"Document.getSelectedDataAsync"` 。</span><span class="sxs-lookup"><span data-stu-id="09097-118">For example, to specify the `getSelectedDataAsync` method, you must specify `"Document.getSelectedDataAsync"`.</span></span>|

## <a name="remarks"></a><span data-ttu-id="09097-119">说明</span><span class="sxs-lookup"><span data-stu-id="09097-119">Remarks</span></span>

<span data-ttu-id="09097-120">`Methods` `Method` 邮件外接程序不支持和元素。有关要求集的详细信息，请参阅[Office 版本和要求集](../../develop/office-versions-and-requirement-sets.md)。</span><span class="sxs-lookup"><span data-stu-id="09097-120">The `Methods` and `Method` elements aren't supported by mail add-ins. For more information about requirement sets, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="09097-121">因为无法指定各个方法的最低版本要求，所以为了确保在运行时提供可用的方法，当您在外接程序的脚本中调用该方法时，还应该使用 **if** 语句。</span><span class="sxs-lookup"><span data-stu-id="09097-121">Because there is no way to specify the minimum version requirement for individual methods, to make sure that a method is available at runtime, you should also use an **if** statement when calling that method in the script of your add-in.</span></span> <span data-ttu-id="09097-122">有关如何执行此操作的详细信息，请参阅[了解 Office JAVASCRIPT API](../../develop/understanding-the-javascript-api-for-office.md)。</span><span class="sxs-lookup"><span data-stu-id="09097-122">For more information about how to do this, see [Understanding the Office JavaScript API](../../develop/understanding-the-javascript-api-for-office.md).</span></span>
