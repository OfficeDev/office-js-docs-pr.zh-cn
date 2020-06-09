---
title: 清单文件中的 Method 元素
description: Method 元素指定 Office JavaScript API 中的单个方法，Office 外接程序需要这些方法才能激活。
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: c3531475a920fd24ce8390170b5f4728d4dcd0e0
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611755"
---
# <a name="method-element"></a><span data-ttu-id="c7218-103">Method 元素</span><span class="sxs-lookup"><span data-stu-id="c7218-103">Method element</span></span>

<span data-ttu-id="c7218-104">指定 office JavaScript API 中的单个方法，Office 外接程序需要这些方法才能激活。</span><span class="sxs-lookup"><span data-stu-id="c7218-104">Specifies an individual method from the Office JavaScript API that your Office Add-in requires in order to activate.</span></span>

<span data-ttu-id="c7218-105">**外接程序类型：** 内容、任务窗格</span><span class="sxs-lookup"><span data-stu-id="c7218-105">**Add-in type:** Content, Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="c7218-106">语法</span><span class="sxs-lookup"><span data-stu-id="c7218-106">Syntax</span></span>

```XML
<Method Name="string"/>
```

## <a name="contained-in"></a><span data-ttu-id="c7218-107">包含于</span><span class="sxs-lookup"><span data-stu-id="c7218-107">Contained in</span></span>

[<span data-ttu-id="c7218-108">Methods</span><span class="sxs-lookup"><span data-stu-id="c7218-108">Methods</span></span>](methods.md)

## <a name="attributes"></a><span data-ttu-id="c7218-109">属性</span><span class="sxs-lookup"><span data-stu-id="c7218-109">Attributes</span></span>

|<span data-ttu-id="c7218-110">**属性**</span><span class="sxs-lookup"><span data-stu-id="c7218-110">**Attribute**</span></span>|<span data-ttu-id="c7218-111">**类型**</span><span class="sxs-lookup"><span data-stu-id="c7218-111">**Type**</span></span>|<span data-ttu-id="c7218-112">**必需**</span><span class="sxs-lookup"><span data-stu-id="c7218-112">**Required**</span></span>|<span data-ttu-id="c7218-113">**说明**</span><span class="sxs-lookup"><span data-stu-id="c7218-113">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="c7218-114">名称</span><span class="sxs-lookup"><span data-stu-id="c7218-114">Name</span></span>|<span data-ttu-id="c7218-115">字符串</span><span class="sxs-lookup"><span data-stu-id="c7218-115">string</span></span>|<span data-ttu-id="c7218-116">必需</span><span class="sxs-lookup"><span data-stu-id="c7218-116">required</span></span>|<span data-ttu-id="c7218-117">指定由其父对象限定的所需方法的名称。</span><span class="sxs-lookup"><span data-stu-id="c7218-117">Specifies the name of the required method qualified with its parent object.</span></span> <span data-ttu-id="c7218-118">例如，若要指定 `getSelectedDataAsync` 方法，必须指定 `"Document.getSelectedDataAsync"` 。</span><span class="sxs-lookup"><span data-stu-id="c7218-118">For example, to specify the `getSelectedDataAsync` method, you must specify `"Document.getSelectedDataAsync"`.</span></span>|

## <a name="remarks"></a><span data-ttu-id="c7218-119">备注</span><span class="sxs-lookup"><span data-stu-id="c7218-119">Remarks</span></span>

<span data-ttu-id="c7218-120">`Methods` `Method` 邮件外接程序不支持和元素。有关要求集的详细信息，请参阅[Office 版本和要求集](../../develop/office-versions-and-requirement-sets.md)。</span><span class="sxs-lookup"><span data-stu-id="c7218-120">The `Methods` and `Method` elements aren't supported by mail add-ins. For more information about requirement sets, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="c7218-121">因为无法指定各个方法的最低版本要求，所以为了确保在运行时提供可用的方法，当您在外接程序的脚本中调用该方法时，还应该使用 **if** 语句。</span><span class="sxs-lookup"><span data-stu-id="c7218-121">Because there is no way to specify the minimum version requirement for individual methods, to make sure that a method is available at runtime, you should also use an **if** statement when calling that method in the script of your add-in.</span></span> <span data-ttu-id="c7218-122">有关如何执行此操作的详细信息，请参阅[了解 Office JAVASCRIPT API](../../develop/understanding-the-javascript-api-for-office.md)。</span><span class="sxs-lookup"><span data-stu-id="c7218-122">For more information about how to do this, see [Understanding the Office JavaScript API](../../develop/understanding-the-javascript-api-for-office.md).</span></span>
