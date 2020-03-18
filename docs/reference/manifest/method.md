---
title: 清单文件中的 Method 元素
description: Method 元素指定 Office JavaScript API 中的单个方法，Office 外接程序需要这些方法才能激活。
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 5da25616d25a8d7454fc847727cda38a9935b5c7
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720580"
---
# <a name="method-element"></a><span data-ttu-id="1a6b8-103">Method 元素</span><span class="sxs-lookup"><span data-stu-id="1a6b8-103">Method element</span></span>

<span data-ttu-id="1a6b8-104">指定 office JavaScript API 中的单个方法，Office 外接程序需要这些方法才能激活。</span><span class="sxs-lookup"><span data-stu-id="1a6b8-104">Specifies an individual method from the Office JavaScript API that your Office Add-in requires in order to activate.</span></span>

<span data-ttu-id="1a6b8-105">**外接程序类型：** 内容、任务窗格</span><span class="sxs-lookup"><span data-stu-id="1a6b8-105">**Add-in type:** Content, Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="1a6b8-106">语法</span><span class="sxs-lookup"><span data-stu-id="1a6b8-106">Syntax</span></span>

```XML
<Method Name="string"/>
```

## <a name="contained-in"></a><span data-ttu-id="1a6b8-107">包含于</span><span class="sxs-lookup"><span data-stu-id="1a6b8-107">Contained in</span></span>

[<span data-ttu-id="1a6b8-108">Methods</span><span class="sxs-lookup"><span data-stu-id="1a6b8-108">Methods</span></span>](methods.md)

## <a name="attributes"></a><span data-ttu-id="1a6b8-109">属性</span><span class="sxs-lookup"><span data-stu-id="1a6b8-109">Attributes</span></span>

|<span data-ttu-id="1a6b8-110">**属性**</span><span class="sxs-lookup"><span data-stu-id="1a6b8-110">**Attribute**</span></span>|<span data-ttu-id="1a6b8-111">**类型**</span><span class="sxs-lookup"><span data-stu-id="1a6b8-111">**Type**</span></span>|<span data-ttu-id="1a6b8-112">**必需**</span><span class="sxs-lookup"><span data-stu-id="1a6b8-112">**Required**</span></span>|<span data-ttu-id="1a6b8-113">**说明**</span><span class="sxs-lookup"><span data-stu-id="1a6b8-113">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="1a6b8-114">名称</span><span class="sxs-lookup"><span data-stu-id="1a6b8-114">Name</span></span>|<span data-ttu-id="1a6b8-115">字符串</span><span class="sxs-lookup"><span data-stu-id="1a6b8-115">string</span></span>|<span data-ttu-id="1a6b8-116">必需</span><span class="sxs-lookup"><span data-stu-id="1a6b8-116">required</span></span>|<span data-ttu-id="1a6b8-117">指定由其父对象限定的所需方法的名称。</span><span class="sxs-lookup"><span data-stu-id="1a6b8-117">Specifies the name of the required method qualified with its parent object.</span></span> <span data-ttu-id="1a6b8-118">例如，若要指定`getSelectedDataAsync`方法，必须指定。 `"Document.getSelectedDataAsync"`</span><span class="sxs-lookup"><span data-stu-id="1a6b8-118">For example, to specify the `getSelectedDataAsync` method, you must specify `"Document.getSelectedDataAsync"`.</span></span>|

## <a name="remarks"></a><span data-ttu-id="1a6b8-119">备注</span><span class="sxs-lookup"><span data-stu-id="1a6b8-119">Remarks</span></span>

<span data-ttu-id="1a6b8-120">邮件`Methods`外`Method`接程序不支持和元素。有关要求集的详细信息，请参阅[Office 版本和要求集](../../develop/office-versions-and-requirement-sets.md)。</span><span class="sxs-lookup"><span data-stu-id="1a6b8-120">The `Methods` and `Method` elements aren't supported by mail add-ins. For more information about requirement sets, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="1a6b8-121">因为无法指定各个方法的最低版本要求，所以为了确保在运行时提供可用的方法，当您在外接程序的脚本中调用该方法时，还应该使用 **if** 语句。</span><span class="sxs-lookup"><span data-stu-id="1a6b8-121">Because there is no way to specify the minimum version requirement for individual methods, to make sure that a method is available at runtime, you should also use an **if** statement when calling that method in the script of your add-in.</span></span> <span data-ttu-id="1a6b8-122">有关如何执行此操作的详细信息，请参阅[了解 Office JAVASCRIPT API](../../develop/understanding-the-javascript-api-for-office.md)。</span><span class="sxs-lookup"><span data-stu-id="1a6b8-122">For more information about how to do this, see [Understanding the Office JavaScript API](../../develop/understanding-the-javascript-api-for-office.md).</span></span>
