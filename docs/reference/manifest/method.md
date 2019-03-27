---
title: 清单文件中的 Method 元素
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 19234b35e1faf8a8cc52a9e893fcc720793cadae
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/27/2019
ms.locfileid: "30870420"
---
# <a name="method-element"></a><span data-ttu-id="fedb3-102">Method 元素</span><span class="sxs-lookup"><span data-stu-id="fedb3-102">Method element</span></span>

<span data-ttu-id="fedb3-103">指定来自适用于 Office 的 JavaScript API 的单个方法，Office 外接程序需要该方法才能激活。</span><span class="sxs-lookup"><span data-stu-id="fedb3-103">Specifies an individual method from the JavaScript API for Office that your Office Add-in requires in order to activate.</span></span>

<span data-ttu-id="fedb3-104">**外接程序类型：** 内容、任务窗格</span><span class="sxs-lookup"><span data-stu-id="fedb3-104">**Add-in type:** Content, Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="fedb3-105">语法</span><span class="sxs-lookup"><span data-stu-id="fedb3-105">Syntax</span></span>

```XML
<Method Name="string"/>
```

## <a name="contained-in"></a><span data-ttu-id="fedb3-106">包含于</span><span class="sxs-lookup"><span data-stu-id="fedb3-106">Contained in</span></span>

[<span data-ttu-id="fedb3-107">Methods</span><span class="sxs-lookup"><span data-stu-id="fedb3-107">Methods</span></span>](methods.md)

## <a name="attributes"></a><span data-ttu-id="fedb3-108">属性</span><span class="sxs-lookup"><span data-stu-id="fedb3-108">Attributes</span></span>

|<span data-ttu-id="fedb3-109">**属性**</span><span class="sxs-lookup"><span data-stu-id="fedb3-109">**Attribute**</span></span>|<span data-ttu-id="fedb3-110">**类型**</span><span class="sxs-lookup"><span data-stu-id="fedb3-110">**Type**</span></span>|<span data-ttu-id="fedb3-111">**必需**</span><span class="sxs-lookup"><span data-stu-id="fedb3-111">**Required**</span></span>|<span data-ttu-id="fedb3-112">**说明**</span><span class="sxs-lookup"><span data-stu-id="fedb3-112">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="fedb3-113">名称</span><span class="sxs-lookup"><span data-stu-id="fedb3-113">Name</span></span>|<span data-ttu-id="fedb3-114">字符串</span><span class="sxs-lookup"><span data-stu-id="fedb3-114">string</span></span>|<span data-ttu-id="fedb3-115">必需</span><span class="sxs-lookup"><span data-stu-id="fedb3-115">required</span></span>|<span data-ttu-id="fedb3-p101">指定由其父对象限定的所需方法的名称。例如，要指定 **getSelectedDataAsync** 方法，必须指定 `"Document.getSelectedDataAsync"`。</span><span class="sxs-lookup"><span data-stu-id="fedb3-p101">Specifies the name of the required method qualified with its parent object. For example, to specify the  **getSelectedDataAsync** method, you must specify `"Document.getSelectedDataAsync"`.</span></span>|

## <a name="remarks"></a><span data-ttu-id="fedb3-118">注释</span><span class="sxs-lookup"><span data-stu-id="fedb3-118">Remarks</span></span>

<span data-ttu-id="fedb3-119">**Methods** 和 **Method** 元素不受邮件外接程序的支持。有关要求集的详细信息，请参阅 [Office 版本和要求集](/office/dev/add-ins/develop/office-versions-and-requirement-sets)。</span><span class="sxs-lookup"><span data-stu-id="fedb3-119">The  **Methods** and **Method** elements aren't supported by mail add-ins. For more information about requirement sets, see [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

> [!IMPORTANT] 
> <span data-ttu-id="fedb3-120">因为无法指定各个方法的最低版本要求，所以为了确保在运行时提供可用的方法，当您在外接程序的脚本中调用该方法时，还应该使用 **if** 语句。</span><span class="sxs-lookup"><span data-stu-id="fedb3-120">Because there is no way to specify the minimum version requirement for individual methods, to make sure that a method is available at runtime, you should also use an **if** statement when calling that method in the script of your add-in.</span></span> <span data-ttu-id="fedb3-121">有关如何执行此操作的详细信息，请参阅[了解适用于 Office 的 JavaScript API](/office/dev/add-ins/develop/understanding-the-javascript-api-for-office)。</span><span class="sxs-lookup"><span data-stu-id="fedb3-121">For more information about how to do this, see [Understanding the JavaScript API for Office](/office/dev/add-ins/develop/understanding-the-javascript-api-for-office).</span></span>

