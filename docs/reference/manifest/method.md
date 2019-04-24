---
title: 清单文件中的 Method 元素
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 19234b35e1faf8a8cc52a9e893fcc720793cadae
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450651"
---
# <a name="method-element"></a><span data-ttu-id="5a1d8-102">Method 元素</span><span class="sxs-lookup"><span data-stu-id="5a1d8-102">Method element</span></span>

<span data-ttu-id="5a1d8-103">指定来自适用于 Office 的 JavaScript API 的单个方法，Office 外接程序需要该方法才能激活。</span><span class="sxs-lookup"><span data-stu-id="5a1d8-103">Specifies an individual method from the JavaScript API for Office that your Office Add-in requires in order to activate.</span></span>

<span data-ttu-id="5a1d8-104">**外接程序类型：** 内容、任务窗格</span><span class="sxs-lookup"><span data-stu-id="5a1d8-104">**Add-in type:** Content, Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="5a1d8-105">语法</span><span class="sxs-lookup"><span data-stu-id="5a1d8-105">Syntax</span></span>

```XML
<Method Name="string"/>
```

## <a name="contained-in"></a><span data-ttu-id="5a1d8-106">包含于</span><span class="sxs-lookup"><span data-stu-id="5a1d8-106">Contained in</span></span>

[<span data-ttu-id="5a1d8-107">Methods</span><span class="sxs-lookup"><span data-stu-id="5a1d8-107">Methods</span></span>](methods.md)

## <a name="attributes"></a><span data-ttu-id="5a1d8-108">属性</span><span class="sxs-lookup"><span data-stu-id="5a1d8-108">Attributes</span></span>

|<span data-ttu-id="5a1d8-109">**属性**</span><span class="sxs-lookup"><span data-stu-id="5a1d8-109">**Attribute**</span></span>|<span data-ttu-id="5a1d8-110">**类型**</span><span class="sxs-lookup"><span data-stu-id="5a1d8-110">**Type**</span></span>|<span data-ttu-id="5a1d8-111">**必需**</span><span class="sxs-lookup"><span data-stu-id="5a1d8-111">**Required**</span></span>|<span data-ttu-id="5a1d8-112">**说明**</span><span class="sxs-lookup"><span data-stu-id="5a1d8-112">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="5a1d8-113">名称</span><span class="sxs-lookup"><span data-stu-id="5a1d8-113">Name</span></span>|<span data-ttu-id="5a1d8-114">字符串</span><span class="sxs-lookup"><span data-stu-id="5a1d8-114">string</span></span>|<span data-ttu-id="5a1d8-115">必需</span><span class="sxs-lookup"><span data-stu-id="5a1d8-115">required</span></span>|<span data-ttu-id="5a1d8-p101">指定由其父对象限定的所需方法的名称。例如，要指定 **getSelectedDataAsync** 方法，必须指定 `"Document.getSelectedDataAsync"`。</span><span class="sxs-lookup"><span data-stu-id="5a1d8-p101">Specifies the name of the required method qualified with its parent object. For example, to specify the  **getSelectedDataAsync** method, you must specify `"Document.getSelectedDataAsync"`.</span></span>|

## <a name="remarks"></a><span data-ttu-id="5a1d8-118">注释</span><span class="sxs-lookup"><span data-stu-id="5a1d8-118">Remarks</span></span>

<span data-ttu-id="5a1d8-119">**Methods** 和 **Method** 元素不受邮件外接程序的支持。有关要求集的详细信息，请参阅 [Office 版本和要求集](/office/dev/add-ins/develop/office-versions-and-requirement-sets)。</span><span class="sxs-lookup"><span data-stu-id="5a1d8-119">The  **Methods** and **Method** elements aren't supported by mail add-ins. For more information about requirement sets, see [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

> [!IMPORTANT] 
> <span data-ttu-id="5a1d8-120">因为无法指定各个方法的最低版本要求，所以为了确保在运行时提供可用的方法，当您在外接程序的脚本中调用该方法时，还应该使用 **if** 语句。</span><span class="sxs-lookup"><span data-stu-id="5a1d8-120">Because there is no way to specify the minimum version requirement for individual methods, to make sure that a method is available at runtime, you should also use an **if** statement when calling that method in the script of your add-in.</span></span> <span data-ttu-id="5a1d8-121">有关如何执行此操作的详细信息，请参阅[了解适用于 Office 的 JavaScript API](/office/dev/add-ins/develop/understanding-the-javascript-api-for-office)。</span><span class="sxs-lookup"><span data-stu-id="5a1d8-121">For more information about how to do this, see [Understanding the JavaScript API for Office](/office/dev/add-ins/develop/understanding-the-javascript-api-for-office).</span></span>

