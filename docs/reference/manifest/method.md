---
title: 清单文件中的 Method 元素
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: fded84344182bb45597b00a794f18defaa44d3b3
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432821"
---
# <a name="method-element"></a><span data-ttu-id="f9bb4-102">Method 元素</span><span class="sxs-lookup"><span data-stu-id="f9bb4-102">Method element</span></span>

<span data-ttu-id="f9bb4-103">指定来自适用于 Office 的 JavaScript API 的单个方法，Office 外接程序需要该方法才能激活。</span><span class="sxs-lookup"><span data-stu-id="f9bb4-103">Specifies an individual method from the JavaScript API for Office that your Office Add-in requires in order to activate.</span></span>

<span data-ttu-id="f9bb4-104">**外接程序类型：** 内容、任务窗格</span><span class="sxs-lookup"><span data-stu-id="f9bb4-104">**Add-in type:** Content, Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="f9bb4-105">语法</span><span class="sxs-lookup"><span data-stu-id="f9bb4-105">Syntax</span></span>

```XML
<Method Name="string"/>
```

## <a name="contained-in"></a><span data-ttu-id="f9bb4-106">包含于</span><span class="sxs-lookup"><span data-stu-id="f9bb4-106">Contained in</span></span>

[<span data-ttu-id="f9bb4-107">Methods</span><span class="sxs-lookup"><span data-stu-id="f9bb4-107">Methods</span></span>](methods.md)

## <a name="attributes"></a><span data-ttu-id="f9bb4-108">属性</span><span class="sxs-lookup"><span data-stu-id="f9bb4-108">Attributes</span></span>

|<span data-ttu-id="f9bb4-109">**属性**</span><span class="sxs-lookup"><span data-stu-id="f9bb4-109">**Attribute**</span></span>|<span data-ttu-id="f9bb4-110">**类型**</span><span class="sxs-lookup"><span data-stu-id="f9bb4-110">**Type**</span></span>|<span data-ttu-id="f9bb4-111">**必需**</span><span class="sxs-lookup"><span data-stu-id="f9bb4-111">**Required**</span></span>|<span data-ttu-id="f9bb4-112">**说明**</span><span class="sxs-lookup"><span data-stu-id="f9bb4-112">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="f9bb4-113">名称</span><span class="sxs-lookup"><span data-stu-id="f9bb4-113">Name</span></span>|<span data-ttu-id="f9bb4-114">字符串</span><span class="sxs-lookup"><span data-stu-id="f9bb4-114">string</span></span>|<span data-ttu-id="f9bb4-115">必需</span><span class="sxs-lookup"><span data-stu-id="f9bb4-115">required</span></span>|<span data-ttu-id="f9bb4-p101">指定由其父对象限定的所需方法的名称。例如，要指定 **getSelectedDataAsync** 方法，必须指定 `"Document.getSelectedDataAsync"`。</span><span class="sxs-lookup"><span data-stu-id="f9bb4-p101">Specifies the name of the required method qualified with its parent object. For example, to specify the  **getSelectedDataAsync** method, you must specify `"Document.getSelectedDataAsync"`.</span></span>|

## <a name="remarks"></a><span data-ttu-id="f9bb4-118">注释</span><span class="sxs-lookup"><span data-stu-id="f9bb4-118">Remarks</span></span>

<span data-ttu-id="f9bb4-119">**Methods** 和 **Method** 元素不受邮件外接程序的支持。有关要求集的详细信息，请参阅 [Office 版本和要求集](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)。</span><span class="sxs-lookup"><span data-stu-id="f9bb4-119">The  Methods and Method elements aren't supported by mail add-ins. For more information about requirement sets, see Specify Office hosts and API requirements.</span></span>

> [!IMPORTANT] 
> <span data-ttu-id="f9bb4-120">因为无法指定各个方法的最低版本要求，所以为了确保在运行时提供可用的方法，当您在外接程序的脚本中调用该方法时，还应该使用 **if** 语句。</span><span class="sxs-lookup"><span data-stu-id="f9bb4-120">Important  Because there is no way to specify the minimum version requirement for individual methods, to make sure that a method is available at runtime, you should also use an  **if** statement when calling that method in the script of your add-in. For more information about how to do this, see Understanding the JavaScript API for Office.</span></span> <span data-ttu-id="f9bb4-121">有关如何执行此操作的详细信息，请参阅[了解适用于 Office 的 JavaScript API](https://docs.microsoft.com/office/dev/add-ins/develop/understanding-the-javascript-api-for-office)。</span><span class="sxs-lookup"><span data-stu-id="f9bb4-121">For more information about how to do this, see [Understanding the JavaScript API for Office](https://docs.microsoft.com/office/dev/add-ins/develop/understanding-the-javascript-api-for-office).</span></span>

