---
title: 清单文件中的 Method 元素
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 74b7a8b3d0f8511d21eb0df150500850e8b93fe9
ms.sourcegitcommit: 4079903c3cc45b7d8c041509a44e9fc38da399b1
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/11/2020
ms.locfileid: "42596891"
---
# <a name="method-element"></a><span data-ttu-id="75433-102">Method 元素</span><span class="sxs-lookup"><span data-stu-id="75433-102">Method element</span></span>

<span data-ttu-id="75433-103">指定 office JavaScript API 中的单个方法，Office 外接程序需要这些方法才能激活。</span><span class="sxs-lookup"><span data-stu-id="75433-103">Specifies an individual method from the Office JavaScript API that your Office Add-in requires in order to activate.</span></span>

<span data-ttu-id="75433-104">**外接程序类型：** 内容、任务窗格</span><span class="sxs-lookup"><span data-stu-id="75433-104">**Add-in type:** Content, Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="75433-105">语法</span><span class="sxs-lookup"><span data-stu-id="75433-105">Syntax</span></span>

```XML
<Method Name="string"/>
```

## <a name="contained-in"></a><span data-ttu-id="75433-106">包含于</span><span class="sxs-lookup"><span data-stu-id="75433-106">Contained in</span></span>

[<span data-ttu-id="75433-107">Methods</span><span class="sxs-lookup"><span data-stu-id="75433-107">Methods</span></span>](methods.md)

## <a name="attributes"></a><span data-ttu-id="75433-108">属性</span><span class="sxs-lookup"><span data-stu-id="75433-108">Attributes</span></span>

|<span data-ttu-id="75433-109">**属性**</span><span class="sxs-lookup"><span data-stu-id="75433-109">**Attribute**</span></span>|<span data-ttu-id="75433-110">**类型**</span><span class="sxs-lookup"><span data-stu-id="75433-110">**Type**</span></span>|<span data-ttu-id="75433-111">**必需**</span><span class="sxs-lookup"><span data-stu-id="75433-111">**Required**</span></span>|<span data-ttu-id="75433-112">**说明**</span><span class="sxs-lookup"><span data-stu-id="75433-112">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="75433-113">名称</span><span class="sxs-lookup"><span data-stu-id="75433-113">Name</span></span>|<span data-ttu-id="75433-114">字符串</span><span class="sxs-lookup"><span data-stu-id="75433-114">string</span></span>|<span data-ttu-id="75433-115">必需</span><span class="sxs-lookup"><span data-stu-id="75433-115">required</span></span>|<span data-ttu-id="75433-116">指定由其父对象限定的所需方法的名称。</span><span class="sxs-lookup"><span data-stu-id="75433-116">Specifies the name of the required method qualified with its parent object.</span></span> <span data-ttu-id="75433-117">例如，若要指定`getSelectedDataAsync`方法，必须指定。 `"Document.getSelectedDataAsync"`</span><span class="sxs-lookup"><span data-stu-id="75433-117">For example, to specify the `getSelectedDataAsync` method, you must specify `"Document.getSelectedDataAsync"`.</span></span>|

## <a name="remarks"></a><span data-ttu-id="75433-118">说明</span><span class="sxs-lookup"><span data-stu-id="75433-118">Remarks</span></span>

<span data-ttu-id="75433-119">邮件`Methods`外`Method`接程序不支持和元素。有关要求集的详细信息，请参阅[Office 版本和要求集](../../develop/office-versions-and-requirement-sets.md)。</span><span class="sxs-lookup"><span data-stu-id="75433-119">The `Methods` and `Method` elements aren't supported by mail add-ins. For more information about requirement sets, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="75433-120">因为无法指定各个方法的最低版本要求，所以为了确保在运行时提供可用的方法，当您在外接程序的脚本中调用该方法时，还应该使用 **if** 语句。</span><span class="sxs-lookup"><span data-stu-id="75433-120">Because there is no way to specify the minimum version requirement for individual methods, to make sure that a method is available at runtime, you should also use an **if** statement when calling that method in the script of your add-in.</span></span> <span data-ttu-id="75433-121">有关如何执行此操作的详细信息，请参阅[了解 Office JAVASCRIPT API](../../develop/understanding-the-javascript-api-for-office.md)。</span><span class="sxs-lookup"><span data-stu-id="75433-121">For more information about how to do this, see [Understanding the Office JavaScript API](../../develop/understanding-the-javascript-api-for-office.md).</span></span>
