---
ms.date: 02/08/2019
description: 了解 Excel 自定义函数名称的要求并避免出现常见命名缺陷。
title: Excel 中自定义函数的命名准则 (预览)
localization_priority: Normal
ms.openlocfilehash: 954753c35d2df59093661e3b8e92adfa1302e595
ms.sourcegitcommit: 8e7b7b0cfb68b91a3a95585d094cf5f5ffd00178
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/09/2019
ms.locfileid: "30512837"
---
# <a name="naming-guidelines"></a><span data-ttu-id="958d6-103">命名准则</span><span class="sxs-lookup"><span data-stu-id="958d6-103">Naming guidelines</span></span>

<span data-ttu-id="958d6-104">自定义函数由 JSON 元数据文件中的**id**和**name**属性标识。</span><span class="sxs-lookup"><span data-stu-id="958d6-104">A custom function is identified by an **id** and **name** property in the JSON metadata file.</span></span> <span data-ttu-id="958d6-105">函数 id 用于唯一标识 JavaScript 代码中的自定义函数。</span><span class="sxs-lookup"><span data-stu-id="958d6-105">The function id is used to uniquely identify custom functions in your JavaScript code.</span></span> <span data-ttu-id="958d6-106">函数名称将用作在 Excel 中向用户显示的显示名称。</span><span class="sxs-lookup"><span data-stu-id="958d6-106">The function name is used as the display name that appears to a user in Excel.</span></span> <span data-ttu-id="958d6-107">函数名可以与函数 ID 不同, 例如出于本地化目的。</span><span class="sxs-lookup"><span data-stu-id="958d6-107">A function name can differ from the function ID, such as for localization purposes.</span></span> <span data-ttu-id="958d6-108">但通常, 如果没有理由让它们不同, 则应将其保持与 ID 相同。</span><span class="sxs-lookup"><span data-stu-id="958d6-108">But in general it should stay the same as the ID if there is no compelling reason for them to differ.</span></span>

<span data-ttu-id="958d6-109">函数名称和函数 id 共享一些常见要求:</span><span class="sxs-lookup"><span data-stu-id="958d6-109">Function names and function IDs share some common requirements:</span></span>

- <span data-ttu-id="958d6-110">函数 id 可能仅使用从 A 到 Z 的字符, 从零到九、下划线和数个句点。</span><span class="sxs-lookup"><span data-stu-id="958d6-110">Function ids may only use characters A through Z, numbers zero through nine, underscores, and periods.</span></span>

- <span data-ttu-id="958d6-111">函数名称可以使用任何 Unicode 字母字符、下划线和句点。</span><span class="sxs-lookup"><span data-stu-id="958d6-111">Function names may use any Unicode alphabetic characters, underscores, and periods.</span></span>

- <span data-ttu-id="958d6-112">它们必须以字母开头, 最小限制为三个字符。</span><span class="sxs-lookup"><span data-stu-id="958d6-112">They must start with a letter and have a minimum limit of three characters.</span></span>

<span data-ttu-id="958d6-113">Excel 使用大写字母作为内置函数名称 (例如`SUM`)。</span><span class="sxs-lookup"><span data-stu-id="958d6-113">Excel uses uppercase letters for built-in function names (such as `SUM`).</span></span> <span data-ttu-id="958d6-114">因此, 请考虑将大写字母用作自定义函数名称和函数 id 作为最佳实践。</span><span class="sxs-lookup"><span data-stu-id="958d6-114">Therefore, consider using uppercase letters for your custom function names and function IDs as a best practice.</span></span>

<span data-ttu-id="958d6-115">函数名称不应按如下方式命名:</span><span class="sxs-lookup"><span data-stu-id="958d6-115">Function names shouldn't be named the same as:</span></span>

- <span data-ttu-id="958d6-116">A1 到 XFD1048576 之间的任何单元格, 或从 R1C1 到 R1048576C16384 之间的任何单元格。</span><span class="sxs-lookup"><span data-stu-id="958d6-116">Any cells between A1 to XFD1048576 or any cells between R1C1 to R1048576C16384.</span></span>

- <span data-ttu-id="958d6-117">任何 Excel 4.0 宏函数 (例如`RUN`, `ECHO`)。</span><span class="sxs-lookup"><span data-stu-id="958d6-117">Any Excel 4.0 Macro Function (such as `RUN`, `ECHO`).</span></span>  <span data-ttu-id="958d6-118">有关这些函数的完整列表, 请参阅[本文](https://www.microsoft.com/en-us/download/details.aspx?id=1465)。</span><span class="sxs-lookup"><span data-stu-id="958d6-118">For a full list of these functions, see [this article](https://www.microsoft.com/en-us/download/details.aspx?id=1465).</span></span>

## <a name="naming-conflicts"></a><span data-ttu-id="958d6-119">命名冲突</span><span class="sxs-lookup"><span data-stu-id="958d6-119">Naming conflicts</span></span>

<span data-ttu-id="958d6-120">如果您的函数名称与已存在的外接程序中的函数名称相同, 则 **#REF!**</span><span class="sxs-lookup"><span data-stu-id="958d6-120">If your function name is the same as a function name in an add-in that already exists, the **#REF!**</span></span> <span data-ttu-id="958d6-121">错误将出现在工作簿中。</span><span class="sxs-lookup"><span data-stu-id="958d6-121">error will appear in your workbook.</span></span>

<span data-ttu-id="958d6-122">若要修复名称冲突, 请更改外接程序中的名称, 然后重试该函数。</span><span class="sxs-lookup"><span data-stu-id="958d6-122">To fix a name conflict, change the name in your add-in and try the function again.</span></span> <span data-ttu-id="958d6-123">此外, 还可以使用冲突的名称卸载加载项。</span><span class="sxs-lookup"><span data-stu-id="958d6-123">You can also uninstall the add-in with the conflicting name.</span></span> <span data-ttu-id="958d6-124">或者, 如果要在不同的环境中测试外接程序, 请尝试使用不同的命名空间来区分您的函数 (如 NAMESPACE_NAMEOFFUNCTION)。</span><span class="sxs-lookup"><span data-stu-id="958d6-124">Or, if you're testing your add-in in different environments, try using a different namespace to differentiate your function (such as NAMESPACE_NAMEOFFUNCTION).</span></span>

<span data-ttu-id="958d6-125">此外, 还应考虑你希望用户在你的外接程序中使用这些功能的方式。</span><span class="sxs-lookup"><span data-stu-id="958d6-125">Also consider how you'd like people to use the functions within your add-in.</span></span> <span data-ttu-id="958d6-126">在许多情况下, 将多个参数添加到函数中是有意义的, 而不是使用相同或相似的名称来创建多个函数。</span><span class="sxs-lookup"><span data-stu-id="958d6-126">In many cases, it makes sense to add multiple arguments to a function rather than create multiple functions with the same or similar names.</span></span>

## <a name="see-also"></a><span data-ttu-id="958d6-127">另请参阅</span><span class="sxs-lookup"><span data-stu-id="958d6-127">See also</span></span>

* [<span data-ttu-id="958d6-128">自定义函数元数据</span><span class="sxs-lookup"><span data-stu-id="958d6-128">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="958d6-129">自定义函数最佳实践</span><span class="sxs-lookup"><span data-stu-id="958d6-129">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="958d6-130">Excel 自定义函数教程</span><span class="sxs-lookup"><span data-stu-id="958d6-130">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
* [<span data-ttu-id="958d6-131">Excel 自定义函数的运行时</span><span class="sxs-lookup"><span data-stu-id="958d6-131">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
