---
ms.date: 07/08/2021
description: 了解Excel窗格及其特定 JavaScript 运行时的自定义函数。
title: 无 UI 的运行时Excel自定义函数
ms.localizationpriority: medium
ms.openlocfilehash: 491e47674d87d99d0adeda952ee65ffc24dff2bd
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/12/2021
ms.locfileid: "59148859"
---
# <a name="runtime-for-ui-less-excel-custom-functions"></a>无 UI 的运行时Excel自定义函数

不使用任务窗格的自定义函数 (无 UI 的自定义) 使用旨在优化计算性能的 JavaScript 运行时。

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

[!include[Shared runtime note](../includes/shared-runtime-note.md)]

此 JavaScript 运行时提供对命名空间中的 API 的访问权限，无 UI 自定义函数和任务窗格可以使用这些 API `OfficeRuntime` 来存储数据。

## <a name="request-external-data"></a>请求外部数据

在无 UI 自定义函数中，可以使用提取等 API 或[XmlHttpRequest (XHR) （](https://developer.mozilla.org/docs/Web/API/XMLHttpRequest)一种发布 HTTP 请求以与服务器交互的标准 Web API）请求外部数据。 [](https://developer.mozilla.org/docs/Web/API/Fetch_API)

请注意，无 UI 函数在生成 XmlHttpRequest 时必须使用额外的安全措施，这需要[同源策略](https://developer.mozilla.org/docs/Web/Security/Same-origin_policy)和简单的[CORS。](https://www.w3.org/TR/cors/)

简单的 CORS 实现不能使用 Cookie，并且仅支持 GET、HEAD、POST () 的简单方法。 简单的 CORS 接受字段名称为 `Accept`、`Accept-Language`、`Content-Language` 的简单标题。 还可以在简单 `Content-Type` CORS 中使用标头，只要内容类型为 、 或 `application/x-www-form-urlencoded` `text/plain` `multipart/form-data` 。

## <a name="store-and-access-data"></a>存储和访问数据

在无 UI 自定义函数中，可以使用 对象存储和访问 `OfficeRuntime.storage` 数据。 `Storage` 是一个持续、未加密的键值存储系统，可提供 [localStorage](https://developer.mozilla.org/docs/Web/API/Window/localStorage)的替代项，而无 UI 自定义函数不能使用它。 `Storage` 每个域提供 10 MB 的数据。 域可以由多个加载项共享。

`Storage` 旨在作为共享存储解决方案，这意味着外接程序的多个部分将能访问相同数据。 例如，用户身份验证令牌可能存储在 中，因为它可以通过无 UI 自定义函数和外接程序 UI 元素（如任务窗格） `storage` 访问。 同样，如果两个加载项共享同一个域 (例如 ，、) ，则还允许它们通过 来回 `www.contoso.com/addin1` `www.contoso.com/addin2` 共享信息 `storage` 。 请注意，具有不同子域的加载项将具有不同的 (`storage` 例如 `subdomain.contoso.com/addin1` `differentsubdomain.contoso.com/addin2` ，) 。

由于 `storage` 可能是共享的位置，因此一定要认识到，可能会存在替代键值对的情况。

以下方法在 对象 `storage` 上可用。

- `getItem`
- `getItems`
- `setItem`
- `setItems`
- `removeItem`
- `removeItems`
- `getKeys`

> [!NOTE]
> 没有用于清除所有信息的方法 (例如 `clear`) 。 相反，需要使用 `removeItems` 来一次性删除多个条目。

### <a name="officeruntimestorage-example"></a>OfficeRuntime.storage 示例

下面的代码示例调用 `OfficeRuntime.storage.setItem` 函数，将键和值设置为 `storage` 。

```js
function StoreValue(key, value) {

  return OfficeRuntime.storage.setItem(key, value).then(function (result) {
      return "Success: Item with key '" + key + "' saved to storage.";
  }, function (error) {
      return "Error: Unable to save item with key '" + key + "' to storage. " + error;
  });
}
```

## <a name="additional-considerations"></a>其他注意事项

如果加载项仅使用无 UI 自定义函数，请注意，不能通过无 UI 自定义函数访问文档对象模型 (DOM) 或使用 jQuery 等依赖于 DOM 的库。

## <a name="next-steps"></a>后续步骤

了解如何调试 [无 UI 自定义函数](custom-functions-debugging.md)。

## <a name="see-also"></a>另请参阅

* [对无 UI 自定义函数进行身份验证](custom-functions-authentication.md)
* [在 Excel 中创建自定义函数](custom-functions-overview.md)
* [自定义函数教程](../tutorials/excel-tutorial-create-custom-functions.md)
