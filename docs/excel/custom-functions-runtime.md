---
ms.date: 06/15/2022
description: 了解不使用共享运行时及其特定 JavaScript 运行时的 Excel 自定义函数。
title: 自定义函数的仅 JavaScript 运行时
ms.localizationpriority: medium
ms.openlocfilehash: 0d3298e95ab39f976c3fbfd5c0cc4ecdd1369721
ms.sourcegitcommit: b6a3815a1ad17f3522ca35247a3fd5d7105e174e
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/22/2022
ms.locfileid: "66958409"
---
# <a name="javascript-only-runtime-for-custom-functions"></a>自定义函数的仅 JavaScript 运行时

不使用共享运行时的自定义函数使用仅限 JavaScript 的运行时，该运行时旨在优化计算性能。

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

[!include[Shared runtime note](../includes/shared-runtime-note.md)]

此 JavaScript 运行时提供对命名空间中 `OfficeRuntime` API 的访问权限，这些 API 可由自定义函数使用，任务窗格 (在不同的运行时) 中运行以存储数据。

## <a name="request-external-data"></a>请求外部数据

在自定义函数中，你可以使用 [Fetch](https://developer.mozilla.org/docs/Web/API/Fetch_API) 等 API 或使用 [XmlHttpRequest (XHR)](https://developer.mozilla.org/docs/Web/API/XMLHttpRequest)（一种发出与服务器交互的 HTTP 请求的标准 Web API）来请求外部数据。

请注意，在创建 XmlHttpRequests 时，自定义函数必须使用其他安全措施，这需要 [相同的源](https://developer.mozilla.org/docs/Web/Security/Same-origin_policy) 策略和简单的 [CORS](https://www.w3.org/TR/cors/)。

简单的 CORS 实现不能使用 Cookie，并且仅支持 GET、HEAD、POST)  (简单方法。 简单的 CORS 接受字段名称为 `Accept`、`Accept-Language`、`Content-Language` 的简单标题。 还可以在简单 CORS 中使用`Content-Type`标头，前提是内容类型为`application/x-www-form-urlencoded`或 `text/plain``multipart/form-data`。

## <a name="store-and-access-data"></a>存储和访问数据

在不使用共享运行时的自定义函数中，可以使用 [OfficeRuntime.storage](/javascript/api/office-runtime/officeruntime.storage) 对象存储和访问数据。 该 `Storage` 对象是一个永久性的、未加密的密钥值存储系统，它提供了 [localStorage](https://developer.mozilla.org/docs/Web/API/Window/localStorage) 的替代方法，使用仅限 JavaScript 的运行时的自定义函数无法使用该存储。 该 `Storage` 对象为每个域提供 10 MB 的数据。 域可由多个加载项共享。

该 `Storage` 对象是共享存储解决方案，这意味着加载项的多个部分能够访问相同的数据。 例如，用户身份验证的令牌可能存储在对象中 `Storage` ，因为使用仅限 JavaScript 的运行时) 的自定义函数 (和使用完整 Web 视图运行时)  (的任务窗格都可以访问它。 同样，如果两个加载项共享相同的域 (例如 `www.contoso.com/addin1``www.contoso.com/addin2` ，) ，则还允许它们通过`Storage`对象来回共享信息。 请注意，具有不同子域的加载项将具有不同的`Storage` (实例， `subdomain.contoso.com/addin1``differentsubdomain.contoso.com/addin2` 例如，) 。

`Storage`由于该对象可以是共享位置，因此请务必认识到，可以替代键值对。

对象上 `Storage` 提供了以下方法。

- `getItem`
- `getItems`
- `setItem`
- `setItems`
- `removeItem`
- `removeItems`
- `getKeys`

> [!NOTE]
> 无法清除所有信息 (（如 `clear`) ）。 相反，需要使用 `removeItems` 来一次性删除多个条目。

### <a name="officeruntimestorage-example"></a>OfficeRuntime.storage 示例

下面的 `OfficeRuntime.storage.setItem` 代码示例调用将键和值设置为 `storage`的方法。

```js
function StoreValue(key, value) {

  return OfficeRuntime.storage.setItem(key, value).then(function (result) {
      return "Success: Item with key '" + key + "' saved to storage.";
  }, function (error) {
      return "Error: Unable to save item with key '" + key + "' to storage. " + error;
  });
}
```

## <a name="next-steps"></a>后续步骤

了解如何 [调试自定义函数](custom-functions-debugging.md)。

## <a name="see-also"></a>另请参阅

- [没有共享运行时的自定义函数的身份验证](custom-functions-authentication.md)
- [在 Excel 中创建自定义函数](custom-functions-overview.md)
- [自定义函数教程](../tutorials/excel-tutorial-create-custom-functions.md)
