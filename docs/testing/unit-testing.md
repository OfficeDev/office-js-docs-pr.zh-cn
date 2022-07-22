---
title: Office 加载项中的单元测试
description: 了解如何对调用 Office JavaScript API 的代码进行单元测试。
ms.date: 02/07/2022
ms.localizationpriority: medium
ms.openlocfilehash: 21858a68734ca5d07621f3e9c88b147ebac7dde6
ms.sourcegitcommit: b6a3815a1ad17f3522ca35247a3fd5d7105e174e
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/22/2022
ms.locfileid: "66958746"
---
# <a name="unit-testing-in-office-add-ins"></a>Office 加载项中的单元测试

单元测试检查外接程序的功能，而无需网络或服务连接，包括与 Office 应用程序的连接。 单元测试服务器端代码和 *不* 调用 [Office JavaScript API](../develop/understanding-the-javascript-api-for-office.md) 的客户端代码在 Office 加载项中与在任何 Web 应用程序中相同，因此不需要特殊文档。 但是调用 Office JavaScript API 的客户端代码很难进行测试。 为了解决这些问题，我们创建了一个库，以简化在单元测试中创建模拟 Office 对象： [Office-Addin-Mock](https://www.npmjs.com/package/office-addin-mock)。 该库通过以下方式简化了测试：

- Office JavaScript API 必须在 Office 应用程序 (Excel、Word 等) 上下文的 Web 视图控件中初始化，因此无法在开发计算机上运行单元测试的进程中加载它们。 可将 Office-Addin-Mock 库导入到测试文件中，从而在运行测试的node.js进程中模拟 Office JavaScript API。
- [特定于应用程序的 API](../develop/understanding-the-javascript-api-for-office.md#api-models) 具有[负载](../develop/application-specific-api-model.md#load)和[同步](../develop/application-specific-api-model.md#sync)方法，必须按相对于其他函数和彼此的特定顺序调用这些方法。 此外，必须使用某些参数调用该方法， `load` 具体取决于要在 *稍后* 测试的函数中，代码将读入哪些 Office 对象的属性。 但单元测试框架本质上是无状态的，因此它们不能记录是`load``sync`调用还是被调用，或者传递给`load`哪些参数。 使用 Office-Addin-Mock 库创建的模拟对象具有跟踪这些内容的内部状态。 这使模拟对象能够模拟实际 Office 对象的错误行为。 例如，如果正在测试的函数尝试读取未首先传递到 `load`的属性，则测试将返回类似于 Office 将返回的错误。

该库不依赖于 Office JavaScript API，并且可以与任何 JavaScript 单元测试框架一起使用，例如：

- [开玩笑](https://jestjs.io)
- [摩 卡](https://mochajs.org/)
- [茉莉花](https://jasmine.github.io/)

本文中的示例使用 Jest 框架。 [Office-Addin-Mock 主页](https://github.com/OfficeDev/Office-Addin-Scripts/tree/master/packages/office-addin-mock#examples)上有使用 Mocha 框架的示例。

## <a name="prerequisites"></a>先决条件

本文假设你熟悉单元测试和模拟的基本概念，包括如何创建和运行测试文件，以及你在单元测试框架方面有一些经验。

> [!TIP]
> 如果使用 Visual Studio，建议阅读 Visual [Studio 中“单元测试 JavaScript”和“TypeScript”](/visualstudio/javascript/unit-testing-javascript-with-visual-studio) 一文，了解有关 Visual Studio 中 JavaScript 单元测试的一些基本信息，然后返回到本文。

## <a name="install-the-tool"></a>安装工具

若要安装库，请打开命令提示符，导航到加载项项目的根目录，然后输入以下命令。

```command&nbsp;line
npm install office-addin-mock --save-dev
```

## <a name="basic-usage"></a>基本使用情况

1. 你的项目将具有一个或多个测试文件。  (请参阅下面) 的示例 (#示例中测试框架的说明和示例测试文件。) 使用该或`import`关键字将库`require`导入到对调用 Office JavaScript API 的函数进行测试的任何测试文件，如以下示例所示。

   ```javascript
   const OfficeAddinMock = require("office-addin-mock");
   ```

1. 导入包含要使用关键字或`import`关键字测试的外接程序函数的`require`模块。 下面是一个示例，假定测试文件位于包含外接程序代码文件的文件夹的子文件夹中。

   ```javascript
   const myOfficeAddinFeature = require("../my-office-add-in");
   ```

1. 创建一个数据对象，该对象具有测试函数所需的属性和子属性。 下面是模拟 Excel [Workbook.range.address](/javascript/api/excel/excel.range#excel-excel-range-address-member) 属性和 [Workbook.getSelectedRange 方法的对象的](/javascript/api/excel/excel.workbook#excel-excel-workbook-getselectedrange-member(1)) 示例。 这不是最终的模拟对象。 将其视为用于创建最终模拟对象的 `OfficeMockObject` 种子对象。

   ```javascript
   const mockData = {
     workbook: {
       range: {
         address: "C2:G3",
       },
       getSelectedRange: function () {
         return this.range;
       },
     },
   };
   ```

1. 将数据对象传递给 `OfficeMockObject` 构造函数。 请注意以下有关返回 `OfficeMockObject` 的对象的信息。

   - 它是 [OfficeExtension.ClientRequestContext](/javascript/api/office/officeextension.clientrequestcontext) 对象的简化模拟。
   - 模拟对象具有数据对象的所有成员，并且还具有模拟实现 `load` 和 `sync` 方法。
   - 模拟对象将模拟对象的关键错误行为 `ClientRequestContext` 。 例如，如果要测试的 Office API 尝试读取属性而不首先加载属性并调用 `sync`，则测试将失败，并出现类似于生产运行时引发的错误：“错误，属性未加载”。

   ```javascript
   const contextMock = new OfficeAddinMock.OfficeMockObject(mockData);
   ```

    > [!NOTE]
    > 该类型的完整参考文档 `OfficeMockObject` 位于 [Office-Addin-Mock](https://github.com/OfficeDev/Office-Addin-Scripts/tree/master/packages/office-addin-mock#reference) 中。

1. 在测试框架的语法中，添加函数的测试。 `OfficeMockObject`在本例`ClientRequestContext`中，使用该对象代替它模拟的对象。 下面的示例在 Jest 中继续。 此示例测试假定正在测试的外接程序函数被调用 `getSelectedRangeAddress`，它采用 `ClientRequestContext` 对象作为参数，并打算返回当前所选范围的地址。 本文 [稍后将](#mocking-a-clientrequestcontext-object)介绍完整示例。

   ```javascript
   test("getSelectedRangeAddress should return the address of the range", async function () {
     expect(await getSelectedRangeAddress(contextMock)).toBe("C2:G3");
   });
   ```

1. 根据测试框架和开发工具的文档运行测试。 通常，有一个 **package.json** 文件，其中包含执行测试框架的脚本。 例如，如果 Jest 是框架， **package.json** 将包含以下内容：

   ```json
   "scripts": {
     "test": "jest",
     -- other scripts omitted --  
   }
   ```

   若要运行测试，请在项目的根目录中的命令提示符中输入以下内容。

   ```command&nbsp;line
   npm test
   ```

## <a name="examples"></a>示例

本部分中的示例使用 Jest 及其默认设置。 这些设置支持 CommonJS 模块。 请参阅 [Jest 文档](https://jestjs.io/docs/getting-started) ，了解如何配置 Jest 和 node.js 以支持 ECMAScript 模块和支持 TypeScript。 若要运行上述任何示例，请执行以下步骤。

1. 为相应的 Office 主机应用程序 (（例如 Excel 或 Word) ）创建 Office 外接程序项目。 快速执行此操作的一种方法是使用 [Office 加载项的 Yeoman 生成器](../develop/yeoman-generator-overview.md)。
1. 在项目的根目录中， [安装 Jest](https://jestjs.io/docs/getting-started)。
1. [安装 office-addin-mock 工具](#install-the-tool)。
1. 创建与示例中的第一个文件完全一样的文件，并将其添加到包含项目的其他源文件（通常调用 `\src`）的文件夹中。
1. 创建源文件夹的子文件夹，并为其提供适当的名称，例如 `\tests`。
1. 创建与示例中的测试文件完全一样的文件，并将其添加到子文件夹。
1. `test`将脚本添加到 **package.json** 文件，然后运行测试，如 [基本用法](#basic-usage)中所述。

### <a name="mocking-the-office-common-apis"></a>模拟 Office 通用 API

此示例假定任何支持 [Office 公用 API](../develop/office-javascript-api-object-model.md) 的主机的 Office 加载项 (例如 Excel、PowerPoint 或 Word) 。 外接程序在名为 `my-common-api-add-in-feature.js`>a0/a0> 的文件中具有其功能之一。 下面显示了文件的内容。 该`addHelloWorldText`函数设置文本“Hello World！” 到文档中当前选择的任何内容;例如;Word 中的区域、Excel 中的单元格或 PowerPoint 中的文本框。

```javascript
const myCommonAPIAddinFeature = {

    addHelloWorldText: async () => {
        const options = { coercionType: Office.CoercionType.Text };
        await Office.context.document.setSelectedDataAsync("Hello World!", options);
    }
}
  
module.exports = myCommonAPIAddinFeature;
```

命名的测试文件 `my-common-api-add-in-feature.test.js` 位于子文件夹中，相对于加载项代码文件的位置。 下面显示了文件的内容。 请注意，顶层属性是 `context`[Office.Context](/javascript/api/office/office.context) 对象，因此被模拟的对象是此属性的父级：[Office](/javascript/api/office) 对象。 关于此代码，请注意以下几点：

- 构 `OfficeMockObject` 造函数 *不会* 将所有 Office 枚举类添加到模拟 `Office` 对象，因此 `CoercionType.Text` 加载项方法中引用的值必须在种子对象中显式添加。
- 由于节点进程中未加载 Office JavaScript 库， `Office` 因此必须声明和初始化加载项代码中引用的对象。

```javascript
const OfficeAddinMock = require("office-addin-mock");
const myCommonAPIAddinFeature = require("../my-common-api-add-in-feature");

// Create the seed mock object.
const mockData = {
    context: {
      document: {
        setSelectedDataAsync: function (data, options) {
          this.data = data;
          this.options = options;
        },
      },
    },
    // Mock the Office.CoercionType enum.
    CoercionType: {
      Text: {},
    },
};
  
// Create the final mock object from the seed object.
const officeMock = new OfficeAddinMock.OfficeMockObject(mockData);

// Create the Office object that is called in the addHelloWorldText function.
global.Office = officeMock;

/* Code that calls the test framework goes below this line. */

// Jest test
test("Text of selection in document should be set to 'Hello World'", async function () {
    await myCommonAPIAddinFeature.addHelloWorldText();
    expect(officeMock.context.document.data).toBe("Hello World!");
});
```

### <a name="mocking-the-outlook-apis"></a>模拟 Outlook API

尽管严格来说，Outlook API 是通用 API 模型的一部分，但它们具有围绕 [邮箱](/javascript/api/outlook/office.mailbox) 对象构建的特殊体系结构，因此我们为 Outlook 提供了一个独特的示例。 此示例假定 Outlook 在名为 `my-outlook-add-in-feature.js`>a0/a0> 的文件中具有其功能之一。 下面显示了文件的内容。 该`addHelloWorldText`函数设置文本“Hello World！” 到消息撰写窗口中当前选择的任何内容。

```javascript
const myOutlookAddinFeature = {

    addHelloWorldText: async () => {
        Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
      }
}

module.exports = myOutlookAddinFeature;
```

命名的测试文件 `my-outlook-add-in-feature.test.js` 位于子文件夹中，相对于加载项代码文件的位置。 下面显示了文件的内容。 请注意，顶层属性是 `context`[Office.Context](/javascript/api/office/office.context) 对象，因此被模拟的对象是此属性的父级：[Office](/javascript/api/office) 对象。 关于此代码，请注意以下几点：

- `host`模拟库在内部使用模拟对象上的属性来标识 Office 应用程序。 Outlook 是必需的。 它当前不适用于任何其他 Office 应用程序。
- 由于节点进程中未加载 Office JavaScript 库， `Office` 因此必须声明和初始化加载项代码中引用的对象。

```javascript
const OfficeAddinMock = require("office-addin-mock");
const myOutlookAddinFeature = require("../my-outlook-add-in-feature");

// Create the seed mock object.
const mockData = {
  // Identify the host to the mock library (required for Outlook).
  host: "outlook",
  context: {
    mailbox: {
      item: {
          setSelectedDataAsync: function (data) {
          this.data = data;
        },
      },
    },
  },
};
  
// Create the final mock object from the seed object.
const officeMock = new OfficeAddinMock.OfficeMockObject(mockData);

// Create the Office object that is called in the addHelloWorldText function.
global.Office = officeMock;

/* Code that calls the test framework goes below this line. */

// Jest test
test("Text of selection in message should be set to 'Hello World'", async function () {
    await myOutlookAddinFeature.addHelloWorldText();
    expect(officeMock.context.mailbox.item.data).toBe("Hello World!");
});
```

### <a name="mocking-the-office-application-specific-apis"></a>模拟特定于 Office 应用程序的 API

测试使用特定于应用程序的 API 的函数时，请务必模拟正确的对象类型。 有两个选项：

- 模拟 [OfficeExtension.ClientRequestObject](/javascript/api/office/officeextension.clientrequestcontext)。 当正在测试的函数满足以下两个条件时，请执行此操作：

  - 它不会调用 *主机*。`run` 函数，如 [Excel.run](/javascript/api/excel#Excel_run_batch_)。
  - 它不引用 *Host* 对象的任何其他直接属性或方法。

- 模拟 *Host* 对象，例如 [Excel](/javascript/api/excel) 或 [Word](/javascript/api/word)。 如果无法使用上述选项，请执行此操作。

下面的子节中提供了这两种类型的测试的示例。

#### <a name="mocking-a-clientrequestcontext-object"></a>模拟 ClientRequestContext 对象

此示例假定 Excel 加载项在名为 `my-excel-add-in-feature.js`> 的文件中具有其功能之一。 下面显示了文件的内容。 请注意， `getSelectedRangeAddress` 该方法是在传递给 `Excel.run`的回调中调用的帮助程序方法。

```javascript
const myExcelAddinFeature = {
    
    getSelectedRangeAddress: async (context) => {
        const range = context.workbook.getSelectedRange();      
        range.load("address");

        await context.sync();
      
        return range.address;
    }
}

module.exports = myExcelAddinFeature;
```

命名的测试文件 `my-excel-add-in-feature.test.js` 位于子文件夹中，相对于加载项代码文件的位置。 下面显示了文件的内容。 请注意，顶层属性是`workbook`，因此被模拟的对象是：对象`ClientRequestContext`的`Excel.Workbook`父级。

```javascript
const OfficeAddinMock = require("office-addin-mock");
const myExcelAddinFeature = require("../my-excel-add-in-feature");

// Create the seed mock object.
const mockData = {
    workbook: {
      range: {
        address: "C2:G3",
      },
      // Mock the Workbook.getSelectedRange method.
      getSelectedRange: function () {
        return this.range;
      },
    },
};

// Create the final mock object from the seed object.
const contextMock = new OfficeAddinMock.OfficeMockObject(mockData);

/* Code that calls the test framework goes below this line. */

// Jest test
test("getSelectedRangeAddress should return address of selected range", async function () {
  expect(await myOfficeAddinFeature.getSelectedRangeAddress(contextMock)).toBe("C2:G3");
});
```

#### <a name="mocking-a-host-object"></a>模拟主机对象

此示例假定 Word 加载项在名为 `my-word-add-in-feature.js`>a0/a0> 的文件中具有其功能之一。 下面显示了文件的内容。

```javascript
const myWordAddinFeature = {

  insertBlueParagraph: async () => {
    return Word.run(async (context) => {
      // Insert a paragraph at the end of the document.
      const paragraph = context.document.body.insertParagraph("Hello World", Word.InsertLocation.end);
  
      // Change the font color to blue.
      paragraph.font.color = "blue";
  
      await context.sync();
    });
  }
}

module.exports = myWordAddinFeature;
```

命名的测试文件 `my-word-add-in-feature.test.js` 位于子文件夹中，相对于加载项代码文件的位置。 下面显示了文件的内容。 请注意，顶层属性是 `context`一个 `ClientRequestContext` 对象，因此被模拟的对象是此属性的父属性：一个 `Word` 对象。 关于此代码，请注意以下几点：

- `OfficeMockObject`当构造函数创建最终模拟对象时，它将确保子`ClientRequestContext`对象具有`sync`和`load`方法。
- 构`OfficeMockObject`造函数 *不* 向模拟`Word`对象添加`run`函数，因此必须在种子对象中显式添加函数。
- 构 `OfficeMockObject` 造函数 *不会* 将所有 Word 枚举类添加到模拟 `Word` 对象，因此 `InsertLocation.end` 加载项方法中引用的值必须在种子对象中显式添加。
- 由于节点进程中未加载 Office JavaScript 库， `Word` 因此必须声明和初始化加载项代码中引用的对象。

```javascript
const OfficeAddinMock = require("office-addin-mock");
const myWordAddinFeature = require("../my-word-add-in-feature");

// Create the seed mock object.
const mockData = {
  context: {
    document: {
      body: {
        paragraph: {
          font: {},
        },
        // Mock the Body.insertParagraph method.
        insertParagraph: function (paragraphText, insertLocation) {
          this.paragraph.text = paragraphText;
          this.paragraph.insertLocation = insertLocation;
          return this.paragraph;
        },
      },
    },
  },
  // Mock the Word.InsertLocation enum.
  InsertLocation: {
    end: "end",
  },
  // Mock the Word.run function.
  run: async function(callback) {
    await callback(this.context);
  },
};

// Create the final mock object from the seed object.
const wordMock = new OfficeAddinMock.OfficeMockObject(mockData);

// Define and initialize the Word object that is called in the insertBlueParagraph function.
global.Word = wordMock;

/* Code that calls the test framework goes below this line. */

// Jest test set
describe("Insert blue paragraph at end tests", () => {

  test("color of paragraph", async function () {
    await myWordAddinFeature.insertBlueParagraph();  
    expect(wordMock.context.document.body.paragraph.font.color).toBe("blue");
  });

  test("text of paragraph", async function () {
    await myWordAddinFeature.insertBlueParagraph();
    expect(wordMock.context.document.body.paragraph.text).toBe("Hello World");
  });
})
```

> [!NOTE]
> 该类型的完整参考文档 `OfficeMockObject` 位于 [Office-Addin-Mock](https://github.com/OfficeDev/Office-Addin-Scripts/tree/master/packages/office-addin-mock#reference) 中。

## <a name="see-also"></a>另请参阅

- [Office-Addin-Mock npm 页面](https://www.npmjs.com/package/office-addin-mock) 安装点。 
- 开放源代码存储库是 [Office-Addin-Mock](https://github.com/OfficeDev/Office-Addin-Scripts/tree/master/packages/office-addin-mock)。
- [开玩笑](https://jestjs.io)
- [摩 卡](https://mochajs.org/)
- [茉莉花](https://jasmine.github.io/)
