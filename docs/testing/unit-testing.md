---
title: 外接程序中的Office测试
description: 了解如何对调用 JavaScript API 的 Office代码
ms.date: 11/14/2021
ms.localizationpriority: medium
ms.openlocfilehash: 3daf47a6221d5c9dbc0ad9fe1c357264d0a2f622
ms.sourcegitcommit: 67b70f5328e4b9c9e9df098ec98f29a02f363464
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/19/2021
ms.locfileid: "61124832"
---
# <a name="unit-testing-in-office-add-ins"></a>外接程序中的Office测试

单元测试无需网络连接或服务连接（包括与加载项应用程序的连接）即可检查Office功能。 单元测试服务器端代码和不调用[Office JavaScript](../develop/understanding-the-javascript-api-for-office.md)API的客户端代码在 Office 外接程序中与在任何 Web 应用程序中相同，因此不需要特殊文档。 但是调用 JavaScript API 的Office代码很难测试。 为了解决这些问题，我们创建了一个库来简化单元测试中的 mock Office 对象的创建[：Office-Addin-Mock](https://www.npmjs.com/package/office-addin-mock)。 该库通过以下方式使测试变得更简单：

- Office JavaScript API 必须在 Office 应用程序 (Excel、Word 等 ) 上下文的 Web 视图控件中初始化，因此无法在开发计算机上运行单元测试的过程中加载它们。 可以将 Office-Addin-Mock 库导入测试文件，从而可以在运行测试的 node.js 进程中模拟 Office JavaScript API。
- 特定于 [应用程序的 API](../develop/understanding-the-javascript-api-for-office.md#api-models) 具有 [加载](../develop/application-specific-api-model.md#load) 和 [同步](../develop/application-specific-api-model.md#sync) 方法，这些方法必须相对于其他函数和彼此以特定顺序调用。 此外，必须使用特定参数调用方法，具体取决于要测试的函数中稍后的代码将读取 Office 对象的属性 `load` 。  但是单元测试框架本身是无状态的，因此它们无法记录是否已调用或传递 `load` `sync` 了哪些参数 `load` 。 使用 Addin-Mock Office创建的 mock 对象具有可跟踪这些内容的内部状态。 这使 mock 对象能够模拟实际对象Office行为。 例如，如果正在测试的函数尝试读取未首先传递到 的属性，则测试将返回一个类似于Office `load` 的错误。

该库不依赖于 javaScript OFFICE，它可用于任何 JavaScript 单元测试框架，例如：

- [Jest](https://jestjs.io)
- [Mocha](https://mochajs.org/)
- [Storybook](https://storybook.js.org/docs/react/workflows/unit-testing)
- [Jasmine](https://jasmine.github.io/)

本文中的示例使用 Jest 框架。 在[Office-Addin-Mock 主页上提供了使用 Mocha 框架的示例](https://github.com/OfficeDev/Office-Addin-Scripts/tree/master/packages/office-addin-mock#examples)。

## <a name="prerequisites"></a>先决条件

本文假定你熟悉单元测试和模拟的基本概念，包括如何创建和运行测试文件，并且你具有单元测试框架的一些经验。

> [!TIP]
> 如果你使用 Visual Studio，我们建议你阅读 Visual Studio 中的单元测试[JavaScript 和 TypeScript](/visualstudio/javascript/unit-testing-javascript-with-visual-studio.md)一文，获取有关 Visual Studio 中 JavaScript 单元测试的一些基本信息，然后返回到本文。

## <a name="install-the-tool"></a>安装工具

若要安装库，请打开命令提示符，导航到加载项项目的根目录，然后输入以下命令。

```command&nbsp;line
npm install office-addin-mock --save-dev
```

## <a name="basic-usage"></a>基本用法

1. 项目将具有一个或多个测试文件。  (请参阅下面的示例 (#examples) 中的测试文件示例的说明。) 将库（带 或 关键字）导入到具有调用 `require` Office JavaScript API 的函数测试的任何测试文件，如以下示例所示。 `import`

   ```javascript
   const OfficeAddinMock = require("office-addin-mock");
   ```

1. 导入包含您想要使用 或 关键字测试的外接程序函数的 `require` `import` 模块。 下面的示例假定测试文件位于包含加载项代码文件的文件夹的子文件夹中。

   ```javascript
   const myOfficeAddinFeature = require("../my-office-add-in");
   ```

1. 创建一个数据对象，该对象具有测试函数需要模拟的属性和子属性。 下面是模拟[Workbook.range.address](/javascript/api/excel/excel.range#address) Excel [Workbook.getSelectedRange](/javascript/api/excel/excel.workbook#getSelectedRange__)方法的对象示例。 这不是最终的 mock 对象。 将该对象视为用于创建最终 mock 对象的 `OfficeMockObject` 种子对象。

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

1. 将数据对象传递给 `OfficeMockObject` 构造函数。 对于返回的对象，请注意 `OfficeMockObject` 以下几点。

   - 它是 [OfficeExtension.ClientRequestContext](/javascript/api/office/officeextension.clientrequestcontext) 对象的简化模型。
   - mock 对象具有 data 对象的所有成员，并且具有 和 方法的 `load` mock `sync` 实现。
   - mock 对象将模拟对象的关键错误 `ClientRequestContext` 行为。 例如，如果你正在测试的 Office API 尝试读取属性，而未首先加载属性并调用 ，则测试将失败，并出现类似于在生产运行时中抛出的错误：" `sync` 错误，未加载属性"。

   ```javascript
   const contextMock = new OfficeAddinMock.OfficeMockObject(mockData);
   ```

    > [!NOTE]
    > 有关类型的完整参考 `OfficeMockObject` 文档，请参阅[Office-Addin-Mock](https://github.com/OfficeDev/Office-Addin-Scripts/tree/master/packages/office-addin-mock#reference)。

1. 在测试框架的语法中，添加 函数的测试。 使用 `OfficeMockObject` 对象来表示它所模拟的对象，本例中为 `ClientRequestContext` 对象。 下面继续 Jest 中的示例。 此示例测试假定正在测试的加载项函数被调用 ，它采用对象作为参数，并打算返回当前选定区域 `getSelectedRangeAddress` `ClientRequestContext` 的地址。 本文稍后将 [介绍完整示例](#mocking-a-clientrequestcontext-object)。

   ```javascript
   test("getSelectedRangeAddress should return the address of the range", async function () {
     expect(await getSelectedRangeAddress(contextMock)).toBe("C2:G3");
   });
   ```

1. 根据测试框架和开发工具的文档运行测试。 通常，存在一个 **package.json** 文件，该文件具有用于执行测试框架的脚本。 例如，如果 Jest 是框架 **，package.json** 将包含以下内容：

   ```json
   "scripts": {
     "test": "jest",
     -- other scripts omitted --  
   }
   ```

   若要运行测试，请在项目根目录的命令提示符中输入以下内容。

   ```command&nbsp;line
   npm test
   ```

## <a name="examples"></a>示例

本节中的示例使用 Jest 及其默认设置。 这些设置支持 CommonJS 模块。 请参阅 [Jest 文档，](https://jestjs.io/docs/getting-started) 了解如何配置 Jest 和 node.js以支持 ECMAScript 模块和支持 TypeScript。 若要运行其中任何示例，请执行以下步骤。

> [!NOTE]
> 

1. 为Office应用程序创建一个Office加载项项目 (，例如Excel Word) 。 快速完成此操作的一个方法就是使用[Yo Office 工具](https://github.com/OfficeDev/generator-office)。
1. 在项目的根中，安装 [Jest](https://jestjs.io/docs/getting-started)。
1. [安装 office-addin-mock 工具](#install-the-tool)。
1. 创建一个与示例中第一个文件完全相同的文件，并将其添加到包含项目的其他源文件（通常称为 ）的文件夹 `\src` 。
1. 创建源文件文件夹的子文件夹，然后为它指定一个合适的名称，例如 `\tests` 。
1. 创建一个与示例中的测试文件完全相同的文件，并将其添加到子文件夹。
1. 将 `test` 脚本添加到 **package.json** 文件，然后运行测试，如 [基本](#basic-usage)用法 中所述。

### <a name="mocking-the-office-common-apis"></a>模拟Office API

本示例为Office通用 API 的任何主机Office [Office](../develop/office-javascript-api-object-model.md)一个加载项 (例如 Excel、PowerPoint 或 Word) 。 加载项在名为 的文件中有一项功能 `my-common-api-add-in-feature.js` 。 下面显示了文件的内容。 函数 `addHelloWorldText` 设置文本"Hello World！" 为文档中当前选择的任何内容;例如;Word 中的一个范围、Excel单元格或 word 中的PowerPoint。

```javascript
const myCommonAPIAddinFeature = {

    addHelloWorldText: async () => {
        const options = { coercionType: Office.CoercionType.Text };
        await Office.context.document.setSelectedDataAsync("Hello World!", options);
    }
}
  
module.exports = myCommonAPIAddinFeature;
```

名为 的测试 `my-common-api-add-in-feature.test.js` 文件位于子文件夹（相对于加载项代码文件的位置） 中。 下面显示了文件的内容。 请注意，顶级属性是 `context` [，Office。Context](/javascript/api/office/office.context)对象，因此要模拟的对象是此属性的父对象：一个[Office对象。](/javascript/api/office) 关于此代码，请注意以下几点：

- 构造函数不会将所有 Office枚举类添加到 mock 对象，因此必须在 seed 对象中显式添加外接程序方法中引用 `OfficeMockObject`  `Office` `CoercionType.Text` 的值。
- 由于Office JavaScript 库未加载到节点进程中，因此加载项代码中引用的对象必须声明 `Office` 和初始化。

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

### <a name="mocking-the-outlook-apis"></a>模拟Outlook API

尽管严格来说，Outlook API 是通用 API 模型的一部分，但是它们具有围绕[Mailbox](/javascript/api/office/office.mailbox)对象构建的特殊体系结构，因此我们为 Outlook 提供了一个Outlook。 此示例假定一Outlook一个在名为 的文件中具有其功能之一的组 `my-outlook-add-in-feature.js` 。 下面显示了文件的内容。 函数 `addHelloWorldText` 设置文本"Hello World！" 为当前在邮件撰写窗口中选择的任何内容。

```javascript
const myOutlookAddinFeature = {

    addHelloWorldText: async () => {
        Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
      }
}

module.exports = myOutlookAddinFeature;
```

名为 的测试 `my-outlook-add-in-feature.test.js` 文件位于子文件夹（相对于加载项代码文件的位置） 中。 下面显示了文件的内容。 请注意，顶级属性是 `context` [，Office。Context](/javascript/api/office/office.context)对象，因此要模拟的对象是此属性的父对象：一个[Office对象。](/javascript/api/office) 关于此代码，请注意以下几点：

- 由于Office JavaScript 库未加载到节点进程中，因此加载项代码中引用的对象必须声明 `Office` 和初始化。

```javascript
const OfficeAddinMock = require("office-addin-mock");
const myOutlookAddinFeature = require("../my-outlook-add-in-feature");

// Create the seed mock object.
const mockData = {
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

### <a name="mocking-the-office-application-specific-apis"></a>模拟Office应用程序特定的 API

在测试使用特定于应用程序的 API 的函数时，请确保模拟的对象类型正确。 有两个选项：

- Mock a [OfficeExtension.ClientRequestObject](/javascript/api/office/officeextension.clientrequestcontext). 当要测试的函数满足以下两个条件时，可执行下列操作：

  - 它不会 *调用主机*。`run` 方法，例如[Excel.run](/javascript/api/excel#Excel_run_batch_)。
  - 它不引用 Host 对象的其他任何直接属性 *或* 方法。

- 模拟 *Host* 对象，如 [Excel](/javascript/api/excel)或 [Word](/javascript/api/word)。 如果上述选项不可行，则执行上述步骤。

下面各小节中提供了这两种类型的测试的示例。

#### <a name="mocking-a-clientrequestcontext-object"></a>模拟 ClientRequestContext 对象

此示例假定一Excel一个外接程序，该加载项在名为 的文件中具有其一个功能 `my-excel-add-in-feature.js` 。 下面显示了文件的内容。 请注意， 是在传递给 的回调内 `getSelectedRangeAddress` 调用的帮助程序方法 `Excel.run` 。

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

名为 的测试 `my-excel-add-in-feature.test.js` 文件位于子文件夹（相对于加载项代码文件的位置） 中。 下面显示了文件的内容。 请注意，顶级属性是 ，因此要模拟的对象是 `workbook` `Excel.Workbook` ： 对象的 `ClientRequestContext` 父对象。

```javascript
const OfficeAddinMock = require("office-addin-mock");
const myExcelAddinFeature = require("../my-excel-add-in-feature");

// Create the seed mock object.
const mockData = {
    workbook: {
      range: {
        address: "C2:G3",
      },
      // Mock the Workbook.getSelectRange method.
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

本示例假定 Word 加载项在名为 的文件中具有其功能之一 `my-word-add-in-feature.js` 。 下面显示了文件的内容。

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

名为 的测试 `my-word-add-in-feature.test.js` 文件位于子文件夹（相对于加载项代码文件的位置） 中。 下面显示了文件的内容。 请注意，顶级属性是 ，一个对象，因此要模拟的对象是此属性的父对象 `context` `ClientRequestContext` ： `Word` 一个对象。 关于此代码，请注意以下几点：

- 当 `OfficeMockObject` 构造函数创建最终 mock 对象时，它将确保子对象具有 `ClientRequestContext` 和 `sync` `load` 方法。
- 构造函数不会向 mock 对象添加方法，因此必须在 seed 对象中 `OfficeMockObject`  `run` `Word` 显式添加该方法。
- 构造函数 `OfficeMockObject` 不会 *将* 所有 Word 枚举类添加到 mock 对象，因此必须在 seed 对象中显式添加在外接程序方法中引用 `Word` `InsertLocation.end` 的值。
- 由于Office JavaScript 库未加载到节点进程中，因此加载项代码中引用的对象必须声明 `Word` 和初始化。

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
  // Mock the Word.run method.
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

## <a name="adding-mock-objects-properties-and-methods-dynamically-when-testing"></a>测试时动态添加 mock 对象、属性和方法

在某些情况下，高效测试要求在运行时创建或修改 mock 对象;即，当测试正在运行时。 示例如下：

- 被测试的函数在被调用第二次时的行为会有所不同。 你需要首先使用一个 mock 对象测试函数，然后更改此 mock 对象，然后使用已更改的 mock 对象再次测试该函数。
- 你需要针对多个相似但不完全相同的 mock 对象测试函数。 例如，你需要使用具有 color 属性的 mock 对象测试函数，然后再次使用具有文本属性但与原始 mock 对象相同的 mock 对象测试该函数。

`OfficeMockObject`有三种方法可帮助这些方案。

- `OfficeMockObject.setMock` 向对象添加属性和 `OfficeMockObject` 值。 以下示例添加 `address` 属性。

    ```javascript
    rangeMock.setMock("address", "G6:K9");
    ```

- `OfficeMockObject.addMockFunction` 向对象添加 mock `OfficeMockObject` 函数，如以下示例所示。

    ```javascript
    workbookMock.addMockFunction("getSelectedRange", function () { 
      const range = {
        address: "B2:G5",
      };
      return range;
    });
    ```

    > [!NOTE]
    > 函数参数是可选的。 如果不存在，将创建一个空函数。

- `OfficeMockObject.addMock` 将新 `OfficeMockObject` 对象作为属性添加到现有对象并命名。 它将具有全部具有的最小 `OfficeMockObject` 成员，例如 和 `load` `sync` 。 可以使用 和 方法添加 `setMock` `addMockFunction` 其他成员。 下面是一个将 mock 对象作为属性 `Excel.WorkbookProtection` 添加到 `protection` mock 工作簿的示例。 然后，它将 `protected` 属性添加到新的 mock 对象。

    ```javascript
    workbookMock.addMock("protection");
    workbookMock.protection.setMock("protected", true);
    ```

> [!NOTE]
> 有关类型的完整参考 `OfficeMockObject` 文档，请参阅[Office-Addin-Mock](https://github.com/OfficeDev/Office-Addin-Scripts/tree/master/packages/office-addin-mock#reference)。

## <a name="see-also"></a>另请参阅

- [Office-Addin-Mock npm 页面](https://www.npmjs.com/package/office-addin-mock)安装点。 
- 开放源存储库是[Office-Addin-Mock 。](https://github.com/OfficeDev/Office-Addin-Scripts/tree/master/packages/office-addin-mock)
- [Jest](https://jestjs.io)
- [Mocha](https://mochajs.org/)
- [Storybook](https://storybook.js.org/docs/react/workflows/unit-testing)
- [Jasmine](https://jasmine.github.io/)
 