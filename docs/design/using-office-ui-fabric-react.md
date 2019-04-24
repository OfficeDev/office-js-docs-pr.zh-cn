---
title: 在 Office 加载项中使用 Office UI Fabric React
description: ''
ms.date: 02/28/2019
localization_priority: Priority
ms.openlocfilehash: 11bb9daf99d85f1c4551363e9f04056870631378
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/24/2019
ms.locfileid: "32449027"
---
# <a name="use-office-ui-fabric-react-in-office-add-ins"></a><span data-ttu-id="99838-102">在 Office 加载项中使用 Office UI Fabric React</span><span class="sxs-lookup"><span data-stu-id="99838-102">Use Office UI Fabric React in Office Add-ins</span></span>

<span data-ttu-id="99838-p101">Office UI Fabric 是用于生成 Office 和 Office 365 用户体验的 JavaScript 前端框架。如果使用 React 生成外接程序，请考虑使用 Fabric React 来创建用户体验。Fabric 提供了多个可在外接程序中使用的基于 React 的 UX 组件，如按钮或复选框。</span><span class="sxs-lookup"><span data-stu-id="99838-p101">Office UI Fabric is a JavaScript front-end framework for building user experiences for Office and Office 365. If you build your add-in using React, consider using Fabric React to create your user experience. Fabric provides several React-based UX components, like buttons or checkboxes, that you can use in your add-in.</span></span>

<span data-ttu-id="99838-106">若要开始在加载项中使用 Fabric React 组件，请执行以下步骤。</span><span class="sxs-lookup"><span data-stu-id="99838-106">To get started using Fabric React's components in your add-in, perform the following steps.</span></span>

> [!NOTE]
> <span data-ttu-id="99838-107">如果按照本文中的步骤操作，也可以在加载项中使用 Fabric Core。</span><span class="sxs-lookup"><span data-stu-id="99838-107">If you follow the steps in this article, Fabric Core is also available in your add-in.</span></span>

## <a name="step-1---create-your-project-with-the-yeoman-generator-for-office"></a><span data-ttu-id="99838-108">第 1 步 - 使用 Office 的 Yeoman 生成器创建项目</span><span class="sxs-lookup"><span data-stu-id="99838-108">Step 1 - Create your project with the Yeoman generator for Office</span></span>

<span data-ttu-id="99838-109">若要创建使用 Fabric React 的外接程序，我们建议使用 Office 的 Yeoman 生成器。</span><span class="sxs-lookup"><span data-stu-id="99838-109">To create an add-in that uses Fabric React, we recommend that you use the Yeoman generator for Office.</span></span> <span data-ttu-id="99838-110">Office 的 Yeoman 生成器提供开发 Office 外接程序所需的项目基架和版本管理。</span><span class="sxs-lookup"><span data-stu-id="99838-110">The Yeoman generator for Office provides the project scaffolding and build management needed to develop an Office Add-in.</span></span>

<span data-ttu-id="99838-111">若要创建项目，请使用 **Windows PowerShell**（而不是命令提示符）执行以下步骤：</span><span class="sxs-lookup"><span data-stu-id="99838-111">To create your project, perform the following steps using **Windows PowerShell** (not the command prompt):</span></span>

1. <span data-ttu-id="99838-112">安装必备组件。</span><span class="sxs-lookup"><span data-stu-id="99838-112">Install the prerequisites.</span></span>
2. <span data-ttu-id="99838-113">运行 `yo office`，为外接程序创建项目文件。</span><span class="sxs-lookup"><span data-stu-id="99838-113">Run `yo office` to create the project files for your add-in.</span></span>
3. <span data-ttu-id="99838-114">当系统提示你选择一个 Office 客户端应用程序时，请选择 **Word**。</span><span class="sxs-lookup"><span data-stu-id="99838-114">When prompted to select an Office client application, choose **Word**.</span></span>
4. <span data-ttu-id="99838-p103">确保位于包含项目文件的目录中，再运行 `npm start`。此时，显示旋转图标的浏览器窗口自动打开。</span><span class="sxs-lookup"><span data-stu-id="99838-p103">Ensure you are in the directory with the project files, and then run `npm start`. A browser window showing a spinner opens automatically.</span></span>
5. <span data-ttu-id="99838-117">[旁加载清单](../testing/test-debug-office-add-ins.md)，以查看加载项的完整 UI。</span><span class="sxs-lookup"><span data-stu-id="99838-117">[Sideload your manifest](../testing/test-debug-office-add-ins.md) to view the full UI of the add-in.</span></span>

## <a name="step-2---add-a-fabric-react-component"></a><span data-ttu-id="99838-118">第 2 步 - 添加 Fabric React 组件</span><span class="sxs-lookup"><span data-stu-id="99838-118">Step 2 - Add a Fabric React component</span></span>

<span data-ttu-id="99838-p104">接下来，将 Fabric React 组件添加到外接程序。创建称为 `ButtonPrimaryExample` 的新的 React 组件，其中包含来自 Fabric React 的标签和 PrimaryButton。创建 `ButtonPrimaryExample`：</span><span class="sxs-lookup"><span data-stu-id="99838-p104">Next, add Fabric React components to your add-in. Create a new React component, called `ButtonPrimaryExample`, that consists of a Label and PrimaryButton from Fabric React. To create `ButtonPrimaryExample`:</span></span>

1. <span data-ttu-id="99838-122">打开 Yeoman 生成器创建的项目文件夹，并转到 **src\components**。</span><span class="sxs-lookup"><span data-stu-id="99838-122">Open the project folder created by the Yeoman generator, and go to **src\components**.</span></span>
2. <span data-ttu-id="99838-123">创建 **button.tsx**。</span><span class="sxs-lookup"><span data-stu-id="99838-123">Create **button.tsx**.</span></span>
3. <span data-ttu-id="99838-124">在 **button.tsx** 中，输入以下代码以创建 `ButtonPrimaryExample` 组件。</span><span class="sxs-lookup"><span data-stu-id="99838-124">In **button.tsx**, enter the following code to create the `ButtonPrimaryExample` component.</span></span>

```typescript
import * as React from 'react';
import { PrimaryButton, IButtonProps } from 'office-ui-fabric-react/lib/Button';
import { Label } from 'office-ui-fabric-react/lib/Label';

export class ButtonPrimaryExample extends React.Component<IButtonProps, {}> {
  public constructor() {
    super();
  }

   insertText = async () => {
        // In the click event, write text to the document.
        await Word.run(async (context) => {
            let body = context.document.body;
            body.insertParagraph('Hello Office UI Fabric React!', Word.InsertLocation.end);
            await context.sync();
        });
    }

  public render() {
    let { disabled } = this.props;
    return (
      <div className='ms-BasicButtonsExample'>
        <Label>Click the button to insert text.</Label>
        <PrimaryButton
          data-automation-id='test'
          disabled={ disabled }
          text='Insert text...'
          onClick={ this.insertText } />
      </div>
    );
  }
}
```

<span data-ttu-id="99838-125">此代码将执行以下操作：</span><span class="sxs-lookup"><span data-stu-id="99838-125">This code does the following:</span></span>

- <span data-ttu-id="99838-126">引用使用 `import * as React from 'react';` 的 React 库。</span><span class="sxs-lookup"><span data-stu-id="99838-126">References the React library using `import * as React from 'react';`.</span></span>
- <span data-ttu-id="99838-127">引用用于创建 `ButtonPrimaryExample` 的 Fabric 组件（PrimaryButton、IButtonProps、标签）。</span><span class="sxs-lookup"><span data-stu-id="99838-127">References the Fabric components (PrimaryButton, IButtonProps, Label) that are used to create `ButtonPrimaryExample`.</span></span>
- <span data-ttu-id="99838-128">使用 `export class ButtonPrimaryExample extends React.Component`，声明并公开新的 `ButtonPrimaryExample` 组件。</span><span class="sxs-lookup"><span data-stu-id="99838-128">Declares and make public the new `ButtonPrimaryExample` component using `export class ButtonPrimaryExample extends React.Component`.</span></span>
- <span data-ttu-id="99838-129">将 `insertText` 函数声明为处理 `onClick` 事件。</span><span class="sxs-lookup"><span data-stu-id="99838-129">Declares the `insertText` function to handle the `onClick` event.</span></span>
- <span data-ttu-id="99838-p105">在 `render` 函数中定义 React 组件的 UI。呈现器定义组件结构。在 `render` 中，使用 `this.insertText` 连接 `onClick` 事件。</span><span class="sxs-lookup"><span data-stu-id="99838-p105">Defines the UI of the React component in the `render` function. Render defines the structure of the component. Within `render`, you wire up the `onClick` event using `this.insertText`.</span></span>

## <a name="step-3---add-the-react-component-to-your-add-in"></a><span data-ttu-id="99838-133">第 3 步 - 将 React 组件添加到加载项</span><span class="sxs-lookup"><span data-stu-id="99838-133">Step 3 - Add the React component to your add-in</span></span>

<span data-ttu-id="99838-134">通过打开 **src\components\app.tsx** 并执行下列操作将 `ButtonPrimaryExample` 添加到外接程序：</span><span class="sxs-lookup"><span data-stu-id="99838-134">Add `ButtonPrimaryExample` to your add-in by opening **src\components\app.tsx** and doing the following:</span></span>

- <span data-ttu-id="99838-135">添加以下导入语句以引用来自步骤 2 中创建的 **button.tsx** 的引用 `ButtonPrimaryExample`（不需要文件扩展名）。</span><span class="sxs-lookup"><span data-stu-id="99838-135">Add the following import statement to reference `ButtonPrimaryExample` from **button.tsx** created in step 2 (no file extension is needed).</span></span>

  ```typescript
  import {ButtonPrimaryExample} from './button';
  ```

- <span data-ttu-id="99838-136">将默认 `render()` 函数替换为以下使用 `<ButtonPrimaryExample />` 的代码。</span><span class="sxs-lookup"><span data-stu-id="99838-136">Replace the default `render()` function with the following code that uses `<ButtonPrimaryExample />`.</span></span>

  ```typescript
  render() {
      return (
          <div className="ms-welcome">
          <Header logo="assets/logo-filled.png" title={this.props.title} message="Welcome" />
          <HeroList message="Discover what this add-in can do for you today!" items={this.state.listItems} >
              <ButtonPrimaryExample />
          </HeroList>
          </div>
      );
  }
  ```

<span data-ttu-id="99838-p106">保存所做的更改。所有打开的浏览器实例（包括外接程序）将自动更新和显示 `ButtonPrimaryExample` React 组件。请注意，默认文本和按钮将替换为 `ButtonPrimaryExample` 中定义的文本和主按钮。</span><span class="sxs-lookup"><span data-stu-id="99838-p106">Save your changes. All open browser instances, including the add-in, update automatically and show the `ButtonPrimaryExample` React component. Notice that the default text and button is replaced with the text and primary button defined in `ButtonPrimaryExample`.</span></span>



## <a name="see-also"></a><span data-ttu-id="99838-140">另请参阅</span><span class="sxs-lookup"><span data-stu-id="99838-140">See also</span></span>

- [<span data-ttu-id="99838-141">Office UI Fabric React</span><span class="sxs-lookup"><span data-stu-id="99838-141">Office UI Fabric React</span></span>](https://developer.microsoft.com/fabric)
- [<span data-ttu-id="99838-142">适用于 Office 加载项的 UX 设计模式</span><span class="sxs-lookup"><span data-stu-id="99838-142">UX design patterns for Office Add-ins</span></span>](../design/ux-design-pattern-templates.md)
- [<span data-ttu-id="99838-143">Fabric React 代码示例入门</span><span class="sxs-lookup"><span data-stu-id="99838-143">Getting started with Fabric React code sample</span></span>](https://github.com/OfficeDev/Word-Add-in-GettingStartedFabricReact)
- [<span data-ttu-id="99838-144">Office 外接程序 Fabric UI 示例（使用 Fabric 1.0）</span><span class="sxs-lookup"><span data-stu-id="99838-144">Office Add-in Fabric UI sample (uses Fabric 1.0)</span></span>](https://github.com/OfficeDev/Office-Add-in-Fabric-UI-Sample)
- [<span data-ttu-id="99838-145">Office 的 yeoman 生成器</span><span class="sxs-lookup"><span data-stu-id="99838-145">Yeoman generator for Office</span></span>](https://github.com/OfficeDev/generator-office)
