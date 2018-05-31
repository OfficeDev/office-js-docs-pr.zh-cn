---
title: 在 Office 加载项中使用 Office UI Fabric React
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: 8ae8bac8c8043b51188d765dd7170922dcc1c84e
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/23/2018
ms.locfileid: "19437596"
---
# <a name="use-office-ui-fabric-react-in-office-add-ins"></a><span data-ttu-id="afe9d-102">在 Office 加载项中使用 Office UI Fabric React</span><span class="sxs-lookup"><span data-stu-id="afe9d-102">Use Office UI Fabric React in Office Add-ins</span></span>

<span data-ttu-id="afe9d-p101">Office UI Fabric 是用于生成 Office 和 Office 365 用户体验的 JavaScript 前端框架。如果使用 React 生成外接程序，请考虑使用 Fabric React 来创建用户体验。Fabric 提供了多个可在外接程序中使用的基于 React 的 UX 组件，如按钮或复选框。</span><span class="sxs-lookup"><span data-stu-id="afe9d-p101">Office UI Fabric is a JavaScript front-end framework for building user experiences for Office and Office 365. If you build your add-in using React, consider using Fabric React to create your user experience. Fabric provides several React-based UX components, like buttons or checkboxes, that you can use in your add-in.</span></span>

<span data-ttu-id="afe9d-106">若要开始在加载项中使用 Fabric React 组件，请执行以下步骤。</span><span class="sxs-lookup"><span data-stu-id="afe9d-106">To get started using Fabric React's components in your add-in, perform the following steps.</span></span>

> [!NOTE]
> <span data-ttu-id="afe9d-107">如果按照本文中的步骤操作，也可以在加载项中使用 Fabric Core。</span><span class="sxs-lookup"><span data-stu-id="afe9d-107">If you follow the steps in this article, Fabric Core is also available in your add-in.</span></span>

## <a name="step-1---create-your-project-with-the-yeoman-generator-for-office"></a><span data-ttu-id="afe9d-108">第 1 步 - 使用适用于 Office 的 Yeoman 生成器创建项目</span><span class="sxs-lookup"><span data-stu-id="afe9d-108">Step 1 - Create your project with the Yeoman generator for Office</span></span>

<span data-ttu-id="afe9d-p102">若要创建使用 Fabric React 的外接程序，我们建议使用 Office 的 Yeoman 生成器。Office 的 Yeoman 生成器提供开发 Office 外接程序所需的项目基架和版本管理。</span><span class="sxs-lookup"><span data-stu-id="afe9d-p102">To create an add-in that uses Fabric React, we recommend that you use the Yeoman generator for Office. The Yeoman generator for Office provides the project scaffolding and build management needed to develop an Office add-in.</span></span>

<span data-ttu-id="afe9d-111">若要创建项目，请使用 **Windows PowerShell**（而不是命令提示符）执行以下步骤：</span><span class="sxs-lookup"><span data-stu-id="afe9d-111">To create your project, perform the following steps using **Windows PowerShell** (not the command prompt):</span></span>

1. <span data-ttu-id="afe9d-112">安装必备组件。</span><span class="sxs-lookup"><span data-stu-id="afe9d-112">Install the prerequisites.</span></span>
2. <span data-ttu-id="afe9d-113">运行 `yo office`，为外接程序创建项目文件。</span><span class="sxs-lookup"><span data-stu-id="afe9d-113">Run `yo office` to create the project files for your add-in.</span></span>
3. <span data-ttu-id="afe9d-114">当系统提示你选择一个 Office 客户端应用程序时，请选择 **Word**。</span><span class="sxs-lookup"><span data-stu-id="afe9d-114">When prompted to select an Office client application, choose **Word**.</span></span>
4. <span data-ttu-id="afe9d-p103">确保位于包含项目文件的目录中，再运行 `npm start`。此时，显示旋转图标的浏览器窗口自动打开。</span><span class="sxs-lookup"><span data-stu-id="afe9d-p103">Ensure you are in the directory with the project files, and then run `npm start`. A browser window showing a spinner opens automatically.</span></span>
5. <span data-ttu-id="afe9d-117">[旁加载清单](..\testing\test-debug-office-add-ins.md)，以查看加载项的完整 UI。</span><span class="sxs-lookup"><span data-stu-id="afe9d-117">[Sideload your manifest](..\testing\test-debug-office-add-ins.md) to view the full UI of the add-in.</span></span>

## <a name="step-2---add-a-fabric-react-component"></a><span data-ttu-id="afe9d-118">第 2 步 - 添加 Fabric React 组件</span><span class="sxs-lookup"><span data-stu-id="afe9d-118">Step 2 - Add a Fabric React component</span></span>

<span data-ttu-id="afe9d-p104">接下来，将 Fabric React 组件添加到外接程序。创建称为 `ButtonPrimaryExample` 的新的 React 组件，其中包含来自 Fabric React 的标签和 PrimaryButton。创建 `ButtonPrimaryExample`：</span><span class="sxs-lookup"><span data-stu-id="afe9d-p104">Next, add Fabric React components to your add-in. Create a new React component, called `ButtonPrimaryExample`, that consists of a Label and PrimaryButton from Fabric React. To create `ButtonPrimaryExample`:</span></span>

1. <span data-ttu-id="afe9d-122">打开 Yeoman 生成器创建的项目文件夹，并转到 **src\components**。</span><span class="sxs-lookup"><span data-stu-id="afe9d-122">Open the project folder created by the Yeoman generator, and go to **src\components**.</span></span>
2. <span data-ttu-id="afe9d-123">创建 **button.tsx**。</span><span class="sxs-lookup"><span data-stu-id="afe9d-123">Create **button.tsx**.</span></span>
3. <span data-ttu-id="afe9d-124">在 **button.tsx** 中，输入以下代码以创建 `ButtonPrimaryExample` 组件。</span><span class="sxs-lookup"><span data-stu-id="afe9d-124">In **button.tsx**, enter the following code to create the `ButtonPrimaryExample` component.</span></span>

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

<span data-ttu-id="afe9d-125">此代码将执行以下操作：</span><span class="sxs-lookup"><span data-stu-id="afe9d-125">This code does the following:</span></span>

- <span data-ttu-id="afe9d-126">引用使用 `import * as React from 'react';` 的 React 库。</span><span class="sxs-lookup"><span data-stu-id="afe9d-126">References the React library using `import * as React from 'react';`.</span></span>
- <span data-ttu-id="afe9d-127">引用用于创建 `ButtonPrimaryExample` 的 Fabric 组件（PrimaryButton、IButtonProps、标签）。</span><span class="sxs-lookup"><span data-stu-id="afe9d-127">References the Fabric components (PrimaryButton, IButtonProps, Label) that are used to create `ButtonPrimaryExample`.</span></span>
- <span data-ttu-id="afe9d-128">使用 `export class ButtonPrimaryExample extends React.Component`，声明并公开新的 `ButtonPrimaryExample` 组件。</span><span class="sxs-lookup"><span data-stu-id="afe9d-128">Declares and make public the new `ButtonPrimaryExample` component using `export class ButtonPrimaryExample extends React.Component`.</span></span>
- <span data-ttu-id="afe9d-129">将 `insertText` 函数声明为处理 `onClick` 事件。</span><span class="sxs-lookup"><span data-stu-id="afe9d-129">Declares the `insertText` function to handle the `onClick` event.</span></span>
- <span data-ttu-id="afe9d-p105">在 `render` 函数中定义 React 组件的 UI。呈现器定义组件结构。在 `render` 中，使用 `this.insertText` 连接 `onClick` 事件。</span><span class="sxs-lookup"><span data-stu-id="afe9d-p105">Defines the UI of the React component in the `render` function. Render defines the structure of the component. Within `render`, you wire up the `onClick` event using `this.insertText`.</span></span>

## <a name="step-3---add-the-react-component-to-your-add-in"></a><span data-ttu-id="afe9d-133">第 3 步 - 将 React 组件添加到加载项</span><span class="sxs-lookup"><span data-stu-id="afe9d-133">Step 3 - Add the React component to your add-in</span></span>

<span data-ttu-id="afe9d-134">通过打开 **src\components\app.tsx** 并执行下列操作将 `ButtonPrimaryExample` 添加到外接程序：</span><span class="sxs-lookup"><span data-stu-id="afe9d-134">Add `ButtonPrimaryExample` to your add-in by opening **src\components\app.tsx** and doing the following:</span></span>

- <span data-ttu-id="afe9d-135">添加以下导入语句以引用来自步骤 2 中创建的 **button.tsx** 的引用 `ButtonPrimaryExample`（不需要文件扩展名）。</span><span class="sxs-lookup"><span data-stu-id="afe9d-135">Add the following import statement to reference `ButtonPrimaryExample` from **button.tsx** created in step 2 (no file extension is needed).</span></span>

  ```typescript
  import {ButtonPrimaryExample} from './button';
  ```

- <span data-ttu-id="afe9d-136">将默认 `render()` 函数替换为以下使用 `<ButtonPrimaryExample />` 的代码。</span><span class="sxs-lookup"><span data-stu-id="afe9d-136">Replace the default `render()` function with the following code that uses `<ButtonPrimaryExample />`.</span></span>

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

<span data-ttu-id="afe9d-p106">保存所做的更改。所有打开的浏览器实例（包括外接程序）将自动更新和显示 `ButtonPrimaryExample` React 组件。请注意，默认文本和按钮将替换为 `ButtonPrimaryExample` 中定义的文本和主按钮。</span><span class="sxs-lookup"><span data-stu-id="afe9d-p106">Save your changes. All open browser instances, including the add-in, update automatically and show the `ButtonPrimaryExample` React component. Notice that the default text and button is replaced with the text and primary button defined in `ButtonPrimaryExample`.</span></span>

## <a name="recommended-components"></a><span data-ttu-id="afe9d-140">建议使用的组件</span><span class="sxs-lookup"><span data-stu-id="afe9d-140">Recommended components</span></span>

<span data-ttu-id="afe9d-141">下面列出了建议用于加载项的 Fabric React 用户体验组件：</span><span class="sxs-lookup"><span data-stu-id="afe9d-141">The following is a list of the Fabric React UX components that we recommend for use in an add-in:</span></span>

- [<span data-ttu-id="afe9d-142">痕迹导航</span><span class="sxs-lookup"><span data-stu-id="afe9d-142">Breadcrumb</span></span>](breadcrumb.md)
- [<span data-ttu-id="afe9d-143">按钮</span><span class="sxs-lookup"><span data-stu-id="afe9d-143">Button</span></span>](button.md)
- [<span data-ttu-id="afe9d-144">复选框</span><span class="sxs-lookup"><span data-stu-id="afe9d-144">Checkbox</span></span>](checkbox.md)
- [<span data-ttu-id="afe9d-145">选择组</span><span class="sxs-lookup"><span data-stu-id="afe9d-145">ChoiceGroup</span></span>](choicegroup.md)
- [<span data-ttu-id="afe9d-146">下拉列表</span><span class="sxs-lookup"><span data-stu-id="afe9d-146">Dropdown</span></span>](dropdown.md)
- [<span data-ttu-id="afe9d-147">标签</span><span class="sxs-lookup"><span data-stu-id="afe9d-147">Label</span></span>](label.md)
- [<span data-ttu-id="afe9d-148">列表</span><span class="sxs-lookup"><span data-stu-id="afe9d-148">List</span></span>](list.md)
- [<span data-ttu-id="afe9d-149">透视</span><span class="sxs-lookup"><span data-stu-id="afe9d-149">Pivot</span></span>](pivot.md)
- [<span data-ttu-id="afe9d-150">文字框</span><span class="sxs-lookup"><span data-stu-id="afe9d-150">TextField</span></span>](textfield.md)
- [<span data-ttu-id="afe9d-151">切换</span><span class="sxs-lookup"><span data-stu-id="afe9d-151">Toggle</span></span>](toggle.md)

> [!NOTE]
> <span data-ttu-id="afe9d-152">今后，我们将陆续添加其他组件。</span><span class="sxs-lookup"><span data-stu-id="afe9d-152">We will add additional components over time.</span></span>

## <a name="see-also"></a><span data-ttu-id="afe9d-153">另请参阅</span><span class="sxs-lookup"><span data-stu-id="afe9d-153">See also</span></span>

- [<span data-ttu-id="afe9d-154">Office UI Fabric React</span><span class="sxs-lookup"><span data-stu-id="afe9d-154">Office UI Fabric React</span></span>](https://dev.office.com/fabric#/)
- [<span data-ttu-id="afe9d-155">Fabric React 代码示例入门</span><span class="sxs-lookup"><span data-stu-id="afe9d-155">Getting started with Fabric React code sample</span></span>](https://github.com/OfficeDev/Word-Add-in-GettingStartedFabricReact)
- [<span data-ttu-id="afe9d-156">用户体验设计模式（使用 Fabric 2.6.1）</span><span class="sxs-lookup"><span data-stu-id="afe9d-156">UX design patterns (uses Fabric 2.6.1)</span></span>](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)
- [<span data-ttu-id="afe9d-157">Office 外接程序 Fabric UI 示例（使用 Fabric 1.0）</span><span class="sxs-lookup"><span data-stu-id="afe9d-157">Office Add-in Fabric UI sample (uses Fabric 1.0)</span></span>](https://github.com/OfficeDev/Office-Add-in-Fabric-UI-Sample)
- [<span data-ttu-id="afe9d-158">Office 的 Yeoman 生成器</span><span class="sxs-lookup"><span data-stu-id="afe9d-158">Yeoman generator for Office</span></span>](https://github.com/OfficeDev/generator-office)
