---
title: 在 Office 加载项中使用 Office UI Fabric React
description: 了解如何在 Office 加载项中使用 Office UI Fabric React。
ms.date: 01/16/2020
localization_priority: Normal
ms.openlocfilehash: d0c46b30f361fcc9d2d7c08170b70c66f513bdae
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608469"
---
# <a name="use-office-ui-fabric-react-in-office-add-ins"></a><span data-ttu-id="6295d-103">在 Office 加载项中使用 Office UI Fabric React</span><span class="sxs-lookup"><span data-stu-id="6295d-103">Use Office UI Fabric React in Office Add-ins</span></span>

<span data-ttu-id="6295d-p101">Office UI Fabric 是用于生成 Office 和 Office 365 用户体验的 JavaScript 前端框架。如果使用 React 生成外接程序，请考虑使用 Fabric React 来创建用户体验。Fabric 提供了多个可在外接程序中使用的基于 React 的 UX 组件，如按钮或复选框。</span><span class="sxs-lookup"><span data-stu-id="6295d-p101">Office UI Fabric is a JavaScript front-end framework for building user experiences for Office and Office 365. If you build your add-in using React, consider using Fabric React to create your user experience. Fabric provides several React-based UX components, like buttons or checkboxes, that you can use in your add-in.</span></span>

<span data-ttu-id="6295d-107">本文介绍如何创建使用 React 构建的加载项, 并使用 Fabric React 组件。</span><span class="sxs-lookup"><span data-stu-id="6295d-107">This article describes how to create an add-in that's built with React and uses Fabric React components.</span></span> 

> [!NOTE]
> <span data-ttu-id="6295d-108">[Fabric Core](office-ui-fabric.md#use-fabric-core-icons-fonts-colors)是 Fabric React 附带的，这意味着在完成本文中的步骤后，你的加载项也有权访问 Fabric Core。</span><span class="sxs-lookup"><span data-stu-id="6295d-108">[Fabric Core](office-ui-fabric.md#use-fabric-core-icons-fonts-colors) is included with Fabric React, which means your add-in will also have access to Fabric Core after you've completed the steps in this article.</span></span>

## <a name="create-an-add-in-project"></a><span data-ttu-id="6295d-109">创建加载项项目</span><span class="sxs-lookup"><span data-stu-id="6295d-109">Create an add-in project</span></span>

<span data-ttu-id="6295d-110">将使用 Office 加载项的 Yeoman 生成器创建使用 React 的加载项项目。</span><span class="sxs-lookup"><span data-stu-id="6295d-110">You'll use the Yeoman generator for Office Add-ins to create an add-in project that uses React.</span></span>

### <a name="install-the-prerequisites"></a><span data-ttu-id="6295d-111">安装必备组件</span><span class="sxs-lookup"><span data-stu-id="6295d-111">Install the prerequisites</span></span>

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

### <a name="create-the-project"></a><span data-ttu-id="6295d-112">创建项目</span><span class="sxs-lookup"><span data-stu-id="6295d-112">Create the project</span></span>

[!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

- <span data-ttu-id="6295d-113">**选择项目类型:** `Office Add-in Task Pane project using React framework`</span><span class="sxs-lookup"><span data-stu-id="6295d-113">**Choose a project type:** `Office Add-in Task Pane project using React framework`</span></span>
- <span data-ttu-id="6295d-114">**选择脚本类型:** `TypeScript`</span><span class="sxs-lookup"><span data-stu-id="6295d-114">**Choose a script type:** `TypeScript`</span></span>
- <span data-ttu-id="6295d-115">**要如何命名加载项?**</span><span class="sxs-lookup"><span data-stu-id="6295d-115">**What do you want to name your add-in?**</span></span> `My Office Add-in`
- <span data-ttu-id="6295d-116">**要支持哪一个 Office 客户端应用程序?**</span><span class="sxs-lookup"><span data-stu-id="6295d-116">**Which Office client application would you like to support?**</span></span> `Word`

![Yeoman 生成器](../images/yo-office-word-react.png)

<span data-ttu-id="6295d-118">完成此向导后，生成器会创建项目，并安装支持的 Node 组件。</span><span class="sxs-lookup"><span data-stu-id="6295d-118">After you complete the wizard, the generator creates the project and installs supporting Node components.</span></span>

[!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

### <a name="try-it-out"></a><span data-ttu-id="6295d-119">试用</span><span class="sxs-lookup"><span data-stu-id="6295d-119">Try it out</span></span>

1. <span data-ttu-id="6295d-120">导航到项目的根文件夹。</span><span class="sxs-lookup"><span data-stu-id="6295d-120">Navigate to the root folder of the project.</span></span>

    ```command&nbsp;line
    cd "My Office Add-in"
    ```

2. <span data-ttu-id="6295d-121">完成以下步骤，以启动本地 Web 服务器并旁加载你的加载项。</span><span class="sxs-lookup"><span data-stu-id="6295d-121">Complete the following steps to start the local web server and sideload your add-in.</span></span>

    > [!NOTE]
    > <span data-ttu-id="6295d-122">Office 加载项应使用 HTTPS，而不是 HTTP（即便是在开发时也是如此）。</span><span class="sxs-lookup"><span data-stu-id="6295d-122">Office Add-ins should use HTTPS, not HTTP, even when you are developing.</span></span> <span data-ttu-id="6295d-123">如果系统在运行以下命令之一后提示你安装证书，请接受提示以安装 Yeoman 生成器提供的证书。</span><span class="sxs-lookup"><span data-stu-id="6295d-123">If you are prompted to install a certificate after you run one of the following commands, accept the prompt to install the certificate that the Yeoman generator provides.</span></span>

    > [!TIP]
    > <span data-ttu-id="6295d-124">如果在 Mac 上测试加载项，请先运行以下命令，然后再继续。</span><span class="sxs-lookup"><span data-stu-id="6295d-124">If you're testing your add-in on Mac, run the following command before proceeding.</span></span> <span data-ttu-id="6295d-125">运行此命令时，本地 Web 服务器将启动。</span><span class="sxs-lookup"><span data-stu-id="6295d-125">When you run this command, the local web server starts.</span></span>
    >
    > ```command&nbsp;line
    > npm run dev-server
    > ```

    - <span data-ttu-id="6295d-126">若要在 Word 中测试加载项，请在项目的根目录中运行以下命令。</span><span class="sxs-lookup"><span data-stu-id="6295d-126">To test your add-in in Word, run the following command in the root directory of your project.</span></span> <span data-ttu-id="6295d-127">这将启动本地的 Web 服务器（如果尚未运行的话），并使用加载的加载项打开 Word。</span><span class="sxs-lookup"><span data-stu-id="6295d-127">This starts the local web server (if it's not already running) and opens Word with your add-in loaded.</span></span>

        ```command&nbsp;line
        npm start
        ```

    - <span data-ttu-id="6295d-128">若要在浏览器版 Word 中测试加载项，请在项目的根目录中运行以下命令。</span><span class="sxs-lookup"><span data-stu-id="6295d-128">To test your add-in in Word on a browser, run the following command in the root directory of your project.</span></span> <span data-ttu-id="6295d-129">如果你运行此命令，本地 Web 服务器将启动（如果尚未运行的话）。</span><span class="sxs-lookup"><span data-stu-id="6295d-129">When you run this command, the local web server will start (if it's not already running).</span></span>

        ```command&nbsp;line
        npm run start:web
        ```

        <span data-ttu-id="6295d-130">若要使用加载项，请在 Word 网页版中打开新的文档，并按照[在 Office 网页版中旁加载 Office 加载项](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web)中的说明操作，以旁加载你的加载项。</span><span class="sxs-lookup"><span data-stu-id="6295d-130">To use your add-in, open a new document in Word on the web and then sideload your add-in by following the instructions in [Sideload Office Add-ins in Office on the web](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web).</span></span>

3. <span data-ttu-id="6295d-131">在 Word 中，依次选择“**开始**”选项卡和功能区中的“**显示任务窗格**”按钮，以打开加载项任务窗格。</span><span class="sxs-lookup"><span data-stu-id="6295d-131">In Word, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span> <span data-ttu-id="6295d-132">请注意任务窗格底部的“默认文本”和 "**运行**" 按钮。</span><span class="sxs-lookup"><span data-stu-id="6295d-132">Notice the default text and the **Run** button at the bottom of the task pane.</span></span> <span data-ttu-id="6295d-133">在本演练的其余部分中, 你将通过创建使用来自 Fabric React 的 UX 组件的 React 组件来重新定义此文本和按钮。</span><span class="sxs-lookup"><span data-stu-id="6295d-133">In the remainder of this walkthrough, you'll redefine this text and button by creating a React component that uses UX components from Fabric React.</span></span>

    ![Word 应用程序的屏幕截图，任务窗格中突出显示了 "显示任务窗格" 功能区按钮以及“运行……”按钮和前面的文本](../images/word-task-pane-yo-default.png)


## <a name="create-a-react-component-that-uses-fabric-react"></a><span data-ttu-id="6295d-135">创建使用 Fabric React 的 React 组件</span><span class="sxs-lookup"><span data-stu-id="6295d-135">Create a React component that uses Fabric React</span></span>

<span data-ttu-id="6295d-136">此时, 你已经创建了一个使用 React 构建的非常基本的任务窗格加载项。</span><span class="sxs-lookup"><span data-stu-id="6295d-136">At this point, you've created a very basic task pane add-in that's built using React.</span></span> <span data-ttu-id="6295d-137">接下来，完成以下步骤，在加载项项目中创建新的 React 组件 (`ButtonPrimaryExample`)。</span><span class="sxs-lookup"><span data-stu-id="6295d-137">Next, complete the following steps to create a new React component (`ButtonPrimaryExample`) within the add-in project.</span></span> <span data-ttu-id="6295d-138">该组件使用 Fabric React 的`Label`和`PrimaryButton`组件。</span><span class="sxs-lookup"><span data-stu-id="6295d-138">The component uses the `Label` and `PrimaryButton` components from Fabric React.</span></span>

1. <span data-ttu-id="6295d-139">打开 Yeoman 生成器创建的项目文件夹，并转到**src\taskpane\components**。</span><span class="sxs-lookup"><span data-stu-id="6295d-139">Open the project folder created by the Yeoman generator, and go to **src\taskpane\components**.</span></span>
2. <span data-ttu-id="6295d-140">在该文件夹中，创建一个名为“**Button.tsx**”的新文件。</span><span class="sxs-lookup"><span data-stu-id="6295d-140">In that folder, create a new file named **Button.tsx**.</span></span>
3. <span data-ttu-id="6295d-141">在 **Button.tsx** 中，输入以下代码以定义`ButtonPrimaryExample`组件。</span><span class="sxs-lookup"><span data-stu-id="6295d-141">In **Button.tsx**, add the following code to define the `ButtonPrimaryExample` component.</span></span>

```typescript
import * as React from 'react';
import { PrimaryButton, IButtonProps } from 'office-ui-fabric-react/lib/Button';
import { Label } from 'office-ui-fabric-react/lib/Label';

export class ButtonPrimaryExample extends React.Component<IButtonProps, {}> {
  public constructor(props) {
    super(props);
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

<span data-ttu-id="6295d-142">此代码将执行以下操作：</span><span class="sxs-lookup"><span data-stu-id="6295d-142">This code does the following:</span></span>

- <span data-ttu-id="6295d-143">引用使用 `import * as React from 'react';` 的 React 库。</span><span class="sxs-lookup"><span data-stu-id="6295d-143">References the React library using `import * as React from 'react';`.</span></span>
- <span data-ttu-id="6295d-144">参考用于创建 `ButtonPrimaryExample` 的 Fabric 组件（`PrimaryButton`、`IButtonProps`、`Label`）。</span><span class="sxs-lookup"><span data-stu-id="6295d-144">References the Fabric components (`PrimaryButton`, `IButtonProps`, `Label`) that are used to create `ButtonPrimaryExample`.</span></span>
- <span data-ttu-id="6295d-145">声明新的 `ButtonPrimaryExample` 组件使用 `export class ButtonPrimaryExample extends React.Component`。</span><span class="sxs-lookup"><span data-stu-id="6295d-145">Declares the new `ButtonPrimaryExample` component using `export class ButtonPrimaryExample extends React.Component`.</span></span>
- <span data-ttu-id="6295d-146">声明 `insertText` 将处理按钮 `onClick` 事件的函数。</span><span class="sxs-lookup"><span data-stu-id="6295d-146">Declares the `insertText` function that will handle the button's `onClick` event.</span></span>
- <span data-ttu-id="6295d-147">定义 `render` 函数中 React 组件的 UI。</span><span class="sxs-lookup"><span data-stu-id="6295d-147">Defines the UI of the React component in the `render` function.</span></span> <span data-ttu-id="6295d-148">HTML 标记使用 Fabric Reac 中的组件 `Label` 和 `PrimaryButton`，并指定当 `onClick` 事件触发时，`insertText` 函数将运行。</span><span class="sxs-lookup"><span data-stu-id="6295d-148">The HTML markup uses the `Label` and `PrimaryButton` components from Fabric React and specifies that when the `onClick` event fires, the `insertText` function will run.</span></span>

## <a name="add-the-react-component-to-your-add-in"></a><span data-ttu-id="6295d-149">将 React 组件添加到加载项</span><span class="sxs-lookup"><span data-stu-id="6295d-149">Add the React component to your add-in</span></span>

<span data-ttu-id="6295d-150">通过打开 **src\components\App.tsx** 并完成下列步骤，将组件 `ButtonPrimaryExample` 添加到加载项：</span><span class="sxs-lookup"><span data-stu-id="6295d-150">Add the `ButtonPrimaryExample` component to your add-in by opening **src\components\App.tsx** and completing the following steps:</span></span>

1. <span data-ttu-id="6295d-151">添加下面导入语句，以`ButtonPrimaryExample`从**Button.tsx**中引用。</span><span class="sxs-lookup"><span data-stu-id="6295d-151">Add the following import statement to reference `ButtonPrimaryExample` from **Button.tsx**.</span></span>

    ```typescript
    import {ButtonPrimaryExample} from './Button';
    ```

2. <span data-ttu-id="6295d-152">删除以下两个导入语句。</span><span class="sxs-lookup"><span data-stu-id="6295d-152">Remove the following two import statements.</span></span>

    ```typescript
    import { Button, ButtonType } from 'office-ui-fabric-react';
    ...
    import Progress from './Progress';
    ```

3. <span data-ttu-id="6295d-153">将默认 `render()` 函数替换为以下使用 `ButtonPrimaryExample` 的代码。</span><span class="sxs-lookup"><span data-stu-id="6295d-153">Replace the default `render()` function with the following code that uses `ButtonPrimaryExample`.</span></span>

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

  4. <span data-ttu-id="6295d-154">将所做的更改保存到**App.tsx**。</span><span class="sxs-lookup"><span data-stu-id="6295d-154">Save the changes you've made to **App.tsx**.</span></span>

## <a name="see-the-result"></a><span data-ttu-id="6295d-155">查看结果</span><span class="sxs-lookup"><span data-stu-id="6295d-155">See the result</span></span>

<span data-ttu-id="6295d-156">在 Word 中, 当你保存对**App.tsx**的更改时，加载项任务窗格会自动更新。</span><span class="sxs-lookup"><span data-stu-id="6295d-156">In Word, the add-in task pane automatically updates when you save changes to **App.tsx**.</span></span> <span data-ttu-id="6295d-157">任务窗格底部的默认文本和按钮现在显示由该`ButtonPrimaryExample`组件定义的 UI。</span><span class="sxs-lookup"><span data-stu-id="6295d-157">The default text and button at the bottom of the task pane now shows the UI that's defined by the `ButtonPrimaryExample` component.</span></span> <span data-ttu-id="6295d-158">选择**插入文本……** 按钮将文本插入到文档中。</span><span class="sxs-lookup"><span data-stu-id="6295d-158">Choose the **Insert text...** button to insert text into the document.</span></span>

![Word 应用程序的屏幕截图，突出显示 "插入文本……" 按钮和前面的文本](../images/word-task-pane-with-react-component.png)

<span data-ttu-id="6295d-160">恭喜，您已使用 React 和 Office UI Fabric React 成功创建了任务窗格加载项！</span><span class="sxs-lookup"><span data-stu-id="6295d-160">Congratulations, you've successfully created a task pane add-in using React and Office UI Fabric React!</span></span> 

## <a name="see-also"></a><span data-ttu-id="6295d-161">另请参阅</span><span class="sxs-lookup"><span data-stu-id="6295d-161">See also</span></span>

- [<span data-ttu-id="6295d-162">Office 加载项中的 Office UI Fabric</span><span class="sxs-lookup"><span data-stu-id="6295d-162">Office UI Fabric in Office Add-ins</span></span>](office-ui-fabric.md)
- [<span data-ttu-id="6295d-163">Office UI Fabric React</span><span class="sxs-lookup"><span data-stu-id="6295d-163">Office UI Fabric React</span></span>](https://developer.microsoft.com/fabric)
- [<span data-ttu-id="6295d-164">适用于 Office 加载项的 UX 设计模式</span><span class="sxs-lookup"><span data-stu-id="6295d-164">UX design patterns for Office Add-ins</span></span>](ux-design-pattern-templates.md)
- [<span data-ttu-id="6295d-165">Fabric React 代码示例入门</span><span class="sxs-lookup"><span data-stu-id="6295d-165">Getting started with Fabric React code sample</span></span>](https://github.com/OfficeDev/Word-Add-in-GettingStartedFabricReact)
