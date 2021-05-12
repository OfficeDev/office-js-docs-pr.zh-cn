---
title: Fluent UI React外接程序Office中的用户界面
description: 了解如何在外接程序React Fluent UI Office。
ms.date: 05/12/2021
localization_priority: Normal
ms.openlocfilehash: cb7f04c21a52a2e4a3f271abc56aa325dd2b02fd
ms.sourcegitcommit: 30f6c620380075e3459cac748ca0c656427b384d
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/12/2021
ms.locfileid: "52330139"
---
# <a name="use-fluent-ui-react-in-office-add-ins"></a><span data-ttu-id="91421-103">在加载项React Fluent UI Office</span><span class="sxs-lookup"><span data-stu-id="91421-103">Use Fluent UI React in Office Add-ins</span></span>

<span data-ttu-id="91421-104">Fluent UI React 是官方开源 JavaScript 前端框架，旨在构建无缝适用于各种 Microsoft 产品（包括 Office）的体验。</span><span class="sxs-lookup"><span data-stu-id="91421-104">Fluent UI React is the official open-source JavaScript front-end framework designed to build experiences that fit seamlessly into a broad range of Microsoft products, including Office.</span></span> <span data-ttu-id="91421-105">它提供可靠的、最新的、可访问的React组件，这些组件使用 CSS-in-JS 进行高度自定义。</span><span class="sxs-lookup"><span data-stu-id="91421-105">It provides robust, up-to-date, accessible React-based components which are highly customizable using CSS-in-JS.</span></span>

> [!NOTE]
> <span data-ttu-id="91421-106">本文介绍 Fluent UI React在加载项上下文中Office使用。但它还用于各种应用Microsoft 365扩展。</span><span class="sxs-lookup"><span data-stu-id="91421-106">This article describes the use of Fluent UI React in the context of Office Add-ins. But it is also used in a wide range of Microsoft 365 apps and extensions.</span></span> <span data-ttu-id="91421-107">有关详细信息，请参阅[Fluent UI React](https://developer.microsoft.com/fluentui#/get-started/web#fluent-ui-react)和开源存储库 Fluent UI [Web。](https://github.com/microsoft/fluentui)</span><span class="sxs-lookup"><span data-stu-id="91421-107">For more information, see [Fluent UI React](https://developer.microsoft.com/fluentui#/get-started/web#fluent-ui-react) and the open source repo [Fluent UI Web](https://github.com/microsoft/fluentui).</span></span>

<span data-ttu-id="91421-108">本文介绍如何创建使用应用程序构建的外接程序，React Fluent UI React组件。</span><span class="sxs-lookup"><span data-stu-id="91421-108">This article describes how to create an add-in that's built with React and uses Fluent UI React components.</span></span>

## <a name="create-an-add-in-project"></a><span data-ttu-id="91421-109">创建加载项项目</span><span class="sxs-lookup"><span data-stu-id="91421-109">Create an add-in project</span></span>

<span data-ttu-id="91421-110">将使用 Office 加载项的 Yeoman 生成器创建使用 React 的加载项项目。</span><span class="sxs-lookup"><span data-stu-id="91421-110">You'll use the Yeoman generator for Office Add-ins to create an add-in project that uses React.</span></span>

### <a name="install-the-prerequisites"></a><span data-ttu-id="91421-111">安装必备组件</span><span class="sxs-lookup"><span data-stu-id="91421-111">Install the prerequisites</span></span>

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

### <a name="create-the-project"></a><span data-ttu-id="91421-112">创建项目</span><span class="sxs-lookup"><span data-stu-id="91421-112">Create the project</span></span>

[!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

- <span data-ttu-id="91421-113">**选择项目类型:** `Office Add-in Task Pane project using React framework`</span><span class="sxs-lookup"><span data-stu-id="91421-113">**Choose a project type:** `Office Add-in Task Pane project using React framework`</span></span>
- <span data-ttu-id="91421-114">**选择脚本类型:** `TypeScript`</span><span class="sxs-lookup"><span data-stu-id="91421-114">**Choose a script type:** `TypeScript`</span></span>
- <span data-ttu-id="91421-115">**要如何命名加载项?**</span><span class="sxs-lookup"><span data-stu-id="91421-115">**What do you want to name your add-in?**</span></span> `My Office Add-in`
- <span data-ttu-id="91421-116">**要支持哪一个 Office 客户端应用程序?**</span><span class="sxs-lookup"><span data-stu-id="91421-116">**Which Office client application would you like to support?**</span></span> `Word`

![显示命令行界面中 Yeoman 生成器的提示和回答的屏幕截图](../images/yo-office-word-react.png)

<span data-ttu-id="91421-118">完成此向导后，生成器会创建项目，并安装支持的 Node 组件。</span><span class="sxs-lookup"><span data-stu-id="91421-118">After you complete the wizard, the generator creates the project and installs supporting Node components.</span></span>

[!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

### <a name="try-it-out"></a><span data-ttu-id="91421-119">试用</span><span class="sxs-lookup"><span data-stu-id="91421-119">Try it out</span></span>

1. <span data-ttu-id="91421-120">导航到项目的根文件夹。</span><span class="sxs-lookup"><span data-stu-id="91421-120">Navigate to the root folder of the project.</span></span>

    ```command&nbsp;line
    cd "My Office Add-in"
    ```

2. <span data-ttu-id="91421-121">完成以下步骤，以启动本地 Web 服务器并旁加载你的加载项。</span><span class="sxs-lookup"><span data-stu-id="91421-121">Complete the following steps to start the local web server and sideload your add-in.</span></span>

    > [!NOTE]
    > <span data-ttu-id="91421-122">Office 加载项应使用 HTTPS，而不是 HTTP（即便是在开发时也是如此）。</span><span class="sxs-lookup"><span data-stu-id="91421-122">Office Add-ins should use HTTPS, not HTTP, even when you are developing.</span></span> <span data-ttu-id="91421-123">如果系统在运行以下命令之一后提示你安装证书，请接受提示以安装 Yeoman 生成器提供的证书。</span><span class="sxs-lookup"><span data-stu-id="91421-123">If you are prompted to install a certificate after you run one of the following commands, accept the prompt to install the certificate that the Yeoman generator provides.</span></span> <span data-ttu-id="91421-124">你可能还必须以管理员身份运行命令提示符或终端才能进行更改。</span><span class="sxs-lookup"><span data-stu-id="91421-124">You may also have to run your command prompt or terminal as an administrator for the changes to be made.</span></span>

    > [!TIP]
    > <span data-ttu-id="91421-125">如果在 Mac 上测试加载项，请先运行以下命令，然后再继续。</span><span class="sxs-lookup"><span data-stu-id="91421-125">If you're testing your add-in on Mac, run the following command before proceeding.</span></span> <span data-ttu-id="91421-126">运行此命令时，本地 Web 服务器将启动。</span><span class="sxs-lookup"><span data-stu-id="91421-126">When you run this command, the local web server starts.</span></span>
    >
    > ```command&nbsp;line
    > npm run dev-server
    > ```

    - <span data-ttu-id="91421-127">若要在 Word 中测试加载项，请在项目的根目录中运行以下命令。</span><span class="sxs-lookup"><span data-stu-id="91421-127">To test your add-in in Word, run the following command in the root directory of your project.</span></span> <span data-ttu-id="91421-128">这将启动本地的 Web 服务器（如果尚未运行的话），并使用加载的加载项打开 Word。</span><span class="sxs-lookup"><span data-stu-id="91421-128">This starts the local web server (if it's not already running) and opens Word with your add-in loaded.</span></span>

        ```command&nbsp;line
        npm start
        ```

    - <span data-ttu-id="91421-129">若要在浏览器版 Word 中测试加载项，请在项目的根目录中运行以下命令。</span><span class="sxs-lookup"><span data-stu-id="91421-129">To test your add-in in Word on a browser, run the following command in the root directory of your project.</span></span> <span data-ttu-id="91421-130">如果你运行此命令，本地 Web 服务器将启动（如果尚未运行的话）。</span><span class="sxs-lookup"><span data-stu-id="91421-130">When you run this command, the local web server will start (if it's not already running).</span></span>

        ```command&nbsp;line
        npm run start:web
        ```

        <span data-ttu-id="91421-131">若要使用加载项，请在 Word 网页版中打开新的文档，并按照[在 Office 网页版中旁加载 Office 加载项](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web)中的说明操作，以旁加载你的加载项。</span><span class="sxs-lookup"><span data-stu-id="91421-131">To use your add-in, open a new document in Word on the web and then sideload your add-in by following the instructions in [Sideload Office Add-ins in Office on the web](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web).</span></span>

3. <span data-ttu-id="91421-132">若要打开加载项任务窗格，在"开始 **"选项卡上** ，选择" **显示任务窗格"** 按钮。</span><span class="sxs-lookup"><span data-stu-id="91421-132">To open the add-in task pane, on the **Home** tab, choose the **Show Taskpane** button.</span></span> <span data-ttu-id="91421-133">请注意任务窗格底部的“默认文本”和 "**运行**" 按钮。</span><span class="sxs-lookup"><span data-stu-id="91421-133">Notice the default text and the **Run** button at the bottom of the task pane.</span></span> <span data-ttu-id="91421-134">在此演练的其余部分中，你将通过创建使用 Fluent UI React 中的 UX 组件的 React 组件来重新定义此文本和React。</span><span class="sxs-lookup"><span data-stu-id="91421-134">In the remainder of this walkthrough, you'll redefine this text and button by creating a React component that uses UX components from Fluent UI React.</span></span>

    ![Screenshot showing the Word application with the Show Taskpane ribbon button highlighted and the Run button and immediately preceding text highlighted in the task pane](../images/word-task-pane-yo-default.png)

## <a name="create-a-react-component-that-uses-fluent-ui-react"></a><span data-ttu-id="91421-136">创建使用 Fluent UI React的 React</span><span class="sxs-lookup"><span data-stu-id="91421-136">Create a React component that uses Fluent UI React</span></span>

<span data-ttu-id="91421-137">此时, 你已经创建了一个使用 React 构建的非常基本的任务窗格加载项。</span><span class="sxs-lookup"><span data-stu-id="91421-137">At this point, you've created a very basic task pane add-in that's built using React.</span></span> <span data-ttu-id="91421-138">接下来，完成以下步骤，在加载项项目中创建新的 React 组件 (`ButtonPrimaryExample`)。</span><span class="sxs-lookup"><span data-stu-id="91421-138">Next, complete the following steps to create a new React component (`ButtonPrimaryExample`) within the add-in project.</span></span> <span data-ttu-id="91421-139">该组件使用 `Label` Fluent UI 元素中的 和 `PrimaryButton` React。</span><span class="sxs-lookup"><span data-stu-id="91421-139">The component uses the `Label` and `PrimaryButton` components from Fluent UI React.</span></span>

1. <span data-ttu-id="91421-140">打开 Yeoman 生成器创建的项目文件夹，并转到 **src\taskpane\components**。</span><span class="sxs-lookup"><span data-stu-id="91421-140">Open the project folder created by the Yeoman generator, and go to **src\taskpane\components**.</span></span>
2. <span data-ttu-id="91421-141">在该文件夹中，创建一个名为“**Button.tsx**”的新文件。</span><span class="sxs-lookup"><span data-stu-id="91421-141">In that folder, create a new file named **Button.tsx**.</span></span>
3. <span data-ttu-id="91421-142">在 **Button.tsx** 中，输入以下代码以定义`ButtonPrimaryExample`组件。</span><span class="sxs-lookup"><span data-stu-id="91421-142">In **Button.tsx**, add the following code to define the `ButtonPrimaryExample` component.</span></span>

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
      body.insertParagraph('Hello Fluent UI React!', Word.InsertLocation.end);
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

<span data-ttu-id="91421-143">此代码将执行以下操作：</span><span class="sxs-lookup"><span data-stu-id="91421-143">This code does the following:</span></span>

- <span data-ttu-id="91421-144">引用使用 `import * as React from 'react';` 的 React 库。</span><span class="sxs-lookup"><span data-stu-id="91421-144">References the React library using `import * as React from 'react';`.</span></span>
- <span data-ttu-id="91421-145">引用 Fluent UI React、 (`PrimaryButton` 、) 用于创建 `IButtonProps` `Label` 的组件 `ButtonPrimaryExample` 。</span><span class="sxs-lookup"><span data-stu-id="91421-145">References the Fluent UI React components (`PrimaryButton`, `IButtonProps`, `Label`) that are used to create `ButtonPrimaryExample`.</span></span>
- <span data-ttu-id="91421-146">声明新的 `ButtonPrimaryExample` 组件使用 `export class ButtonPrimaryExample extends React.Component`。</span><span class="sxs-lookup"><span data-stu-id="91421-146">Declares the new `ButtonPrimaryExample` component using `export class ButtonPrimaryExample extends React.Component`.</span></span>
- <span data-ttu-id="91421-147">声明 `insertText` 将处理按钮 `onClick` 事件的函数。</span><span class="sxs-lookup"><span data-stu-id="91421-147">Declares the `insertText` function that will handle the button's `onClick` event.</span></span>
- <span data-ttu-id="91421-148">定义 `render` 函数中 React 组件的 UI。</span><span class="sxs-lookup"><span data-stu-id="91421-148">Defines the UI of the React component in the `render` function.</span></span> <span data-ttu-id="91421-149">HTML 标记使用 Fluent UI 元素中的 和 React并指定当事件触发时 `Label` `PrimaryButton` `onClick` `insertText` ，函数将运行。</span><span class="sxs-lookup"><span data-stu-id="91421-149">The HTML markup uses the `Label` and `PrimaryButton` components from Fluent UI React and specifies that when the `onClick` event fires, the `insertText` function will run.</span></span>

## <a name="add-the-react-component-to-your-add-in"></a><span data-ttu-id="91421-150">将 React 组件添加到加载项</span><span class="sxs-lookup"><span data-stu-id="91421-150">Add the React component to your add-in</span></span>

<span data-ttu-id="91421-151">通过打开 **src\components\App.tsx** 并完成下列步骤，将组件 `ButtonPrimaryExample` 添加到加载项：</span><span class="sxs-lookup"><span data-stu-id="91421-151">Add the `ButtonPrimaryExample` component to your add-in by opening **src\components\App.tsx** and completing the following steps:</span></span>

1. <span data-ttu-id="91421-152">添加下面导入语句，以`ButtonPrimaryExample`从 **Button.tsx** 中引用。</span><span class="sxs-lookup"><span data-stu-id="91421-152">Add the following import statement to reference `ButtonPrimaryExample` from **Button.tsx**.</span></span>

    ```typescript
    import {ButtonPrimaryExample} from './Button';
    ```

2. <span data-ttu-id="91421-153">删除以下两个导入语句。</span><span class="sxs-lookup"><span data-stu-id="91421-153">Remove the following two import statements.</span></span>

    ```typescript
    import { Button, ButtonType } from 'office-ui-fabric-react';
    ...
    import Progress from './Progress';
    ```

3. <span data-ttu-id="91421-154">将默认 `render()` 函数替换为以下使用 `ButtonPrimaryExample` 的代码。</span><span class="sxs-lookup"><span data-stu-id="91421-154">Replace the default `render()` function with the following code that uses `ButtonPrimaryExample`.</span></span>

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

4. <span data-ttu-id="91421-155">将所做的更改保存到 **App.tsx**。</span><span class="sxs-lookup"><span data-stu-id="91421-155">Save the changes you've made to **App.tsx**.</span></span>

## <a name="see-the-result"></a><span data-ttu-id="91421-156">查看结果</span><span class="sxs-lookup"><span data-stu-id="91421-156">See the result</span></span>

<span data-ttu-id="91421-157">在 Word 中, 当你保存对 **App.tsx** 的更改时，加载项任务窗格会自动更新。</span><span class="sxs-lookup"><span data-stu-id="91421-157">In Word, the add-in task pane automatically updates when you save changes to **App.tsx**.</span></span> <span data-ttu-id="91421-158">任务窗格底部的默认文本和按钮现在显示由该`ButtonPrimaryExample`组件定义的 UI。</span><span class="sxs-lookup"><span data-stu-id="91421-158">The default text and button at the bottom of the task pane now shows the UI that's defined by the `ButtonPrimaryExample` component.</span></span> <span data-ttu-id="91421-159">选择 **插入文本……** 按钮将文本插入到文档中。</span><span class="sxs-lookup"><span data-stu-id="91421-159">Choose the **Insert text...** button to insert text into the document.</span></span>

![显示具有"插入文本..."的 Word 应用程序的屏幕截图按钮和紧接突出显示的文本](../images/word-task-pane-with-react-component.png)

<span data-ttu-id="91421-161">恭喜！你已成功使用 React 和 Fluent UI 加载项创建React！</span><span class="sxs-lookup"><span data-stu-id="91421-161">Congratulations, you've successfully created a task pane add-in using React and Fluent UI React!</span></span>

## <a name="see-also"></a><span data-ttu-id="91421-162">另请参阅</span><span class="sxs-lookup"><span data-stu-id="91421-162">See also</span></span>

- [<span data-ttu-id="91421-163">Word 外接程序 GettingStartedFabricReact</span><span class="sxs-lookup"><span data-stu-id="91421-163">Word Add-in GettingStartedFabricReact</span></span>](https://github.com/OfficeDev/Word-Add-in-GettingStartedFabricReact)
- [<span data-ttu-id="91421-164">Office外接程序中的 Fabric Core</span><span class="sxs-lookup"><span data-stu-id="91421-164">Fabric Core in Office Add-ins</span></span>](fabric-core.md)
- [<span data-ttu-id="91421-165">适用于 Office 加载项的 UX 设计模式</span><span class="sxs-lookup"><span data-stu-id="91421-165">UX design patterns for Office Add-ins</span></span>](ux-design-pattern-templates.md)
