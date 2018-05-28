# <a name="build-an-excel-add-in-using-jquery"></a><span data-ttu-id="c6818-101">?? jQuery ?? Excel ???</span><span class="sxs-lookup"><span data-stu-id="c6818-101">Build an Excel add-in using jQuery</span></span>

<span data-ttu-id="c6818-102">??????????? jQuery ? Excel JavaScript API ?? Excel ????</span><span class="sxs-lookup"><span data-stu-id="c6818-102">In this article, you'll walk through the process of building an Excel add-in by using jQuery and the Excel JavaScript API.</span></span> 

## <a name="create-the-add-in"></a><span data-ttu-id="c6818-103">?????</span><span class="sxs-lookup"><span data-stu-id="c6818-103">Create the add-in</span></span> 

[!include[Choose your editor](../includes/quickstart-choose-editor.md)]

# <a name="visual-studiotabvisual-studio"></a>[<span data-ttu-id="c6818-104">Visual Studio</span><span class="sxs-lookup"><span data-stu-id="c6818-104">Visual Studio</span></span>](#tab/visual-studio)

### <a name="prerequisites"></a><span data-ttu-id="c6818-105">????</span><span class="sxs-lookup"><span data-stu-id="c6818-105">Prerequisites</span></span>

[!include[Quickstart prerequisites](../includes/quickstart-vs-prerequisites.md)]

### <a name="create-the-add-in-project"></a><span data-ttu-id="c6818-106">???????</span><span class="sxs-lookup"><span data-stu-id="c6818-106">Create the add-in project</span></span>

1. <span data-ttu-id="c6818-107">? Visual Studio ?????????????**** > ????**** > ????****?</span><span class="sxs-lookup"><span data-stu-id="c6818-107">On the Visual Studio menu bar, choose  **File** > **New** > **Project**.</span></span>
    
2. <span data-ttu-id="c6818-108">??Visual C#?****??Visual Basic?****?????????????Office/SharePoint?****????????****?????Excel Web ????****???????</span><span class="sxs-lookup"><span data-stu-id="c6818-108">In the list of project types under **Visual C#** or **Visual Basic**, expand  **Office/SharePoint**, choose **Add-ins**, and then choose **Excel Web Add-in** as the project type.</span></span> 

3. <span data-ttu-id="c6818-109">?????????????****?</span><span class="sxs-lookup"><span data-stu-id="c6818-109">Name the project, and then choose **OK**.</span></span>

4. <span data-ttu-id="c6818-110">???? Office ????****????????????????? Excel?****????????****??????</span><span class="sxs-lookup"><span data-stu-id="c6818-110">In the **Create Office Add-in** dialog window, choose **Add new functionalities to Excel**, and then choose **Finish** to create the project.</span></span>

5. <span data-ttu-id="c6818-p101">???Visual Studio ????????????????????????????****??**Home.html** ??? Visual Studio ????</span><span class="sxs-lookup"><span data-stu-id="c6818-p101">Visual Studio creates a solution and its two projects appear in **Solution Explorer**. The **Home.html** file opens in Visual Studio.</span></span>
    
### <a name="explore-the-visual-studio-solution"></a><span data-ttu-id="c6818-113">?? Visual Studio ????</span><span class="sxs-lookup"><span data-stu-id="c6818-113">Explore the Visual Studio solution</span></span>

[!include[Description of Visual Studio projects](../includes/quickstart-vs-solution.md)]

### <a name="update-the-code"></a><span data-ttu-id="c6818-114">????</span><span class="sxs-lookup"><span data-stu-id="c6818-114">Update the code</span></span>

1. <span data-ttu-id="c6818-115">**Home.html** ??????????????? HTML?</span><span class="sxs-lookup"><span data-stu-id="c6818-115">**Home.html** specifies the HTML that will be rendered in the add-in's task pane.</span></span> <span data-ttu-id="c6818-116">? **Home.html** ??? `<body>` ????????????????</span><span class="sxs-lookup"><span data-stu-id="c6818-116">In **Home.html**, replace the `<body>` element with the following markup and save the file.</span></span>
 
    ```html
    <body class="ms-font-m ms-welcome">
        <div id="content-header">
            <div class="padding">
                <h1>Welcome</h1>
            </div>
        </div>
        <div id="content-main">
            <div class="padding">
                <p>Choose the button below to set the color of the selected range to green.</p>
                <br />
                <h3>Try it out</h3>
                <button class="ms-Button" id="set-color">Set color</button>
            </div>
        </div>
    </body>
    ```

2. <span data-ttu-id="c6818-117">?? Web ????????????Home.js?****?</span><span class="sxs-lookup"><span data-stu-id="c6818-117">Open the file **Home.js** in the root of the web application project.</span></span> <span data-ttu-id="c6818-118">???????????</span><span class="sxs-lookup"><span data-stu-id="c6818-118">This file specifies the script for the add-in.</span></span> <span data-ttu-id="c6818-119">???????????????????</span><span class="sxs-lookup"><span data-stu-id="c6818-119">Replace the entire contents with the following code and save the file.</span></span> 

    ```js
    'use strict';

    (function () {
        Office.initialize = function (reason) {
            $(document).ready(function () {
                $('#set-color').click(setColor);
            });
        };

        function setColor() {
            Excel.run(function (context) {
                var range = context.workbook.getSelectedRange();
                range.format.fill.color = 'green';

                return context.sync();
            }).catch(function (error) {
                console.log("Error: " + error);
                if (error instanceof OfficeExtension.Error) {
                    console.log("Debug info: " + JSON.stringify(error.debugInfo));
                }
            });
        }
    })();
    ```

3. <span data-ttu-id="c6818-120">?? Web ????????????Home.css?****?</span><span class="sxs-lookup"><span data-stu-id="c6818-120">Open the file **Home.css** in the root of the web application project.</span></span> <span data-ttu-id="c6818-121">??????????????</span><span class="sxs-lookup"><span data-stu-id="c6818-121">This file specifies the custom styles for the add-in.</span></span> <span data-ttu-id="c6818-122">???????????????????</span><span class="sxs-lookup"><span data-stu-id="c6818-122">Replace the entire contents with the following code and save the file.</span></span> 

    ```css
    #content-header {
        background: #2a8dd4;
        color: #fff;
        position: absolute;
        top: 0;
        left: 0;
        width: 100%;
        height: 80px; 
        overflow: hidden;
    }

    #content-main {
        background: #fff;
        position: fixed;
        top: 80px;
        left: 0;
        right: 0;
        bottom: 0;
        overflow: auto; 
    }

    .padding {
        padding: 15px;
    }
    ```

### <a name="update-the-manifest"></a><span data-ttu-id="c6818-123">????</span><span class="sxs-lookup"><span data-stu-id="c6818-123">Update the manifest</span></span>

1. <span data-ttu-id="c6818-124">????????? XML ?????</span><span class="sxs-lookup"><span data-stu-id="c6818-124">Open the XML manifest file in the Add-in project.</span></span> <span data-ttu-id="c6818-125">????????????????</span><span class="sxs-lookup"><span data-stu-id="c6818-125">This file defines the add-in's settings and capabilities.</span></span>

2. <span data-ttu-id="c6818-126">?????????`ProviderName`</span><span class="sxs-lookup"><span data-stu-id="c6818-126">The `ProviderName` element has a placeholder value.</span></span> <span data-ttu-id="c6818-127">??????????</span><span class="sxs-lookup"><span data-stu-id="c6818-127">Replace it with your name.</span></span>

3. <span data-ttu-id="c6818-128">??? `DefaultValue` ???????`DisplayName`</span><span class="sxs-lookup"><span data-stu-id="c6818-128">The `DefaultValue` attribute of the `DisplayName` element has a placeholder.</span></span> <span data-ttu-id="c6818-129">??????My Office Add-in?****?</span><span class="sxs-lookup"><span data-stu-id="c6818-129">Replace it with **My Office Add-in**.</span></span>

4. <span data-ttu-id="c6818-130">??? `DefaultValue` ???????`Description`</span><span class="sxs-lookup"><span data-stu-id="c6818-130">The `DefaultValue` attribute of the `Description` element has a placeholder.</span></span> <span data-ttu-id="c6818-131">??????A task pane add-in for Excel?****?</span><span class="sxs-lookup"><span data-stu-id="c6818-131">Replace it with **A task pane add-in for Excel**.</span></span>

5. <span data-ttu-id="c6818-132">?????</span><span class="sxs-lookup"><span data-stu-id="c6818-132">Save the file.</span></span>

    ```xml
    ...
    <ProviderName>John Doe</ProviderName>
    <DefaultLocale>en-US</DefaultLocale>
    <!-- The display name of your add-in. Used on the store and various places of the Office UI such as the add-ins dialog. -->
    <DisplayName DefaultValue="My Office Add-in" />
    <Description DefaultValue="A task pane add-in for Excel"/>
    ...
    ```

### <a name="try-it-out"></a><span data-ttu-id="c6818-133">??</span><span class="sxs-lookup"><span data-stu-id="c6818-133">Try it out</span></span>

1. <span data-ttu-id="c6818-p109">?? Visual Studio ????? F5 ???????****???? Excel??????? Excel ???????????????????****?????????????? IIS ??</span><span class="sxs-lookup"><span data-stu-id="c6818-p109">Using Visual Studio, test the newly created Excel add-in by pressing F5 or choosing the **Start** button to launch Excel with the **Show Taskpane** add-in button displayed in the ribbon. The add-in will be hosted locally on IIS.</span></span>

2. <span data-ttu-id="c6818-136">? Excel ??????????****?????????????????****?????????????</span><span class="sxs-lookup"><span data-stu-id="c6818-136">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Excel ?????](../images/excel-quickstart-addin-2a.png)

3. <span data-ttu-id="c6818-138">????????????????</span><span class="sxs-lookup"><span data-stu-id="c6818-138">Select any range of cells in the worksheet.</span></span>

4. <span data-ttu-id="c6818-139">???????????????****?????????????????</span><span class="sxs-lookup"><span data-stu-id="c6818-139">In the task pane, choose the **Set color** button to set the color of the selected range to green.</span></span>

    ![Excel ???](../images/excel-quickstart-addin-2c.png)

# <a name="any-editortabvisual-studio-code"></a>[<span data-ttu-id="c6818-141">?????</span><span class="sxs-lookup"><span data-stu-id="c6818-141">Any editor</span></span>](#tab/visual-studio-code)

### <a name="prerequisites"></a><span data-ttu-id="c6818-142">????</span><span class="sxs-lookup"><span data-stu-id="c6818-142">Prerequisites</span></span>

- [<span data-ttu-id="c6818-143">Node.js</span><span class="sxs-lookup"><span data-stu-id="c6818-143">Node.js</span></span>](https://nodejs.org)

- <span data-ttu-id="c6818-144">??????? [Yeoman](https://github.com/yeoman/yo) ? [Office ???? Yeoman ???](https://github.com/OfficeDev/generator-office)?</span><span class="sxs-lookup"><span data-stu-id="c6818-144">Install the latest version of [Yeoman](https://github.com/yeoman/yo) and the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office) globally.</span></span>

    ```bash
    npm install -g yo generator-office
    ```

### <a name="create-the-web-app"></a><span data-ttu-id="c6818-145">?? Web ??</span><span class="sxs-lookup"><span data-stu-id="c6818-145">Create the web app</span></span>

1. <span data-ttu-id="c6818-146">????????????????????my-addin?****?</span><span class="sxs-lookup"><span data-stu-id="c6818-146">Create a folder on your local drive and name it **my-addin**.</span></span> <span data-ttu-id="c6818-147">?????????????</span><span class="sxs-lookup"><span data-stu-id="c6818-147">This is where you'll create the files for your app.</span></span>

2. <span data-ttu-id="c6818-148">??????????</span><span class="sxs-lookup"><span data-stu-id="c6818-148">Navigate to your app folder.</span></span>

    ```bash
    cd my-addin
    ```

3. <span data-ttu-id="c6818-149">?? Yeoman ??????????????</span><span class="sxs-lookup"><span data-stu-id="c6818-149">Use the Yeoman generator to generate the manifest file for your add-in.</span></span> <span data-ttu-id="c6818-150">??????????????????????????</span><span class="sxs-lookup"><span data-stu-id="c6818-150">Run the following command and then answer the prompts as shown in the following screenshot:</span></span>

    ```bash
    yo office
    ```

    - <span data-ttu-id="c6818-151">**??????????????** `No`</span><span class="sxs-lookup"><span data-stu-id="c6818-151">**Would you like to create a new subfolder for your project?:** `No`</span></span>
    - <span data-ttu-id="c6818-152">**??????????????:** `My Office Add-in`</span><span class="sxs-lookup"><span data-stu-id="c6818-152">**What do you want to name your add-in?:** `My Office Add-in`</span></span>
    - <span data-ttu-id="c6818-153">**?????? Office ????????:** `Excel`</span><span class="sxs-lookup"><span data-stu-id="c6818-153">**Which Office client application would you like to support?:** `Excel`</span></span>
    - <span data-ttu-id="c6818-154">**??????????:** `Yes`</span><span class="sxs-lookup"><span data-stu-id="c6818-154">**Would you like to create a new add-in?:** `Yes`</span></span>
    - <span data-ttu-id="c6818-155">**????? TypeScript?:** `No`</span><span class="sxs-lookup"><span data-stu-id="c6818-155">**Would you like to use TypeScript?:** `No`</span></span>
    - <span data-ttu-id="c6818-156">**?????** `Jquery`</span><span class="sxs-lookup"><span data-stu-id="c6818-156">**Choose a framework:** `Jquery`</span></span>

    <span data-ttu-id="c6818-p112">???????????????resource.html?****???????????????????????????????????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="c6818-p112">The generator will then ask you if you want to open **resource.html**. It isn't necessary to open it for this tutorial, but feel free to open it if you're curious! Choose yes or no to complete the wizard and allow the generator to do its work.</span></span>

    ![Yeoman ???](../images/yo-office-jquery.png)


4. <span data-ttu-id="c6818-161">????????????????? **index.html**?</span><span class="sxs-lookup"><span data-stu-id="c6818-161">In your code editor, open **index.html** in the root of the project.</span></span> <span data-ttu-id="c6818-162">????????????????? HTML?</span><span class="sxs-lookup"><span data-stu-id="c6818-162">This file specifies the HTML that will be rendered in the add-in's task pane.</span></span> 
 
5. <span data-ttu-id="c6818-163">? **index.html** ?????? `header` ??????????</span><span class="sxs-lookup"><span data-stu-id="c6818-163">Within **index.html**, replace the generated `header` tag with the following markup.</span></span>
 
    ```html
    <div id="content-header">
        <div class="padding">
            <h1>Welcome</h1>
        </div>
    </div>
    ```

6. <span data-ttu-id="c6818-164">? **index.html** ?????? `main` ????????????????</span><span class="sxs-lookup"><span data-stu-id="c6818-164">Within **index.html**, replace the generated `main` tag with the following markup, and save the file.</span></span>

    ```html
    <div id="content-main">
        <div class="padding">
            <p>Choose the button below to set the color of the selected range to green.</p>
            <br />
            <h3>Try it out</h3>
            <button class="ms-Button" id="set-color">Set color</button>
        </div>
    </div>
    ```

7. <span data-ttu-id="c6818-165">?????app.js?****??????????</span><span class="sxs-lookup"><span data-stu-id="c6818-165">Open the file **app.js** to specify the script for the add-in.</span></span> <span data-ttu-id="c6818-166">???????????????????</span><span class="sxs-lookup"><span data-stu-id="c6818-166">Replace the entire contents with the following code and save the file.</span></span>

    ```js
    'use strict';
    
    (function () {
        Office.initialize = function (reason) {
            $(document).ready(function () {
                $('#set-color').click(setColor);
            });
        };

        function setColor() {
            Excel.run(function (context) {
                var range = context.workbook.getSelectedRange();
                range.format.fill.color = 'green';

                return context.sync();
            }).catch(function (error) {
                console.log("Error: " + error);
                if (error instanceof OfficeExtension.Error) {
                    console.log("Debug info: " + JSON.stringify(error.debugInfo));
                }
            });
        }
    })();
    ```

8. <span data-ttu-id="c6818-167">?????app.css?****?????????????</span><span class="sxs-lookup"><span data-stu-id="c6818-167">Open the file **app.css** to specify the custom styles for the add-in.</span></span> <span data-ttu-id="c6818-168">???????????????????</span><span class="sxs-lookup"><span data-stu-id="c6818-168">Replace the entire contents with the following code and save the file.</span></span>

    ```css
    #content-header {
        background: #2a8dd4;
        color: #fff;
        position: absolute;
        top: 0;
        left: 0;
        width: 100%;
        height: 80px; 
        overflow: hidden;
    }

    #content-main {
        background: #fff;
        position: fixed;
        top: 80px;
        left: 0;
        right: 0;
        bottom: 0;
        overflow: auto; 
    }

    .padding {
        padding: 15px;
    }
    ```

### <a name="update-the-manifest"></a><span data-ttu-id="c6818-169">????</span><span class="sxs-lookup"><span data-stu-id="c6818-169">Update the manifest</span></span>

1. <span data-ttu-id="c6818-170">?????my-office-add-in-manifest.xml?****??????????????</span><span class="sxs-lookup"><span data-stu-id="c6818-170">Open the file **my-office-add-in-manifest.xml** to define the add-in's settings and capabilities.</span></span> 

2. <span data-ttu-id="c6818-171">?????????`ProviderName`</span><span class="sxs-lookup"><span data-stu-id="c6818-171">The `ProviderName` element has a placeholder value.</span></span> <span data-ttu-id="c6818-172">??????????</span><span class="sxs-lookup"><span data-stu-id="c6818-172">Replace it with your name.</span></span>

3. <span data-ttu-id="c6818-173">??? `DefaultValue` ???????`DisplayName`</span><span class="sxs-lookup"><span data-stu-id="c6818-173">The `DefaultValue` attribute of the `DisplayName` element has a placeholder.</span></span> <span data-ttu-id="c6818-174">??????My Office Add-in?****?</span><span class="sxs-lookup"><span data-stu-id="c6818-174">Replace it with **My Office Add-in**.</span></span>

4. <span data-ttu-id="c6818-175">??? `DefaultValue` ???????`Description`</span><span class="sxs-lookup"><span data-stu-id="c6818-175">The `DefaultValue` attribute of the `Description` element has a placeholder.</span></span> <span data-ttu-id="c6818-176">??????A task pane add-in for Excel?****?</span><span class="sxs-lookup"><span data-stu-id="c6818-176">Replace it with **A task pane add-in for Excel**.</span></span>

5. <span data-ttu-id="c6818-177">?????</span><span class="sxs-lookup"><span data-stu-id="c6818-177">Save the file.</span></span>

    ```xml
    ...
    <ProviderName>John Doe</ProviderName>
    <DefaultLocale>en-US</DefaultLocale>
    <!-- The display name of your add-in. Used on the store and various places of the Office UI such as the add-ins dialog. -->
    <DisplayName DefaultValue="My Office Add-in" />
    <Description DefaultValue="A task pane add-in for Excel"/>
    ...
    ```

### <a name="start-the-dev-server"></a><span data-ttu-id="c6818-178">?????????</span><span class="sxs-lookup"><span data-stu-id="c6818-178">Start the dev server</span></span>

[!include[Start server section](../includes/quickstart-yo-start-server.md)] 

### <a name="try-it-out"></a><span data-ttu-id="c6818-179">??</span><span class="sxs-lookup"><span data-stu-id="c6818-179">Try it out</span></span>

1. <span data-ttu-id="c6818-180">?????????????????????? Excel ????????</span><span class="sxs-lookup"><span data-stu-id="c6818-180">Follow the instructions for the platform you'll use to run your add-in to sideload the add-in within Excel.</span></span>

    - <span data-ttu-id="c6818-181">Windows?[? Windows ???? Office ???](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span><span class="sxs-lookup"><span data-stu-id="c6818-181">Windows: [Sideload Office Add-ins on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span></span>
    - <span data-ttu-id="c6818-182">Excel Online?[? Office Online ???? Office ???](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)</span><span class="sxs-lookup"><span data-stu-id="c6818-182">Excel Online: [Sideload Office Add-ins in Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)</span></span>
    - <span data-ttu-id="c6818-183">iPad ? Mac?[? iPad ? Mac ???? Office ???](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span><span class="sxs-lookup"><span data-stu-id="c6818-183">iPad and Mac: [Sideload Office Add-ins on iPad and Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span></span>

2. <span data-ttu-id="c6818-184">? Excel ??????????****?????????????????****??????????????</span><span class="sxs-lookup"><span data-stu-id="c6818-184">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Excel ?????](../images/excel-quickstart-addin-2b.png)

3. <span data-ttu-id="c6818-186">????????????????</span><span class="sxs-lookup"><span data-stu-id="c6818-186">Select any range of cells in the worksheet.</span></span>

4. <span data-ttu-id="c6818-187">???????????????****?????????????????</span><span class="sxs-lookup"><span data-stu-id="c6818-187">In the task pane, choose the **Set color** button to set the color of the selected range to green.</span></span>

    ![Excel ???](../images/excel-quickstart-addin-2c.png)

---

## <a name="next-steps"></a><span data-ttu-id="c6818-189">????</span><span class="sxs-lookup"><span data-stu-id="c6818-189">Next steps</span></span>

<span data-ttu-id="c6818-p119">?????? jQuery ???? Excel ????????????? Excel ????????? Excel ????????????????????</span><span class="sxs-lookup"><span data-stu-id="c6818-p119">Congratulations, you've successfully created an Excel add-in using jQuery! Next, learn more about the capabilities of an Excel add-in and build a more complex add-in by following along with the Excel add-in tutorial.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="c6818-192">Excel ?????</span><span class="sxs-lookup"><span data-stu-id="c6818-192">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial.yml)

## <a name="see-also"></a><span data-ttu-id="c6818-193">????</span><span class="sxs-lookup"><span data-stu-id="c6818-193">See also</span></span>

* [<span data-ttu-id="c6818-194">Excel ?????</span><span class="sxs-lookup"><span data-stu-id="c6818-194">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial-create-table.md)
* [<span data-ttu-id="c6818-195">Excel JavaScript API ????</span><span class="sxs-lookup"><span data-stu-id="c6818-195">Excel JavaScript API core concepts</span></span>](../excel/excel-add-ins-core-concepts.md)
* [<span data-ttu-id="c6818-196">Excel ???????</span><span class="sxs-lookup"><span data-stu-id="c6818-196">Excel add-in code samples</span></span>](http://dev.office.com/code-samples#?filters=excel,office%20add-ins)
* [<span data-ttu-id="c6818-197">Excel JavaScript API ??</span><span class="sxs-lookup"><span data-stu-id="c6818-197">Excel JavaScript API reference</span></span>](https://dev.office.com/reference/add-ins/excel/excel-add-ins-reference-overview)
