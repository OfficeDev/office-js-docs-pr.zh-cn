# <a name="build-your-first-word-add-in"></a><span data-ttu-id="b6f42-101">??????? Word ????</span><span class="sxs-lookup"><span data-stu-id="b6f42-101">Build your first Word add-in</span></span>

<span data-ttu-id="b6f42-102">_????Word 2016?Word for iPad?Word for Mac_</span><span class="sxs-lookup"><span data-stu-id="b6f42-102">_Applies to: Word 2016, Word for iPad, Word for Mac_</span></span>

<span data-ttu-id="b6f42-103">??????????? jQuery ? Word JavaScript API ?? Word ????</span><span class="sxs-lookup"><span data-stu-id="b6f42-103">In this article, you'll walk through the process of building a Word add-in by using jQuery and the Word JavaScript API.</span></span> 

## <a name="create-the-add-in"></a><span data-ttu-id="b6f42-104">?????</span><span class="sxs-lookup"><span data-stu-id="b6f42-104">Create the add-in</span></span> 

[!include[Choose your editor](../includes/quickstart-choose-editor.md)]

# <a name="visual-studiotabvisual-studio"></a>[<span data-ttu-id="b6f42-105">Visual Studio</span><span class="sxs-lookup"><span data-stu-id="b6f42-105">Visual Studio</span></span>](#tab/visual-studio)

### <a name="prerequisites"></a><span data-ttu-id="b6f42-106">????</span><span class="sxs-lookup"><span data-stu-id="b6f42-106">Prerequisites</span></span>

[!include[Quick Start prerequisites](../includes/quickstart-vs-prerequisites.md)]

### <a name="create-the-add-in-project"></a><span data-ttu-id="b6f42-107">???????</span><span class="sxs-lookup"><span data-stu-id="b6f42-107">Create the add-in project</span></span>

1. <span data-ttu-id="b6f42-108">? Visual Studio ?????????????**** > ????**** > ????****?</span><span class="sxs-lookup"><span data-stu-id="b6f42-108">On the Visual Studio menu bar, choose  **File** > **New** > **Project**.</span></span>
    
2. <span data-ttu-id="b6f42-109">??Visual C#?****??Visual Basic?****?????????????Office/SharePoint?****????????****?????Word Web ????****???????</span><span class="sxs-lookup"><span data-stu-id="b6f42-109">In the list of project types under **Visual C#** or **Visual Basic**, expand  **Office/SharePoint**, choose **Add-ins**, and then choose **Word Web Add-in** as the project type.</span></span> 

3. <span data-ttu-id="b6f42-110">?????????????****?</span><span class="sxs-lookup"><span data-stu-id="b6f42-110">Name the project, and then choose **OK**.</span></span>

4. <span data-ttu-id="b6f42-p101">???Visual Studio ????????????????????????????****??**Home.html** ??? Visual Studio ????</span><span class="sxs-lookup"><span data-stu-id="b6f42-p101">Visual Studio creates a solution and its two projects appear in **Solution Explorer**. The **Home.html** file opens in Visual Studio.</span></span>
    
### <a name="explore-the-visual-studio-solution"></a><span data-ttu-id="b6f42-113">?? Visual Studio ????</span><span class="sxs-lookup"><span data-stu-id="b6f42-113">Explore the Visual Studio solution</span></span>

[!include[Description of Visual Studio projects](../includes/quickstart-vs-solution.md)]

### <a name="update-the-code"></a><span data-ttu-id="b6f42-114">????</span><span class="sxs-lookup"><span data-stu-id="b6f42-114">Update the code</span></span>

1. <span data-ttu-id="b6f42-115">**Home.html** ??????????????? HTML?</span><span class="sxs-lookup"><span data-stu-id="b6f42-115">**Home.html** specifies the HTML that will be rendered in the add-in's task pane.</span></span> <span data-ttu-id="b6f42-116">? **Home.html** ??? `<body>` ????????????????</span><span class="sxs-lookup"><span data-stu-id="b6f42-116">In **Home.html**, replace the `<body>` element with the following markup and save the file.</span></span>
 
    ```html
    <body>
        <div id="content-header">
            <div class="padding">
                <h1>Welcome</h1>
            </div>
        </div>    
        <div id="content-main">
            <div class="padding">
                <p>Choose the buttons below to add boilerplate text to the document by using the Word JavaScript API.</p>
                <br />
                <h3>Try it out</h3>
                <button id="emerson">Add quote from Ralph Waldo Emerson</button>
                <br /><br />
                <button id="checkhov">Add quote from Anton Chekhov</button>
                <br /><br />
                <button id="proverb">Add Chinese proverb</button>
            </div>
        </div>
        <br />
        <div id="supportedVersion"/>
    </body>
    ```

2. <span data-ttu-id="b6f42-117">?? Web ????????????Home.js?****?</span><span class="sxs-lookup"><span data-stu-id="b6f42-117">Open the file **Home.js** in the root of the web application project.</span></span> <span data-ttu-id="b6f42-118">???????????</span><span class="sxs-lookup"><span data-stu-id="b6f42-118">This file specifies the script for the add-in.</span></span> <span data-ttu-id="b6f42-119">???????????????????</span><span class="sxs-lookup"><span data-stu-id="b6f42-119">Replace the entire contents with the following code and save the file.</span></span>

    ```js
    'use strict';
    
    (function () {

        // The initialize function is run each time the page is loaded.
        Office.initialize = function (reason) {
            $(document).ready(function () {

                // Use this to check whether the API is supported in the Word client.
                if (Office.context.requirements.isSetSupported('WordApi', 1.1)) {
                    // Do something that is only available via the new APIs
                    $('#emerson').click(insertEmersonQuoteAtSelection);
                    $('#checkhov').click(insertChekhovQuoteAtTheBeginning);
                    $('#proverb').click(insertChineseProverbAtTheEnd);
                    $('#supportedVersion').html('This code is using Word 2016 or greater.');
                }
                else {
                    // Just letting you know that this code will not work with your version of Word.
                    $('#supportedVersion').html('This code requires Word 2016 or greater.');
                }
            });
        };

        function insertEmersonQuoteAtSelection() {
            Word.run(function (context) {

                // Create a proxy object for the document.
                var thisDocument = context.document;

                // Queue a command to get the current selection.
                // Create a proxy range object for the selection.
                var range = thisDocument.getSelection();

                // Queue a command to replace the selected text.
                range.insertText('"Hitch your wagon to a star."\n', Word.InsertLocation.replace);

                // Synchronize the document state by executing the queued commands,
                // and return a promise to indicate task completion.
                return context.sync().then(function () {
                    console.log('Added a quote from Ralph Waldo Emerson.');
                });
            })
            .catch(function (error) {
                console.log('Error: ' + JSON.stringify(error));
                if (error instanceof OfficeExtension.Error) {
                    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
                }
            });
        }

        function insertChekhovQuoteAtTheBeginning() {
            Word.run(function (context) {

                // Create a proxy object for the document body.
                var body = context.document.body;

                // Queue a command to insert text at the start of the document body.
                body.insertText('"Knowledge is of no value unless you put it into practice."\n', Word.InsertLocation.start);

                // Synchronize the document state by executing the queued commands,
                // and return a promise to indicate task completion.
                return context.sync().then(function () {
                    console.log('Added a quote from Anton Chekhov.');
                });
            })
            .catch(function (error) {
                console.log('Error: ' + JSON.stringify(error));
                if (error instanceof OfficeExtension.Error) {
                    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
                }
            });
        }

        function insertChineseProverbAtTheEnd() {
            Word.run(function (context) {

                // Create a proxy object for the document body.
                var body = context.document.body;

                // Queue a command to insert text at the end of the document body.
                body.insertText('"To know the road ahead, ask those coming back."\n', Word.InsertLocation.end);

                // Synchronize the document state by executing the queued commands,
                // and return a promise to indicate task completion.
                return context.sync().then(function () {
                    console.log('Added a quote from a Chinese proverb.');
                });
            })
            .catch(function (error) {
                console.log('Error: ' + JSON.stringify(error));
                if (error instanceof OfficeExtension.Error) {
                    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
                }
            });
        }
    })();
    ```

3. <span data-ttu-id="b6f42-120">?? Web ????????????Home.css?****?</span><span class="sxs-lookup"><span data-stu-id="b6f42-120">Open the file **Home.css** in the root of the web application project.</span></span> <span data-ttu-id="b6f42-121">??????????????</span><span class="sxs-lookup"><span data-stu-id="b6f42-121">This file specifies the custom styles for the add-in.</span></span> <span data-ttu-id="b6f42-122">???????????????????</span><span class="sxs-lookup"><span data-stu-id="b6f42-122">Replace the entire contents with the following code and save the file.</span></span>

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

### <a name="update-the-manifest"></a><span data-ttu-id="b6f42-123">????</span><span class="sxs-lookup"><span data-stu-id="b6f42-123">Update the manifest</span></span>

1. <span data-ttu-id="b6f42-124">????????? XML ?????</span><span class="sxs-lookup"><span data-stu-id="b6f42-124">Open the XML manifest file in the Add-in project.</span></span> <span data-ttu-id="b6f42-125">????????????????</span><span class="sxs-lookup"><span data-stu-id="b6f42-125">This file defines the add-in's settings and capabilities.</span></span>

2. <span data-ttu-id="b6f42-126">?????????`ProviderName`</span><span class="sxs-lookup"><span data-stu-id="b6f42-126">The `ProviderName` element has a placeholder value.</span></span> <span data-ttu-id="b6f42-127">??????????</span><span class="sxs-lookup"><span data-stu-id="b6f42-127">Replace it with your name.</span></span>

3. <span data-ttu-id="b6f42-128">??? `DefaultValue` ???????`DisplayName`</span><span class="sxs-lookup"><span data-stu-id="b6f42-128">The `DefaultValue` attribute of the `DisplayName` element has a placeholder.</span></span> <span data-ttu-id="b6f42-129">??????My Office Add-in?****?</span><span class="sxs-lookup"><span data-stu-id="b6f42-129">Replace it with **My Office Add-in**.</span></span>

4. <span data-ttu-id="b6f42-130">??? `DefaultValue` ???????`Description`</span><span class="sxs-lookup"><span data-stu-id="b6f42-130">The `DefaultValue` attribute of the `Description` element has a placeholder.</span></span> <span data-ttu-id="b6f42-131">??????A task pane add-in for Word?****?</span><span class="sxs-lookup"><span data-stu-id="b6f42-131">Replace it with **A task pane add-in for Word**.</span></span>

5. <span data-ttu-id="b6f42-132">?????</span><span class="sxs-lookup"><span data-stu-id="b6f42-132">Save the file.</span></span>

    ```xml
    ...
    <ProviderName>John Doe</ProviderName>
    <DefaultLocale>en-US</DefaultLocale>
    <!-- The display name of your add-in. Used on the store and various places of the Office UI such as the add-ins dialog. -->
    <DisplayName DefaultValue="My Office Add-in" />
    <Description DefaultValue="A task pane add-in for Word"/>
    ...
    ```

### <a name="try-it-out"></a><span data-ttu-id="b6f42-133">??</span><span class="sxs-lookup"><span data-stu-id="b6f42-133">Try it out</span></span>

1. <span data-ttu-id="b6f42-p109">?? Visual Studio ????? F5 ???????****???? Word??????? Word ???????????????????****?????????????? IIS ??</span><span class="sxs-lookup"><span data-stu-id="b6f42-p109">Using Visual Studio, test the newly created Word add-in by pressing F5 or choosing the **Start** button to launch Word with the **Show Taskpane** add-in button displayed in the ribbon. The add-in will be hosted locally on IIS.</span></span>

2. <span data-ttu-id="b6f42-136">? Word ??????????****?????????????????****??????????????</span><span class="sxs-lookup"><span data-stu-id="b6f42-136">In Word, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![???????????????? Word ??????](../images/word-quickstart-addin-0.png)

3. <span data-ttu-id="b6f42-138">????????????????????????</span><span class="sxs-lookup"><span data-stu-id="b6f42-138">In the task pane, choose any of the buttons to add boilerplate text to the document.</span></span>

    ![????????? Word ???????](../images/word-quickstart-addin-1b.png)

# <a name="any-editortabvisual-studio-code"></a>[<span data-ttu-id="b6f42-140">?????</span><span class="sxs-lookup"><span data-stu-id="b6f42-140">Any editor</span></span>](#tab/visual-studio-code)

### <a name="prerequisites"></a><span data-ttu-id="b6f42-141">????</span><span class="sxs-lookup"><span data-stu-id="b6f42-141">Prerequisites</span></span>

- [<span data-ttu-id="b6f42-142">Node.js</span><span class="sxs-lookup"><span data-stu-id="b6f42-142">Node.js</span></span>](https://nodejs.org)

- <span data-ttu-id="b6f42-143">??????? [Yeoman](https://github.com/yeoman/yo) ? [Office ???? Yeoman ???](https://github.com/OfficeDev/generator-office)?</span><span class="sxs-lookup"><span data-stu-id="b6f42-143">Install the latest version of [Yeoman](https://github.com/yeoman/yo) and the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office) globally.</span></span>

    ```bash
    npm install -g yo generator-office
    ```

### <a name="create-the-add-in-project"></a><span data-ttu-id="b6f42-144">???????</span><span class="sxs-lookup"><span data-stu-id="b6f42-144">Create the add-in project</span></span>

1. <span data-ttu-id="b6f42-145">????????????????????`my-word-addin`??</span><span class="sxs-lookup"><span data-stu-id="b6f42-145">Create a folder on your local drive and name it `my-word-addin`.</span></span> <span data-ttu-id="b6f42-146">?????????????</span><span class="sxs-lookup"><span data-stu-id="b6f42-146">This is where you'll create the files for your add-in.</span></span>

2. <span data-ttu-id="b6f42-147">???????</span><span class="sxs-lookup"><span data-stu-id="b6f42-147">Navigate to your new folder.</span></span>

    ```bash
    cd my-word-addin
    ```

3. <span data-ttu-id="b6f42-148">?? Yeoman ????? Word ??????</span><span class="sxs-lookup"><span data-stu-id="b6f42-148">Use the Yeoman generator to create a Word add-in project.</span></span> <span data-ttu-id="b6f42-149">?????????????????????</span><span class="sxs-lookup"><span data-stu-id="b6f42-149">Run the following command and then answer the prompts as follows:</span></span>

    ```bash
    yo office
    ```

    - <span data-ttu-id="b6f42-150">**???????????????:** `No`</span><span class="sxs-lookup"><span data-stu-id="b6f42-150">**Would you like to create a new subfolder for your project?:** `No`</span></span>
    - <span data-ttu-id="b6f42-151">**??????????????:** `My Office Add-in`</span><span class="sxs-lookup"><span data-stu-id="b6f42-151">**What do you want to name your add-in?:** `My Office Add-in`</span></span>
    - <span data-ttu-id="b6f42-152">**?????? Office ????????:** `Word`</span><span class="sxs-lookup"><span data-stu-id="b6f42-152">**Which Office client application would you like to support?:** `Word`</span></span>
    - <span data-ttu-id="b6f42-153">**??????????:** `Yes`</span><span class="sxs-lookup"><span data-stu-id="b6f42-153">**Would you like to create a new add-in?:** `Yes`</span></span>
    - <span data-ttu-id="b6f42-154">**????? TypeScript?:** `No`</span><span class="sxs-lookup"><span data-stu-id="b6f42-154">**Would you like to use TypeScript?:** `No`</span></span>
    - <span data-ttu-id="b6f42-155">**?????** `Jquery`</span><span class="sxs-lookup"><span data-stu-id="b6f42-155">**Choose a framework:** `Jquery`</span></span>

    <span data-ttu-id="b6f42-p112">???????????????resource.html?****???????????????????????????????????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="b6f42-p112">The generator will then ask you if you want to open **resource.html**. It isn't necessary to open it for this tutorial, but feel free to open it if you're curious! Choose yes or no to complete the wizard and allow the generator to do its work.</span></span>

    ![?? Yeoman ?????????????](../images/yo-office-word-jquery.png)

### <a name="update-the-code"></a><span data-ttu-id="b6f42-160">????</span><span class="sxs-lookup"><span data-stu-id="b6f42-160">Update the code</span></span>

1. <span data-ttu-id="b6f42-161">??????????????????index.html?****?</span><span class="sxs-lookup"><span data-stu-id="b6f42-161">In your code editor, open **index.html** in the root of the project.</span></span> <span data-ttu-id="b6f42-162">????????????????? HTML?</span><span class="sxs-lookup"><span data-stu-id="b6f42-162">This file contains the HTML that will be rendered in the add-in's task pane.</span></span> <span data-ttu-id="b6f42-163">???????????????????</span><span class="sxs-lookup"><span data-stu-id="b6f42-163">Replace the entire contents with the following code and save the file.</span></span> <span data-ttu-id="b6f42-164">??????????????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="b6f42-164">This add-in will display three buttons and when any of the buttons are chosen, boilerplate text will be added to the document.</span></span>

    ```html
    <!DOCTYPE html>
    <html>
        <head>
            <meta charset="UTF-8" />
            <meta http-equiv="X-UA-Compatible" content="IE=Edge" />
            <title>Boilerplate text app</title>
            <script src="https://ajax.aspnetcdn.com/ajax/jQuery/jquery-2.1.4.min.js"></script>
            <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
            <script src="app.js" type="text/javascript"></script>
            <link href="app.css" rel="stylesheet" type="text/css" />
        </head>
        <body>
            <div id="content-header">
                <div class="padding">
                    <h1>Welcome</h1>
                </div>
            </div>    
            <div id="content-main">
                <div class="padding">
                    <p>Choose the buttons below to add boilerplate text to the document by using the Word JavaScript API.</p>
                    <br />
                    <h3>Try it out</h3>
                    <button id="emerson">Add quote from Ralph Waldo Emerson</button>
                    <br /><br />
                    <button id="checkhov">Add quote from Anton Chekhov</button>
                    <br /><br />
                    <button id="proverb">Add Chinese proverb</button>
                </div>
            </div>
            <br />
            <div id="supportedVersion"/>
        </body>
    </html>
    ```

2. <span data-ttu-id="b6f42-165">?????app.js?****??????????</span><span class="sxs-lookup"><span data-stu-id="b6f42-165">Open the file **app.js** to specify the script for the add-in.</span></span> <span data-ttu-id="b6f42-166">???????????????????</span><span class="sxs-lookup"><span data-stu-id="b6f42-166">Replace the entire contents with the following code and save the file.</span></span> <span data-ttu-id="b6f42-167">???????????????? Word ????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="b6f42-167">This script contains initialization code as well as the code that makes changes to the Word document, by inserting text into the document when a button is chosen.</span></span> 

    ```js
    'use strict';
    
    (function () {

        // The initialize function is run each time the page is loaded.
        Office.initialize = function (reason) {
            $(document).ready(function () {

                // Use this to check whether the API is supported in the Word client.
                if (Office.context.requirements.isSetSupported('WordApi', 1.1)) {
                    // Do something that is only available via the new APIs
                    $('#emerson').click(insertEmersonQuoteAtSelection);
                    $('#checkhov').click(insertChekhovQuoteAtTheBeginning);
                    $('#proverb').click(insertChineseProverbAtTheEnd);
                    $('#supportedVersion').html('This code is using Word 2016 or greater.');
                }
                else {
                    // Just letting you know that this code will not work with your version of Word.
                    $('#supportedVersion').html('This code requires Word 2016 or greater.');
                }
            });
        };

        function insertEmersonQuoteAtSelection() {
            Word.run(function (context) {

                // Create a proxy object for the document.
                var thisDocument = context.document;

                // Queue a command to get the current selection.
                // Create a proxy range object for the selection.
                var range = thisDocument.getSelection();

                // Queue a command to replace the selected text.
                range.insertText('"Hitch your wagon to a star."\n', Word.InsertLocation.replace);

                // Synchronize the document state by executing the queued commands,
                // and return a promise to indicate task completion.
                return context.sync().then(function () {
                    console.log('Added a quote from Ralph Waldo Emerson.');
                });
            })
            .catch(function (error) {
                console.log('Error: ' + JSON.stringify(error));
                if (error instanceof OfficeExtension.Error) {
                    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
                }
            });
        }

        function insertChekhovQuoteAtTheBeginning() {
            Word.run(function (context) {

                // Create a proxy object for the document body.
                var body = context.document.body;

                // Queue a command to insert text at the start of the document body.
                body.insertText('"Knowledge is of no value unless you put it into practice."\n', Word.InsertLocation.start);

                // Synchronize the document state by executing the queued commands,
                // and return a promise to indicate task completion.
                return context.sync().then(function () {
                    console.log('Added a quote from Anton Chekhov.');
                });
            })
            .catch(function (error) {
                console.log('Error: ' + JSON.stringify(error));
                if (error instanceof OfficeExtension.Error) {
                    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
                }
            });
        }

        function insertChineseProverbAtTheEnd() {
            Word.run(function (context) {

                // Create a proxy object for the document body.
                var body = context.document.body;

                // Queue a command to insert text at the end of the document body.
                body.insertText('"To know the road ahead, ask those coming back."\n', Word.InsertLocation.end);

                // Synchronize the document state by executing the queued commands,
                // and return a promise to indicate task completion.
                return context.sync().then(function () {
                    console.log('Added a quote from a Chinese proverb.');
                });
            })
            .catch(function (error) {
                console.log('Error: ' + JSON.stringify(error));
                if (error instanceof OfficeExtension.Error) {
                    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
                }
            });
        }
    })();
    ```

3. <span data-ttu-id="b6f42-168">????????????app.css?****?????????????</span><span class="sxs-lookup"><span data-stu-id="b6f42-168">Open the file **app.css** in the root of the project to specify the custom styles for the add-in.</span></span> <span data-ttu-id="b6f42-169">???????????????????</span><span class="sxs-lookup"><span data-stu-id="b6f42-169">Replace the entire contents with the following and save the file.</span></span>

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

### <a name="update-the-manifest"></a><span data-ttu-id="b6f42-170">????</span><span class="sxs-lookup"><span data-stu-id="b6f42-170">Update the manifest</span></span>

1. <span data-ttu-id="b6f42-171">?????my-office-add-in-manifest.xml?****??????????????</span><span class="sxs-lookup"><span data-stu-id="b6f42-171">Open the file **my-office-add-in-manifest.xml** to define the add-in's settings and capabilities.</span></span>

2. <span data-ttu-id="b6f42-172">?????????`ProviderName`</span><span class="sxs-lookup"><span data-stu-id="b6f42-172">The `ProviderName` element has a placeholder value.</span></span> <span data-ttu-id="b6f42-173">??????????</span><span class="sxs-lookup"><span data-stu-id="b6f42-173">Replace it with your name.</span></span>

3. <span data-ttu-id="b6f42-174">??? `DefaultValue` ???????`Description`</span><span class="sxs-lookup"><span data-stu-id="b6f42-174">The `DefaultValue` attribute of the `Description` element has a placeholder.</span></span> <span data-ttu-id="b6f42-175">??????A task pane add-in for Word?****?</span><span class="sxs-lookup"><span data-stu-id="b6f42-175">Replace it with **A task pane add-in for Word**.</span></span>

4. <span data-ttu-id="b6f42-176">?????</span><span class="sxs-lookup"><span data-stu-id="b6f42-176">Save the file.</span></span>

    ```xml
    ...
    <ProviderName>John Doe</ProviderName>
    <DefaultLocale>en-US</DefaultLocale>
    <!-- The display name of your add-in. Used on the store and various places of the Office UI such as the add-ins dialog. -->
    <DisplayName DefaultValue="My Office Add-in" />
    <Description DefaultValue="A task pane add-in for Word"/>
    ...
    ```

### <a name="start-the-dev-server"></a><span data-ttu-id="b6f42-177">?????????</span><span class="sxs-lookup"><span data-stu-id="b6f42-177">Start the dev server</span></span>

[!include[Start server section](../includes/quickstart-yo-start-server.md)] 

### <a name="try-it-out"></a><span data-ttu-id="b6f42-178">??</span><span class="sxs-lookup"><span data-stu-id="b6f42-178">Try it out</span></span>

1. <span data-ttu-id="b6f42-179">?????????????????????? Word ????????</span><span class="sxs-lookup"><span data-stu-id="b6f42-179">To sideload the add-in within Word, follow the instructions for the platform you'll use to run your add-in.</span></span>

    - <span data-ttu-id="b6f42-180">Windows?[? Windows ???? Office ???](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span><span class="sxs-lookup"><span data-stu-id="b6f42-180">Windows: [Sideload Office Add-ins on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span></span>
    - <span data-ttu-id="b6f42-181">Word Online?[? Office Online ???? Office ???](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)</span><span class="sxs-lookup"><span data-stu-id="b6f42-181">Word Online: [Sideload Office Add-ins in Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)</span></span>
    - <span data-ttu-id="b6f42-182">iPad ? Mac?[? iPad ? Mac ???? Office ???](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span><span class="sxs-lookup"><span data-stu-id="b6f42-182">iPad and Mac: [Sideload Office Add-ins on iPad and Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span></span>

2. <span data-ttu-id="b6f42-183">? Word ??????????****?????????????????****??????????????</span><span class="sxs-lookup"><span data-stu-id="b6f42-183">In Word, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![???????????????? Word ??????](../images/word-quickstart-addin-2.png)

3. <span data-ttu-id="b6f42-185">????????????????????????</span><span class="sxs-lookup"><span data-stu-id="b6f42-185">In the task pane, choose any of the buttons to add boilerplate text to the document.</span></span>

    ![????????? Word ???????](../images/word-quickstart-addin-1.png)

---

## <a name="next-steps"></a><span data-ttu-id="b6f42-187">????</span><span class="sxs-lookup"><span data-stu-id="b6f42-187">Next steps</span></span>

<span data-ttu-id="b6f42-188">?????? jQuery ???? Word ?????</span><span class="sxs-lookup"><span data-stu-id="b6f42-188">Congratulations, you've successfully created a Word add-in using jQuery!</span></span> <span data-ttu-id="b6f42-189">????????? Excel ??????????? Excel ??????????????????????</span><span class="sxs-lookup"><span data-stu-id="b6f42-189">Next, learn more about the capabilities of an Excel add-in and build a more complex add-in by following along with the Excel add-in tutorial.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="b6f42-190">Word ??????</span><span class="sxs-lookup"><span data-stu-id="b6f42-190">Word add-in tutorial</span></span>](../tutorials/word-tutorial.yml)

## <a name="see-also"></a><span data-ttu-id="b6f42-191">????</span><span class="sxs-lookup"><span data-stu-id="b6f42-191">See also</span></span>

* [<span data-ttu-id="b6f42-192">Word ?????</span><span class="sxs-lookup"><span data-stu-id="b6f42-192">Word add-ins overview</span></span>](../word/word-add-ins-programming-overview.md)
* [<span data-ttu-id="b6f42-193">Word ???????</span><span class="sxs-lookup"><span data-stu-id="b6f42-193">Word add-in code samples</span></span>](http://dev.office.com/code-samples#?filters=word,office%20add-ins)
* [<span data-ttu-id="b6f42-194">Word JavaScript API ??</span><span class="sxs-lookup"><span data-stu-id="b6f42-194">Word JavaScript API reference</span></span>](https://dev.office.com/reference/add-ins/word/word-add-ins-reference-overview)
