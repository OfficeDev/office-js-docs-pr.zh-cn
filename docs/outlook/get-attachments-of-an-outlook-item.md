---
title: 获取 Outlook 加载项中的附件
description: 加载项可使用附件 API 将与附件相关的信息发送至远程服务。
ms.date: 01/14/2021
localization_priority: Normal
ms.openlocfilehash: b572893e93c747e155f643e99c0a3a67c323e5e6e12be9fa996adefc5de47780
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/07/2021
ms.locfileid: "57095172"
---
# <a name="get-attachments-of-an-outlook-item-from-the-server"></a>从服务器获取 Outlook 项的附件

可以通过几种方法获取Outlook项的附件，但具体使用哪个选项取决于你的方案。

1. 将附件信息发送到远程服务。

    外接程序可以使用附件 API 将有关附件的信息发送到远程服务。 然后，该服务可以直接联系 Exchange 服务器来检索附件。

1. 使用要求集 1.8 中提供的 [getAttachmentContentAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) API。 支持的格式 [：AttachmentContentFormat](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat)。

    如果 EWS/REST (例如，由于 Exchange 服务器) 的管理员配置，或者您的外接程序希望直接在 HTML 或 JavaScript 中使用 base64 内容，则此 API 可能很方便。 此外，该 API 在撰写方案中可用，其中附件可能尚未同步到 Exchange;有关详细信息，请参阅在 Outlook 中的撰写窗体中管理 `getAttachmentContentAsync` [项目的](add-and-remove-attachments-to-an-item-in-a-compose-form.md)附件。

本文详细介绍了第一个选项。 若要将附件信息发送到远程服务，请使用以下属性和函数。

- [Office.context.mailbox.ewsUrl](/javascript/api/outlook/office.entities) 属性 &ndash; 提供托管邮箱的 Exchange 服务器上 Exchange Web 服务 (EWS) 的 URL。服务使用此 URL 调用 [ExchangeService.GetAttachments](/exchange/client-developer/exchange-web-services/how-to-get-attachments-by-using-ews-in-exchange) 方法或 [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) EWS 操作。

- [Office.context.mailbox.item.attachments](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) 属性 &ndash; 获取一组 [AttachmentDetails](/javascript/api/outlook/office.attachmentdetails) 对象，每个对象对应于一个项附件。

- [Office.context.mailbox.getCallbackTokenAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) 函数 &ndash; 对托管邮箱的 Exchange 服务执行异步调用，以获取服务器发回给 Exchange 服务器的回调令牌，从而验证附件请求。

## <a name="using-the-attachments-api"></a>使用附件 API

若要使用附件 API 从邮箱Exchange附件，请执行以下步骤。

1. 当用户查看包含附件的邮件或约会时显示外接程序。

1. 从 Exchange 服务器获取回调令牌。

1. 将回调令牌和附件信息发给远程服务。

1. 使用 `ExchangeService.GetAttachments` 方法或 `GetAttachment` 操作获取来自 Exchange 服务器的附件。

以下各节使用 [Outlook-Add-in-JavaScript-GetAttachments](https://github.com/OfficeDev/Outlook-Add-in-JavaScript-GetAttachments) 示例中的代码详细介绍了每个步骤。

> [!NOTE]
> 为强调附件信息，这些示例中的代码已被缩短。该示例包含用于向远程服务器进行外接程序身份验证和管理请求状态的额外代码。

## <a name="get-a-callback-token"></a>获取回调令牌

[Office.context.mailbox](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md) 对象提供 `getCallbackTokenAsync` 函数来获取远程服务器可以用来对 Exchange Server 进行身份验证的令牌。 以下代码显示了加载项中的一个函数，它启动异步请求来获取回调令牌以及用于获取响应的回调函数。 回调令牌存储在在下一部分定义的服务请求对象中。

```js
function getAttachmentToken() {
    if (serviceRequest.attachmentToken == "") {
        Office.context.mailbox.getCallbackTokenAsync(attachmentTokenCallback);
    }
}

function attachmentTokenCallback(asyncResult, userContext) {
    if (asyncResult.status === "succeeded") {
        // Cache the result from the server.
        serviceRequest.attachmentToken = asyncResult.value;
        serviceRequest.state = 3;
        testAttachments();
    } else {
        showToast("Error", "Could not get callback token: " + asyncResult.error.message);
    }
}
```

## <a name="send-attachment-information-to-the-remote-service"></a>向远程服务发送附件信息

外接程序调用的远程服务定义了应该如何向该服务发送附件信息的详情。在本示例中，远程服务是通过使用 Visual Studio 2013 创建的一个 Web API 应用程序。远程服务需要 JSON 对象中的附件信息。以下代码初始化一个包含附件信息的对象。

```js
// Initialize a context object for the add-in.
//   Set the fields that are used on the request
//   object to default values.
 var serviceRequest = {
    attachmentToken: '',
    ewsUrl         : Office.context.mailbox.ewsUrl,
    attachments    : []
 };
```

<br/>

`Office.context.mailbox.item.attachments` 属性包含 `AttachmentDetails` 对象的集合，每个对象对应于项的每个附件。 在大多数情况下，加载项可以只将 `AttachmentDetails` 对象的附件 ID 属性传递到远程服务。 如果远程服务需要有关附件的更多详细信息，可以传递全部或部分 `AttachmentDetails` 对象。 以下代码定义了将整个 `AttachmentDetails` 数组置于 `serviceRequest` 对象中并向远程服务发送请求的方法。

```js
function makeServiceRequest() {
  // Format the attachment details for sending.
  for (var i = 0; i < mailbox.item.attachments.length; i++) {
    serviceRequest.attachments[i] = JSON.parse(JSON.stringify(mailbox.item.attachments[i]));
  }

  $.ajax({
    url: '../../api/Default',
    type: 'POST',
    data: JSON.stringify(serviceRequest),
    contentType: 'application/json;charset=utf-8'
  }).done(function (response) {
    if (!response.isError) {
      var names = "<h2>Attachments processed using " +
                    serviceRequest.service +
                    ": " +
                    response.attachmentsProcessed +
                    "</h2>";
      for (i = 0; i < response.attachmentNames.length; i++) {
        names += response.attachmentNames[i] + "<br />";
      }
      document.getElementById("names").innerHTML = names;
    } else {
      app.showNotification("Runtime error", response.message);
    }
  }).fail(function (status) {

  }).always(function () {
    $('.disable-while-sending').prop('disabled', false);
  })
}
```

## <a name="get-the-attachments-from-the-exchange-server"></a>从 Exchange 服务器获取附件

您的远程服务可以使用 [GetAttachments](/exchange/client-developer/exchange-web-services/how-to-get-attachments-by-using-ews-in-exchange) EWS 托管 API 方法或 [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) EWS 操作从服务器检索附件。服务器应用程序需要两个对象将 JSON 字符串反序列化为可在服务器上使用的 .NET Framework 对象。以下代码显示了反序列化对象的定义。

```cs
namespace AttachmentsSample
{
  public class AttachmentSampleServiceRequest
  {
    public string attachmentToken { get; set; }
    public string ewsUrl { get; set; }
    public string service { get; set; }
    public AttachmentDetails [] attachments { get; set; }
  }

  public class AttachmentDetails
  {
    public string attachmentType { get; set; }
    public string contentType { get; set; }
    public string id { get; set; }
    public bool isInline { get; set; }
    public string name { get; set; }
    public int size { get; set; }
  }
}
```

### <a name="use-the-ews-managed-api-to-get-the-attachments"></a>使用 EWS 托管 API 来获取附件

如果使用远程服务中的 [EWS 托管 API](/exchange/client-developer/web-service-reference/ews-managed-api-reference-for-exchange)，则可以使用 [GetAttachments](/exchange/client-developer/exchange-web-services/how-to-get-attachments-by-using-ews-in-exchange) 方法构建、发送和接收 EWS SOAP 请求以获取附件。我们建议您使用 EWS 托管 API，因为它所需的代码行数较少，并可提供用于调用 EWS 的更为直观的界面。以下代码发出一个检索所有附件的请求，并返回已处理附件的数目和名称。

```cs
private AttachmentSampleServiceResponse GetAtttachmentsFromExchangeServerUsingEWSManagedApi(AttachmentSampleServiceRequest request)
{
  var attachmentsProcessedCount = 0;
  var attachmentNames = new List<string>();

  // Create an ExchangeService object, set the credentials and the EWS URL.
  ExchangeService service = new ExchangeService();
  service.Credentials = new OAuthCredentials(request.attachmentToken);
  service.Url = new Uri(request.ewsUrl);

  var attachmentIds = new List<string>();

  foreach (AttachmentDetails attachment in request.attachments)
  {
    attachmentIds.Add(attachment.id);
  }

  // Call the GetAttachments method to retrieve the attachments on the message.
  // This method results in a GetAttachments EWS SOAP request and response
  // from the Exchange server.
  var getAttachmentsResponse =
    service.GetAttachments(attachmentIds.ToArray(),
                            null,
                            new PropertySet(BasePropertySet.FirstClassProperties,
                                            ItemSchema.MimeContent));

  if (getAttachmentsResponse.OverallResult == ServiceResult.Success)
  {
    foreach (var attachmentResponse in getAttachmentsResponse)
    {
      attachmentNames.Add(attachmentResponse.Attachment.Name);

      // Write the content of each attachment to a stream.
      if (attachmentResponse.Attachment is FileAttachment)
      {
        FileAttachment fileAttachment = attachmentResponse.Attachment as FileAttachment;
        Stream s = new MemoryStream(fileAttachment.Content);
        // Process the contents of the attachment here.
      }

      if (attachmentResponse.Attachment is ItemAttachment)
      {
        ItemAttachment itemAttachment = attachmentResponse.Attachment as ItemAttachment;
        Stream s = new MemoryStream(itemAttachment.Item.MimeContent.Content);
        // Process the contents of the attachment here.
      }

      attachmentsProcessedCount++;
    }
  }

  // Return the names and number of attachments processed for display
  // in the add-in UI.
  var response = new AttachmentSampleServiceResponse();
  response.attachmentNames = attachmentNames.ToArray();
  response.attachmentsProcessed = attachmentsProcessedCount;

  return response;
}
```

### <a name="use-ews-to-get-the-attachments"></a>使用 EWS 获取附件

如果在远程服务中使用 EWS，则需要构造 [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) SOAP 请求，从 Exchange Server 获取附件。 下面的代码返回了提供 SOAP 请求的字符串。 远程服务使用 `String.Format` 方法将附件的附件 ID 插入字符串。


```cs
private const string GetAttachmentSoapRequest =
@"<?xml version=""1.0"" encoding=""utf-8""?>
<soap:Envelope xmlns:xsi=""https://www.w3.org/2001/XMLSchema-instance""
xmlns:xsd=""https://www.w3.org/2001/XMLSchema""
xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/""
xmlns:t=""http://schemas.microsoft.com/exchange/services/2006/types"">
<soap:Header>
<t:RequestServerVersion Version=""Exchange2013"" />
</soap:Header>
  <soap:Body>
    <GetAttachment xmlns=""http://schemas.microsoft.com/exchange/services/2006/messages""
    xmlns:t=""http://schemas.microsoft.com/exchange/services/2006/types"">
      <AttachmentShape/>
      <AttachmentIds>
        <t:AttachmentId Id=""{0}""/>
      </AttachmentIds>
    </GetAttachment>
  </soap:Body>
</soap:Envelope>";
```

<br/>

最后，下列方法使用 EWS `GetAttachment` 请求从 Exchange Server 获取附件。 此实现分别对各个附件发起请求并返回已处理附件的计数。 在单独的 `ProcessXmlResponse` 方法中处理每个响应，然后进行定义。

```cs
private AttachmentSampleServiceResponse GetAttachmentsFromExchangeServerUsingEWS(AttachmentSampleServiceRequest request)
{
  var attachmentsProcessedCount = 0;
  var attachmentNames = new List<string>();

  foreach (var attachment in request.attachments)
  {
    // Prepare a web request object.
    HttpWebRequest webRequest = WebRequest.CreateHttp(request.ewsUrl);
    webRequest.Headers.Add("Authorization",
      string.Format("Bearer {0}", request.attachmentToken));
    webRequest.PreAuthenticate = true;
    webRequest.AllowAutoRedirect = false;
    webRequest.Method = "POST";
    webRequest.ContentType = "text/xml; charset=utf-8";

    // Construct the SOAP message for the GetAttachment operation.
    byte[] bodyBytes = Encoding.UTF8.GetBytes(
      string.Format(GetAttachmentSoapRequest, attachment.id));
    webRequest.ContentLength = bodyBytes.Length;

    Stream requestStream = webRequest.GetRequestStream();
    requestStream.Write(bodyBytes, 0, bodyBytes.Length);
    requestStream.Close();

    // Make the request to the Exchange server and get the response.
    HttpWebResponse webResponse = (HttpWebResponse)webRequest.GetResponse();

    // If the response is okay, create an XML document from the response
    // and process the request.
    if (webResponse.StatusCode == HttpStatusCode.OK)
    {
      var responseStream = webResponse.GetResponseStream();

      var responseEnvelope = XElement.Load(responseStream);

      // After creating a memory stream containing the contents of the
      // attachment, this method writes the XML document to the trace output.
      // Your service would perform it's processing here.
      if (responseEnvelope != null)
      {
        var processResult = ProcessXmlResponse(responseEnvelope);
        attachmentNames.Add(string.Format("{0} {1}", attachment.name, processResult));

      }

      // Close the response stream.
      responseStream.Close();
      webResponse.Close();

    }
    // If the response is not OK, return an error message for the
    // attachment.
    else
    {
      var errorString = string.Format("Attachment \"{0}\" could not be processed. " +
        "Error message: {1}.", attachment.name, webResponse.StatusDescription);
      attachmentNames.Add(errorString);
    }
    attachmentsProcessedCount++;
  }

  // Return the names and number of attachments processed for display
  // in the add-in UI.
  var response = new AttachmentSampleServiceResponse();
  response.attachmentNames = attachmentNames.ToArray();
  response.attachmentsProcessed = attachmentsProcessedCount;

  return response;
}
```

<br/>

将 `GetAttachment` 操作的每个响应发送给 `ProcessXmlResponse` 方法。 此方法检查响应是否出错。 如果它未找到任何错误，则会处理文件附件和项附件。 `ProcessXmlResponse` 方法执行大量工作来处理附件。

```cs
// This method processes the response from the Exchange server.
// In your application the bulk of the processing occurs here.
private string ProcessXmlResponse(XElement responseEnvelope)
{
  // First, check the response for web service errors.
  var errorCodes = from errorCode in responseEnvelope.Descendants
                    ("{http://schemas.microsoft.com/exchange/services/2006/messages}ResponseCode")
                    select errorCode;
  // Return the first error code found.
  foreach (var errorCode in errorCodes)
  {
    if (errorCode.Value != "NoError")
    {
      return string.Format("Could not process result. Error: {0}", errorCode.Value);
    }
  }

  // No errors found, proceed with processing the content.
  // First, get and process file attachments.
  var fileAttachments = from fileAttachment in responseEnvelope.Descendants
                    ("{http://schemas.microsoft.com/exchange/services/2006/types}FileAttachment")
                        select fileAttachment;
  foreach(var fileAttachment in fileAttachments)
  {
    var fileContent = fileAttachment.Element("{http://schemas.microsoft.com/exchange/services/2006/types}Content");
    var fileData = System.Convert.FromBase64String(fileContent.Value);
    var s = new MemoryStream(fileData);
    // Process the file attachment here.
  }

  // Second, get and process item attachments.
  var itemAttachments = from itemAttachment in responseEnvelope.Descendants
                        ("{http://schemas.microsoft.com/exchange/services/2006/types}ItemAttachment")
                        select itemAttachment;
  foreach(var itemAttachment in itemAttachments)
  {
    var message = itemAttachment.Element("{http://schemas.microsoft.com/exchange/services/2006/types}Message");
    if (message != null)
    {
      // Process a message here.
      break;
    }
    var calendarItem = itemAttachment.Element("{http://schemas.microsoft.com/exchange/services/2006/types}CalendarItem");
    if (calendarItem != null)
    {
      // Process calendar item here.
      break;
    }
    var contact = itemAttachment.Element("{http://schemas.microsoft.com/exchange/services/2006/types}Contact");
    if (contact != null)
    {
      // Process contact here.
      break;
    }
    var task = itemAttachment.Element("{http://schemas.microsoft.com/exchange/services/2006/types}Tontact");
    if (task != null)
    {
      // Process task here.
      break;
    }
    var meetingMessage = itemAttachment.Element("{http://schemas.microsoft.com/exchange/services/2006/types}MeetingMessage");
    if (meetingMessage != null)
    {
      // Process meeting message here.
      break;
    }
    var meetingRequest = itemAttachment.Element("{http://schemas.microsoft.com/exchange/services/2006/types}MeetingRequest");
    if (meetingRequest != null)
    {
      // Process meeting request here.
      break;
    }
    var meetingResponse = itemAttachment.Element("{http://schemas.microsoft.com/exchange/services/2006/types}MeetingResponse");
    if (meetingResponse != null)
    {
      // Process meeting response here.
      break;
    }
    var meetingCancellation = itemAttachment.Element("{http://schemas.microsoft.com/exchange/services/2006/types}MeetingCancellation");
    if (meetingCancellation != null)
    {
      // Process meeting cancellation here.
      break;
    }
  }

  return string.Empty;
}
```

## <a name="see-also"></a>另请参阅

- [创建适用于阅读窗体的 Outlook 加载项](read-scenario.md)
- [在 Exchange 中浏览 EWS 托管 API、EWS 和 Web 服务](/exchange/client-developer/exchange-web-services/explore-the-ews-managed-api-ews-and-web-services-in-exchange)
- [EWS 托管 API 客户端应用程序入门](/exchange/client-developer/exchange-web-services/get-started-with-ews-managed-api-client-applications)
- [Outlook外接程序 SSO](https://github.com/OfficeDev/Outlook-Add-in-SSO)
