using DocumentFormat.OpenXml.Packaging;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using System;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;

namespace DocsOne.Controllers
{
    [ApiController]
    [Route("api/")]
    public class DocsController : ControllerBase
    {
     
        private readonly ILogger<DocsController> _logger;
        public DocsController(ILogger<DocsController> logger)
        {
            _logger = logger;
        }

        [HttpGet]
        [Route("")]
        public IActionResult Welcome()
        {
            return Ok("Welcome to the docs api");

        }
            [HttpGet]
        [Route("getdoc")]
        public HttpResponseMessage Get()
        {
            HttpRequestMessage httpRequestMessage;

            string fullPath = @"D:\VS\sample.docx";
            HttpResponseMessage responseMsg;
            byte[] docBytes = System.IO.File.ReadAllBytes(fullPath);
            var dataStream = new MemoryStream(docBytes);
            responseMsg = new HttpResponseMessage(HttpStatusCode.OK);
            responseMsg.Content = new StreamContent(dataStream);
            responseMsg.Content.Headers.ContentDisposition = new System.Net.Http.Headers.ContentDispositionHeaderValue("attachment");
            responseMsg.Content.Headers.ContentDisposition.FileName = "sample.docx";
            responseMsg.Content.Headers.ContentType = new MediaTypeHeaderValue("application/octet-stream");

            return responseMsg;

        }

        [HttpPost]
        [Route("toHtml")]
        public IActionResult docToHtml()
        {
                string htmlText = string.Empty;
                DocProcessor processor = new DocProcessor();
                try
                {
                    htmlText = processor.ParseDOCX();
                }
                catch (OpenXmlPackageException e)
                {
                    if (e.ToString().Contains("Invalid Hyperlink"))
                    {
                        htmlText = processor.ParseDOCX();
                    }
                }
                return Ok(htmlText);
        }

        [HttpPost]
        [Route("toDoc")]
        public IActionResult htmlToDoc()
        {
            string htmlText = string.Empty;
            DocProcessor processor = new DocProcessor();
            try
            {
                 processor.ConvertToDocx();
            }
            catch (OpenXmlPackageException e)
            {
                if (e.ToString().Contains("Invalid Hyperlink"))
                {
                    processor.ConvertToDocx();
                }
            }
            return Ok("Successful");
        }

        [HttpPost, DisableRequestSizeLimit]
        [Route("makehtml")]
        public IActionResult Upload()
        {
            try
            {
                var file = Request.Form.Files[0];
                var folderName = "uploads"; // Path.Combine("Resources", "Docs");
                var tempDi = @"D:\VS\";
                var pathToSave = Path.Combine(tempDi, folderName);
                if (file.Length > 0)
                {
                    var fileName = ContentDispositionHeaderValue.Parse(file.ContentDisposition).FileName.Trim('"');
                    var fullPath = Path.Combine(pathToSave, fileName);
                    var dbPath = Path.Combine(folderName, fileName);
                    using (var stream = new FileStream(fullPath, FileMode.Create))
                    {
                        file.CopyTo(stream);
                    }

                    string localFilePath = fullPath;

                    string htmlText = string.Empty;
                    DocProcessor processor = new DocProcessor();
                    try
                    {
                        htmlText = processor.ParseDOCX();
                    }
                    catch (OpenXmlPackageException e)
                    {
                        if (e.ToString().Contains("Invalid Hyperlink"))
                        {
                            htmlText = processor.ParseDOCX();
                        }
                    }

                   return Ok(new { htmlText });
                }
                else
                {
                    return BadRequest();
                }
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Internal server error: {ex}");
            }
        }
        
        [HttpPost]
        [Route("makedoc")]
        public HttpResponseMessage makeDoc()
        {
            HttpResponseMessage responseMsg;
            try
            {
                var file = Request.Form.Files[0];
                var folderName = "uploads"; // Path.Combine("Resources", "Docs");
                var tempDi = @"D:\VS\";
                var pathToSave = Path.Combine(tempDi, folderName);
                if (file.Length > 0)
                {
                    var fileName = ContentDispositionHeaderValue.Parse(file.ContentDisposition).FileName.Trim('"');
                    var fullPath = Path.Combine(pathToSave, fileName);
                    var dbPath = Path.Combine(folderName, fileName);
                    using (var stream = new FileStream(fullPath, FileMode.Create))
                    {
                        file.CopyTo(stream);
                    }

                    string localFilePath = fullPath;

                    string htmlText = string.Empty;
                    DocProcessor processor = new DocProcessor();
                    try
                    {
                        htmlText = processor.ParseDOCX();
                    }
                    catch (OpenXmlPackageException e)
                    {
                        if (e.ToString().Contains("Invalid Hyperlink"))
                        {
                            htmlText = processor.ParseDOCX();
                        }
                    }

                    byte[] docBytes = System.IO.File.ReadAllBytes(fullPath);
                    var dataStream = new MemoryStream(docBytes);
                    responseMsg = new HttpResponseMessage(HttpStatusCode.OK);
                    responseMsg.Content = new StreamContent(dataStream);
                    responseMsg.Content.Headers.ContentDisposition = new System.Net.Http.Headers.ContentDispositionHeaderValue("attachment");
                    responseMsg.Content.Headers.ContentDisposition.FileName = file.FileName;
                    responseMsg.Content.Headers.ContentType = new MediaTypeHeaderValue("application/octet-stream");
                    //
                }
                else
                {
                     responseMsg =  new HttpResponseMessage(HttpStatusCode.BadRequest);
                }
            }
            catch (Exception ex)
            {
                responseMsg = new HttpResponseMessage(HttpStatusCode.InternalServerError);
            }

            return responseMsg;
        }

    }
}
