﻿using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Web.Http;

namespace ReleaseNotesWeb.Controllers
{
    [RoutePrefix("api")]
    public class ReleaseNotesController : ApiController
    {
        [Route("ReleaseNotes")]
        [HttpPost]
        public HttpResponseMessage ReleaseNotes([FromBody] JObject fields)//string data)
        {
            // JObject fields = JObject.Parse(data);
            if (fields == null) return Request.CreateResponse(HttpStatusCode.BadRequest);

            ReleaseNotesLibrary.Utility.NamedLookup settings = 
                new ReleaseNotesLibrary.Utility.NamedLookup("Settings");

            string generatorType = fields.GetValue("generator").ToString();
            settings["Team Project Path"] = fields.GetValue("teamProjectPath").ToString();
            settings["Project Name"] = fields.GetValue("projectName").ToString();
            settings["Project Subpath"] = fields.GetValue("projectSubpath").ToString();
            settings["Iteration"] = fields.GetValue("iteration").ToString();
            settings["Database"] = fields.GetValue("database").ToString();
            settings["Database Server"] = fields.GetValue("databaseServer").ToString();
            settings["Web Server"] = fields.GetValue("webServer").ToString();
            settings["Doc Type"] = "APPLICATION BUILD/RELEASE NOTES\n";
            settings["Web Location"] = fields.GetValue("webLocation").ToString();

            // only support server side Excel
            ReleaseNotesLibrary.Generators.ReleaseNotesGenerator g;
            if (generatorType == "excel")
            {
                g = ReleaseNotesLibrary.Generators.ExcelServerGenerator.ExcelServerGeneratorFactory(settings, true);
            }
            else if (generatorType == "html")
            {
                g = ReleaseNotesLibrary.Generators.HTMLGenerator.HTMLGeneratorFactory(settings, true);
            }
            else
            {
                generatorType = "html";
                g = ReleaseNotesLibrary.Generators.HTMLGenerator.HTMLGeneratorFactory(settings, true);
            }

            byte[] result = g.GenerateReleaseNotes();

            HttpResponseMessage response = new HttpResponseMessage(HttpStatusCode.OK);
            response.Content = new StreamContent(new MemoryStream(result));
            string outputFileName = "";

            if (generatorType == "excel")
            {
                response.Content.Headers.ContentType = new MediaTypeHeaderValue("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
                outputFileName = settings["Project Name"] + " " + settings["Iteration"] + " Release Notes.xlsx";
            }
            else if (generatorType == "html")
            {
                response.Content.Headers.ContentType = new MediaTypeHeaderValue("text/html");
                outputFileName = settings["Project Name"] + " " + settings["Iteration"] + " Release Notes.html";
            }
            else
            {
                response.Content.Headers.ContentType = new MediaTypeHeaderValue("application/octet-stream");
            }

            response.Content.Headers.ContentDisposition = new System.Net.Http.Headers.ContentDispositionHeaderValue("attachment")
            {
                FileName = outputFileName,
                Name = "Release Notes"
            };

            return response;
        }
    }
}
