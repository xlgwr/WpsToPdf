using ConsoleApp1;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;

namespace WpsToPdf.WebApi.Controllers
{
    public class ValuesController : ApiController
    {
        [HttpGet]
        [Route("api/getPdf")]
        public string GetPdf(string fileSource, string wpsFilename)
        {
           
            ToPdfHelper toPdfHelper = new ToPdfHelper(fileSource);
            var saveFile = toPdfHelper.SavePdf(wpsFilename);
            return saveFile;
             
        }
        // GET api/values
        public IEnumerable<string> Get()
        {
            return new string[] { "value1", "value2" };
        }

        // GET api/values/5
        public string Get(int id)
        {
            return "value";
        }

        // POST api/values
        public void Post([FromBody]string value)
        {
        }

        // PUT api/values/5
        public void Put(int id, [FromBody]string value)
        {
        }

        // DELETE api/values/5
        public void Delete(int id)
        {
        }
    }
}
