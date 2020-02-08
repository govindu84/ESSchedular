using System;
using System.Net.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using System.Net;

namespace ESScheduler.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ESAPPController : ControllerBase
    {
        public IConfiguration _IConfiguration { get; set; }    
        public string FilePath { get; set; }
        public string KibanaURI { get; set; }

        public ESAPPController(IConfiguration configuration)
        {
            this._IConfiguration = configuration;
            FilePath = _IConfiguration["PathName"];
            KibanaURI = _IConfiguration["KibanaURI"];
        }


        /// <summary>
        /// Call ES Service
        /// </summary>
        /// <returns></returns>
        [HttpGet]
        [Route("[action]")]
        public IActionResult GetCallESSService()
        {
            try
            {
              ESScheduler.CombineInputData(FilePath, KibanaURI);            
            }
            catch(Exception ex)
            {
                var response = new HttpResponseMessage(HttpStatusCode.NotFound)
                {
                    Content = new StringContent("Failed due to " + ex.Message + " Excel File path: " + FilePath, System.Text.Encoding.UTF8, "text/plain"),
                    StatusCode = HttpStatusCode.NotFound
                };
                throw new System.Web.Http.HttpResponseException(response);
            }
            return Ok();
        }

       /// <summary>
       /// Call ESS Service
       /// </summary>
       /// <param name="filePath">Excel File Path</param>
       /// <returns></returns>
        [HttpGet("{filePath}")]
        [Route("[action]")]
        public ActionResult GetCallESSService(string filePath)
        {
            try
            {
                ESScheduler.CombineInputData(filePath, KibanaURI);

            }
            catch (Exception ex)
            {
                var response = new HttpResponseMessage(HttpStatusCode.NotFound)
                {
                    Content = new StringContent("Failed due to " + ex.Message + " Excel File path: " + filePath, System.Text.Encoding.UTF8, "text/plain"),
                    StatusCode = HttpStatusCode.NotFound
                };
                throw new System.Web.Http.HttpResponseException(response);
            }
            return Ok();
        }

        // POST api/values
        [HttpPost]
        public void Post([FromBody] string value)
        {
        }

        // PUT api/values/5
        [HttpPut("{id}")]
        public void Put(int id, [FromBody] string value)
        {
        }

        // DELETE api/values/5
        [HttpDelete("{id}")]
        public void Delete(int id)
        {
        }
    }
}
