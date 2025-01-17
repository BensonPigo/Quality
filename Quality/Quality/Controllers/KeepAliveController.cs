using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Http;

namespace Quality.Controllers
{
    public class KeepAliveController : ApiController
    {
        // GET: KeepAlive
        [HttpGet]
        [Route("api/keepalive")]
        public IHttpActionResult KeepAlive()
        {
            // 確保 Session 保持活躍
            return Ok("Session is alive.");
        }
    }
}