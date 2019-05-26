using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using aspnet_modelbinder.Controllers.Web.ModelBinder;

namespace aspnet_modelbinder.Controllers
{
    [Route("api/[controller]/[action]")]
    [ApiController]
    public class ValuesController : ControllerBase
    {
        // GET api/values
        [HttpPost]
        public IActionResult CreateCustomer( [ModelBinder(typeof(JsonWithFilesFormDataModelBinder))]Customer customer)
        {
            byte[] passport = new byte[customer.Passport.Length]; 
            var stream = customer.Passport.OpenReadStream();
            stream.Read(passport);
            return Ok();
        }
    }
}
