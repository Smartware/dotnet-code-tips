using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using aspnet_modelbinder.Controllers.Web.ModelBinder;
using aspnet_modelbinder.Utility.Excel;

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
            var excelMgr = new ExcelManager(stream);

            var list = excelMgr.ReadSheetAtToList<Customer>(0, true,
            c => new Customer
            {
                LastName = c.LastName,
                FirstName = c.FirstName
            });

            return Ok();
        }
    }
}
