using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using ExcelValidator;

namespace ExcelBulkInsert.Controllers
{
    public class UserController : ApiController
    {
        [HttpGet]
        [Route("api/users/bulkinsert")]
        public string BulkInsert(string fileName)
        {
            ReadExcel excel = new ReadExcel(fileName);
            excel.ReadFile();

            return "Ok";

        }
    }
}
