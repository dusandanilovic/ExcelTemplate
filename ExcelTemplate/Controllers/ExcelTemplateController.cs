using ExcelTemplate.Data;
using Microsoft.AspNetCore.Mvc;
using System.Configuration;

namespace ExcelTemplate.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class ExcelTemplateController : ControllerBase
    {
        private IConfiguration Configuration;

        public ExcelTemplateController(IConfiguration configuration)
        {
            Configuration = configuration;
        }

        [HttpGet("GetValidation")]
        public async Task<ActionResult> GetTemplate(string reportDate, string reportType)
        {
            var connectionString = Configuration.GetConnectionString("destination");
            var dataContext = new DataContext(connectionString);

            var result = dataContext.ExecuteProcedure("CR_SA_xls", reportDate, reportType);
            return Ok();
        }
    }
}