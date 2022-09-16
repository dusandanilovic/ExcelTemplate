using ExcelTemplate.Data;
using Microsoft.AspNetCore.Mvc;
using System.Configuration;
using StagingAlgorithm_newdata;
using MathWorks.MATLAB.NET.Arrays;

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

        [HttpGet("GetTemplate")]
        public async Task<ActionResult> GetTemplate(string reportDate, string reportType)
        {
            StagingAlgorithm s = new StagingAlgorithm();
            s.StagingAlgorithm_newdata();
            var connectionString = Configuration.GetConnectionString("destination");
            var dataContext = new DataContext(connectionString);
            dataContext.GenerateExcel();
            var result = dataContext.ExecuteProcedure("CR_SA_xls", reportDate, reportType);
            return Ok();
        }

        [HttpPost("PostStagingAlgorithm")]
        public ActionResult PostStagingAlgorithm()
        {
            string[] strings = new string[] { "{-db}", "{hercules}", "{-ip}", "{10.10.0.46}", "{-u}", "{root}", "{-p}", "{admin@2022}" };
            //var charArray = new MWCharArray(strings);
            //MWCellArray cellArray = new MWCellArray(charArray);
            StagingAlgorithm s = new StagingAlgorithm();
            MWStringArray stringArray = new MWStringArray(strings);
            s.StagingAlgorithm_newdata(stringArray);
            return Ok();
        }
    }
}