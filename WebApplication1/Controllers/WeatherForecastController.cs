using Microsoft.AspNetCore.Mvc;
using System.Net;
using System.Text;

namespace WebApplication1.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class WeatherForecastController : ControllerBase
    {
        private static readonly string[] Summaries = new[]
        {
        "Freezing", "Bracing", "Chilly", "Cool", "Mild", "Warm", "Balmy", "Hot", "Sweltering", "Scorching"
    };

        private readonly ILogger<WeatherForecastController> _logger;

        public WeatherForecastController(ILogger<WeatherForecastController> logger)
        {
            _logger = logger;
        }

        [HttpGet(Name = "GetWeatherForecast")]
        public IEnumerable<WeatherForecast> Get()
        {
            return Enumerable.Range(1, 5).Select(index => new WeatherForecast
            {
                Date = DateTime.Now.AddDays(index),
                TemperatureC = Random.Shared.Next(-20, 55),
                Summary = Summaries[Random.Shared.Next(Summaries.Length)]
            })
            .ToArray();
        }

        [HttpGet("Test")]
        public IActionResult Test()
        {

            var html = @$"MIME-Version: 1.0
Content-Type: multipart/related; boundary=""----=_NextPart_ZROIIZO.ZCZYUACXV.ZARTUI""


------=_NextPart_ZROIIZO.ZCZYUACXV.ZARTUI
Content-Location: file:///C:/mydocument.htm
Content-Type: text/html; charset=""utf-8""


{"<html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'>\r\n<head><title>Microsoft Office HTML Example</title>\r\n<link rel=File-List href=\"mydocument_files/filelist.xml\">\r\n<style><!-- \r\n@page\r\n{\r\n    size:21cm 29.7cmt;  /* A4 */\r\n    margin:1cm 1cm 1cm 1cm; /* Margins: 2.5 cm on each side */\r\n    mso-page-orientation: portrait;  \r\n    mso-header: url(\"mydocument_files/headerfooter.htm\") h1;\r\n    mso-footer: url(\"mydocument_files/headerfooter.htm\") f1;\t\r\n}\r\n@page Section1 { }\r\ndiv.Section1 { page:Section1; }\r\np.MsoHeader, p.MsoFooter { border: 1px solid black; }\r\n--></style>\r\n</head>\r\n<body>\r\n<div class=Section1>\r\nI'm page 1.\r\n<br clear=all style='mso-special-character:line-break;page-break-before:always'>\r\nI'm page 2.\r\n</div>\r\n</body>\r\n</html>"}


------=_NextPart_ZROIIZO.ZCZYUACXV.ZARTUI
Content-Location: file:///C:/mydocument_files/headerfooter.htm
Content-Type: text/html; charset=""utf-8""


<html xmlns:v=""urn:schemas-microsoft-com:vml"" xmlns:o=""urn:schemas-microsoft-com:office:office"" xmlns:w=""urn:schemas-microsoft-com:office:word"" xmlns:m=""http://schemas.microsoft.com/office/2004/12/omml""= xmlns=""http://www.w3.org/TR/REC-html40"">
<body>
 
<div style=""mso-element:header; border:none"" id=""h1"">
<p style=""margin:0;padding:0"">Header</p>
<p style=""margin:0;padding:0"">Header</p>
</div>
 
<div style='mso-element:footer' id=f1>
<p class=MsoFooter><span class=SpellE>Footer</span> page <!--[if supportFields]><span
class=MsoPageNumber><span style='mso-element:field-begin'></span><span
style='mso-spacerun:yes'> </span>PAGE <span style='mso-element:field-separator'></span></span><![endif]--><span
class=MsoPageNumber><span style='mso-no-proof:yes'>1</span></span><!--[if supportFields]><span
class=MsoPageNumber><span style='mso-element:field-end'></span></span><![endif]--><span
class=MsoPageNumber>/</span><!--[if supportFields]><span class=MsoPageNumber><span
style='mso-element:field-begin'></span> NUMPAGES <span style='mso-element:field-separator'></span></span><![endif]--><span
class=MsoPageNumber><span style='mso-no-proof:yes'>1</span></span><!--[if supportFields]><span
class=MsoPageNumber><span style='mso-element:field-end'></span></span><![endif]-->
</p>
</div>
 </body>
</html>


------=_NextPart_ZROIIZO.ZCZYUACXV.ZARTUI--";

            Response.Headers.Add("Content-disposition", "attachment; filename=myfile.doc");
            var _byte=System.Text.Encoding.UTF8.GetBytes(html);
            return File(_byte, "application/msword");
        }

        private  string Base64Encode(string content)
        {
            byte[] bytes = Encoding.UTF8.GetBytes(content);
            return Convert.ToBase64String(bytes);
        }
    }
}