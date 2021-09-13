using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Microsoft.Identity.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace DataverseTeamsIntegrationSample.Controllers
{
    public class TokenRequest
    {
        public string Token { get; set; }
    }

    [Route("api/1.0/auth")]
    [ApiController]
    public class AuthController : ControllerBase
    {
        private readonly IConfiguration _configuration;

        public AuthController(IConfiguration configuration)
        {
            _configuration = configuration;
        }

        [HttpPost]
        public async Task<IActionResult> GetOnBehalfOf([FromBody] TokenRequest request)
        {
            var builder = ConfidentialClientApplicationBuilder.Create(_configuration["Msal:ClientId"])
                .WithClientSecret(_configuration["Msal:ClientSecret"]);
                //.WithRedirectUri(_configuration["Msal:RedirectUri"]);

            var client = builder.Build();

            var tokenBuilder = client.AcquireTokenOnBehalfOf(new[] { "https://itorgdev.crm11.dynamics.com/user_impersonation" }, new UserAssertion(request.Token));

            var result = await tokenBuilder.ExecuteAsync();

            return Ok(result.AccessToken);
        }
    }
}
