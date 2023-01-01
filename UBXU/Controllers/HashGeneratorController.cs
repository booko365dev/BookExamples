using Microsoft.AspNetCore.Mvc;
using System.Security.Cryptography;
using System.Text;

namespace UBXU.Controllers
{
    //gavdcodebegin 01
    /// <summary>  
    /// Generate Hash Class
    /// Contains methods to generate a Hash
    /// </summary>  
    [Route("api/[controller]")]
    [ApiController]
    public class HashGeneratorController : Controller
    {
        //public IActionResult Index()
        //{
        //    return View();
        //}

        /// <summary>  
        /// Generate Hash 
        /// Call: https://localhost:[port]/api/generatehash?RawData=abcdef
        /// Returns: the hash code 
        /// </summary>  
        [HttpGet(Name = "GetHash")]
        public string Get([FromQuery] string RawData)
        {
            StringBuilder dataBuilder = new StringBuilder();

            using (SHA256 sha256Hash = SHA256.Create())
            {
                byte[] dataBytes = sha256Hash.ComputeHash(Encoding.UTF8.GetBytes(RawData));

                for (int i = 0; i < dataBytes.Length; i++)
                {
                    dataBuilder.Append(dataBytes[i].ToString("x2"));
                }
            }

            return dataBuilder.ToString();
        }
    }
    //gavdcodeend 01
}
