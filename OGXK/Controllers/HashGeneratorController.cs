using System.Security.Cryptography;
using System.Text;
using System.Web.Http;

namespace OGXK.Controllers
{
    //gavdcodebegin 01
    public class HashGeneratorController : ApiController
    {
        /// <summary>  
        /// Generate Hash 
        /// Call: http://localhost:51867/api/generatehash?RawData=abcdef
        /// </summary>  
        /// <returns the hash code></returns> 
        [HttpGet]  
        public string Get([FromUri] string RawData)
        {
            // Legacy code. This solutions uses the "ASP.NET Web Application (.Net Framework)"
            // See repo UBXU for the same result using the DOT.NET Core framework
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
