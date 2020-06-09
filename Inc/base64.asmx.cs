using System.Web.Services;

namespace Base64
{
  [WebService(Namespace = "http://tempuri.org/")]
  [WebServiceBinding(ConformsTo = WsiProfiles.BasicProfile1_1)]
  [System.ComponentModel.ToolboxItem(false)]
  [System.Web.Script.Services.ScriptService]

  public class Base64 : System.Web.Services.WebService
  {
    // this is used with V5 Certificates to properly create a Base64Encoded set of values
    // the resulting string (values) is sent to Alex's Certificate Service
    // https://stackoverflow.com/questions/11743160/how-do-i-encode-and-decode-a-base64-string

    [WebMethod]
    public string base64Encode(string plainText)
    {
      //plainText = "vFirstName=Péter&vLastName=Bulloch";
      if (plainText == null) { return null; }
      var plainTextBytes = System.Text.Encoding.UTF8.GetBytes(plainText);
      return System.Convert.ToBase64String(plainTextBytes);
    }

    [WebMethod]
    public string base64Decode(string base64EncodedData)
    {
      //base64EncodedData = "dkZpcnN0TmFtZT1Qw6l0ZXImdkxhc3ROYW1lPUJ1bGxvY2g=";
      if (base64EncodedData == null) { return null; }
      var base64EncodedBytes = System.Convert.FromBase64String(base64EncodedData);
      return System.Text.Encoding.UTF8.GetString(base64EncodedBytes);
    }

  }
}
