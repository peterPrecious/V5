using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace vuBuild20
{
    public partial class conv : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
//            foreach (var f in new DirectoryInfo(@"D:\vusystem\Webs\Active\V5\Code").GetFiles("*.asp", SearchOption.AllDirectories))

            foreach (var f in new DirectoryInfo(@"D:\OneDrive - VUBIZ\Webs\V5").GetFiles("*.asp"))
            {
		var text = File.ReadAllText(f.FullName, Encoding.Default);
	        File.WriteAllText(f.FullName, text, Encoding.UTF8);                	
            }
            foreach (var f in new DirectoryInfo(@"D:\OneDrive - VUBIZ\Webs\V5\inc").GetFiles("*.asp"))
            {
		var text = File.ReadAllText(f.FullName, Encoding.Default);
	        File.WriteAllText(f.FullName, text, Encoding.UTF8);                	
            }

            foreach (var f in new DirectoryInfo(@"D:\OneDrive - VUBIZ\Webs\V5\source").GetFiles("*.asp"))
            {
		var text = File.ReadAllText(f.FullName, Encoding.Default);
	        File.WriteAllText(f.FullName, text, Encoding.UTF8);                	
            }


        }
    }
}