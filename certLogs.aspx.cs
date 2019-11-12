using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace ecomLogs
{
  public partial class _default : System.Web.UI.Page
  {
    public int svMembNo;

    protected void Page_Load(object sender, EventArgs e)
    {
      // this is the membNo of the administrator launching this service from the V5 menu
      string temp = Request["svMembNo"];
      if (temp == null) temp = "1252951"; // this is a random level 5 learner used only if nothing comes through
      svMembNo = Int32.Parse(temp);

    }
  }
}