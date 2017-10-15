using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace YPApps.Entities
{
    public interface IExport
    {
        Dictionary<string, object> ToDictionary();
    }
}
