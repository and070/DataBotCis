using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataBotV5.App.Global.Interfaces
{
    public interface IDataTable
    {
        DataTable DataTableValue { get; set; }
        bool ValidTable { get; set; }
    }
}
