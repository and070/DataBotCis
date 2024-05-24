using System.Collections.Generic;

namespace DataBotV5.Data.Projects.Freelance

{
    interface IFreelanceBase
    {
        List<string> Emails { get; set; }
        string Copias { get; set; }
    }
}
