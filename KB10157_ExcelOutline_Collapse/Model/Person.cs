using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KB10157_ExcelOutline_Collapse.Model;

internal class Person
{
    public int ID { get; set; }
    public String FamilyName { get; set; } = String.Empty;
    public String GivenName { get; set; } = String.Empty;
    public String Prefecture { get; set; } = String.Empty;
    public String City { get; set; } = String.Empty;

    public Person()
    {
    }
}
