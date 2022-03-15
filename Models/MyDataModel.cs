using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CSharpDocxLab.Models;

public class MyDataModel
{
  public string Name { get; set; }
  public decimal Salary { get; set; }
  public DateTime Birthday { get; set; }
  public string Remark { get; set; }
  public List<MyDetail> ItemList { get; set; }
}

public class MyDetail
{
  public int ItemSn { get; set; }
  public string ItemName { get; set; }
  public decimal? ItemCount { get; set; }
}
