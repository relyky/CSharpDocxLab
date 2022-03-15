using CSharpDocxLab.Services;
using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CSharpDocxLab
{
  internal class App
  {
    readonly IConfiguration _config;
    readonly DocxApplyService _wordSvc;

    public App(IConfiguration config, DocxApplyService wordSvc)
    {
      _config = config;
      _wordSvc = wordSvc;
    }

    /// <summary>
    /// 取代原本 Program.Main() 函式的效用。
    /// </summary>
    public void Run(string[] args)
    {
      _wordSvc.ApplyHelloDocx();
      Console.WriteLine("產生 Hello.docx 完成。");

#if DEBUG
      System.Diagnostics.Debugger.Break();
#else
      Console.WriteLine("Press any key to continue.");
      Console.ReadKey();
#endif
    }
  }
}