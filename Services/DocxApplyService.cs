using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using CSharpDocxLab.Models;
using SharpDocx;

namespace CSharpDocxLab.Services;

/// <summary>
/// DocX 文件範本套用
/// </summary>
internal class DocxApplyService
{
  public MemoryStream ApplyHelloDocx(MyDataModel model)
  {
    try
    {
      //※注意：帶入 ShaprDocx 的 data model 的命名空間必需是『xxx.Models』。  

      string docTemplates = @"C:\GitHubRepos\CSharpDocxLab\DocTemplates";
      FileInfo docxTplPath = new FileInfo(Path.Combine(docTemplates, "Hello.cs.docx"));
      var document = DocumentFactory.Create(docxTplPath.FullName, model, forceCompile: true);
      return document.Generate();
    }
    catch(Exception ex)
    {
      throw new ApplicationException("生成 Hello.docx 文件失敗！", ex);
    }
  }
}
