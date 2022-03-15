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
  public void ApplyHelloDocx()
  {
    var model = new MyDataModel
    {
      Name = "甄美麗",
      Salary = 987654321.1234m,
      Birthday = DateTime.Parse("2008-12-25 01:02:03"),
      Remark = @"成立於２００２年７月，致力於客製化的軟體專案開發、企業流程自動化（事務、業務、生產及管理流程）及系統整合等服務給我們的客戶；並為個人金融服務業務（信用卡徵審流程、代償／信貸管理系統、風險管理系統、話務中心流程、信用卡帳務管理流程、郵購系統、費用代繳系統、行銷決策分析系統、外匯系統…）、證券交易服務（基金事務系統、基金會計系統、證券交易平台…）、知識入口網站（入口網站平台、搜尋引擎、工作流程、文件管理）及資訊部門業務流程（需求單、作業單、問題反應單…）等資訊科技整體解決方案之提供廠商。藉由提供企業優質的專案規劃、建置、維護與系統整合等服務，以協助客戶成為高效能之電子化企業為公司願景。了解更多…",
      ItemList = new List<MyDetail> {
                new MyDetail { ItemSn = 11, ItemName = "我是項目11", ItemCount = 54321m },
                new MyDetail { ItemSn = 12, ItemName = "我是項目12", ItemCount = null },
                new MyDetail { ItemSn = 13, ItemName = "我是項目13", ItemCount = null },
                new MyDetail { ItemSn = 14, ItemName = "我是項目14", ItemCount = 4444m },
                new MyDetail { ItemSn = 15, ItemName = "我是項目15", ItemCount = 55555m },
            }
    };

    string docTemplates = @"C:\GitHubRepos\CSharpDocxLab\DocTemplates";
    FileInfo docxTplPath = new FileInfo(Path.Combine(docTemplates, "Hello.cs.docx"));
    var document = DocumentFactory.Create(docxTplPath.FullName, model, forceCompile: true);
    document.Generate("Hello.docx");
  }
}
