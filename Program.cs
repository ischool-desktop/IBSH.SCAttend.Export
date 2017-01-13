using FISCA.Permission;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace IBSH.SCAttend.Export
{
    public static class Program
    {
        [FISCA.MainMethod]
        public static void Main()
        {
            {
                var btn = K12.Presentation.NLDPanels.Student.RibbonBarItems["資料統計"]["匯出"]["成績相關匯出"]["匯出學生修課及評量成績"];
                btn.Enable = false;
                RoleAclSource.Instance["學生"]["功能按鈕"].Add(new RibbonFeature("CD098CF9-8D79-477C-AC81-EE846F41566E", "匯出學生修課及評量成績"));
                if (UserAcl.Current["CD098CF9-8D79-477C-AC81-EE846F41566E"].Executable)
                {
                    K12.Presentation.NLDPanels.Student.SelectedSourceChanged += delegate
                    {
                        btn.Enable = K12.Presentation.NLDPanels.Student.SelectedSource.Count > 0;
                    };
                }
                btn.Click += delegate
                {
                    SmartSchool.API.PlugIn.Export.Exporter exporter = new StudentSCAttendExport();
                    ExportStudentV2 wizard = new ExportStudentV2(exporter.Text, exporter.Image);
                    exporter.InitializeExport(wizard);
                    wizard.ShowDialog();
                };
            }
        }
    }
}
