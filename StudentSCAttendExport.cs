using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SmartSchool.API.PlugIn.Export;

namespace IBSH.SCAttend.Export
{
    class StudentSCAttendExport : SmartSchool.API.PlugIn.Export.Exporter
    {
        public StudentSCAttendExport()
        {
            this.Text = "匯出學生修課及評量成績";
        }

        public override void InitializeExport(ExportWizard wizard)
        {
            Dictionary<string, K12.Data.CourseRecord> dicCourse = null;
            Dictionary<string, K12.Data.ExamRecord> dicExam = new Dictionary<string, K12.Data.ExamRecord>();
            #region 匯出欄位
            wizard.ExportableFields.Add("學年度");
            wizard.ExportableFields.Add("學期");
            wizard.ExportableFields.Add("課程名稱");
            wizard.ExportableFields.Add("科目");
            wizard.ExportableFields.Add("節數");
            wizard.ExportableFields.Add("學分數");
            wizard.ExportableFields.Add("總成績");

            foreach (var examRec in K12.Data.Exam.SelectAll())
            {
                wizard.ExportableFields.Add(examRec.Name);
                dicExam.Add(examRec.ID, examRec);
            }
            #endregion

            #region 模式切換
            var mode1 = "currentSemester";
            var mode2 = "allSemester";
            var currentMode = mode1;

            var radMode1 = new SmartSchool.API.PlugIn.VirtualRadioButton() { Text = "匯出目前學期資料" };
            var radMode2 = new SmartSchool.API.PlugIn.VirtualRadioButton() { Text = "匯出所有學期資料" };

            wizard.Options.Add(radMode1);
            wizard.Options.Add(radMode2);

            wizard.ControlPanelOpen += delegate
            {
                radMode1.Checked = currentMode == mode1;
                radMode2.Checked = currentMode == mode2;
            };
            wizard.ControlPanelClose += delegate
            {
                if (radMode1.Checked)
                    currentMode = mode1;
                else
                    currentMode = mode2;
                dicCourse = null;
            };
            #endregion

            wizard.ExportPackage += delegate (object sender, ExportPackageEventArgs e)
            {
                lock (this)
                {
                    if (dicCourse == null)
                    {
                        dicCourse = new Dictionary<string, K12.Data.CourseRecord>();
                        if (currentMode == mode1)
                        {
                            foreach (var courseRec in K12.Data.Course.SelectBySchoolYearAndSemester(int.Parse(K12.Data.School.DefaultSchoolYear), int.Parse(K12.Data.School.DefaultSemester)))
                            {
                                dicCourse.Add(courseRec.ID, courseRec);
                            }
                        }
                        else
                        {
                            foreach (var courseRec in K12.Data.Course.SelectAll())
                            {
                                dicCourse.Add(courseRec.ID, courseRec);
                            }
                        }
                    }
                }

                var scAttendList = K12.Data.SCAttend.SelectByStudentIDAndCourseID(e.List, dicCourse.Keys);
                scAttendList.Sort((K12.Data.SCAttendRecord r1, K12.Data.SCAttendRecord r2) =>
                {
                    if (r1.RefStudentID != r2.RefStudentID)
                        return e.List.IndexOf(r1.RefStudentID).CompareTo(e.List.IndexOf(r2.RefStudentID));

                    if (dicCourse[r1.RefCourseID].SchoolYear != null
                        && dicCourse[r2.RefCourseID].SchoolYear != null
                        && dicCourse[r1.RefCourseID].SchoolYear.Value != dicCourse[r2.RefCourseID].SchoolYear.Value)
                        return dicCourse[r2.RefCourseID].SchoolYear.Value.CompareTo(dicCourse[r1.RefCourseID].SchoolYear.Value);

                    if (dicCourse[r1.RefCourseID].Semester != null
                        && dicCourse[r2.RefCourseID].Semester != null
                        && dicCourse[r1.RefCourseID].Semester.Value != dicCourse[r2.RefCourseID].Semester.Value)
                        return dicCourse[r2.RefCourseID].Semester.Value.CompareTo(dicCourse[r1.RefCourseID].Semester.Value);

                    if (dicCourse[r1.RefCourseID].Credit != null
                        && dicCourse[r2.RefCourseID].Credit != null
                        && dicCourse[r1.RefCourseID].Credit.Value != dicCourse[r2.RefCourseID].Credit.Value)
                        return dicCourse[r2.RefCourseID].Credit.Value.CompareTo(dicCourse[r1.RefCourseID].Credit.Value);

                    if (dicCourse[r1.RefCourseID].Subject != dicCourse[r2.RefCourseID].Subject)
                        return dicCourse[r1.RefCourseID].Subject.CompareTo(dicCourse[r2.RefCourseID].Subject);

                    return r1.ID.CompareTo(r2.ID);
                });
                var dicSCETake = new Dictionary<string, List<K12.Data.SCETakeRecord>>();
                foreach (var sceTakeRec in K12.Data.SCETake.SelectByStudentAndCourse(e.List, dicCourse.Keys))
                {
                    if (!dicSCETake.ContainsKey(sceTakeRec.RefSCAttendID))
                        dicSCETake.Add(sceTakeRec.RefSCAttendID, new List<K12.Data.SCETakeRecord>());
                    dicSCETake[sceTakeRec.RefSCAttendID].Add(sceTakeRec);
                }

                foreach (var scAttendRec in scAttendList)
                {
                    var row = new SmartSchool.API.PlugIn.RowData();
                    row.ID = scAttendRec.RefStudentID;
                    //wizard.ExportableFields.Add("學年度");
                    //wizard.ExportableFields.Add("學期");
                    //wizard.ExportableFields.Add("課程名稱");
                    //wizard.ExportableFields.Add("科目");
                    //wizard.ExportableFields.Add("節數");
                    //wizard.ExportableFields.Add("學分數");

                    //foreach (var examRec in K12.Data.Exam.SelectAll())
                    //{
                    //    wizard.ExportableFields.Add(examRec.Name);
                    //}
                    if (wizard.SelectedFields.Contains("學年度"))
                        row["學年度"] = "" + dicCourse[scAttendRec.RefCourseID].SchoolYear;
                    if (wizard.SelectedFields.Contains("學期"))
                        row["學期"] = "" + dicCourse[scAttendRec.RefCourseID].Semester;
                    if (wizard.SelectedFields.Contains("課程名稱"))
                        row["課程名稱"] = "" + dicCourse[scAttendRec.RefCourseID].Name;
                    if (wizard.SelectedFields.Contains("科目"))
                        row["科目"] = "" + dicCourse[scAttendRec.RefCourseID].Subject;
                    if (wizard.SelectedFields.Contains("節數"))
                        row["節數"] = "" + dicCourse[scAttendRec.RefCourseID].Period;
                    if (wizard.SelectedFields.Contains("學分數"))
                        row["學分數"] = "" + dicCourse[scAttendRec.RefCourseID].Credit;
                    if (wizard.SelectedFields.Contains("總成績"))
                        row["總成績"] = "" + scAttendRec.Score;
                    if (dicSCETake.ContainsKey(scAttendRec.ID))
                    {
                        foreach (var sceTakeRec in dicSCETake[scAttendRec.ID])
                        {
                            if (sceTakeRec.RefStudentID == scAttendRec.RefStudentID)
                            {
                                if (wizard.SelectedFields.Contains(dicExam[sceTakeRec.RefExamID].Name))
                                {
                                    var node = sceTakeRec.ToXML().SelectSingleNode("Extension/Extension/Score");
                                    if (node != null)
                                    {
                                        row[dicExam[sceTakeRec.RefExamID].Name] = node.InnerText;
                                    }
                                }
                            }
                        }
                    }


                    e.Items.Add(row);
                }
            };
        }
    }
}
