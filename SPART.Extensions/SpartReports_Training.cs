namespace SPART.Extensions
{
  using System;
  using System.Collections.Specialized;
  using System.Linq;
  using Kcsar.Database.Model;
  using OfficeOpenXml;

  public partial class SpartReports
  {
    private static void TrainingReport(SpartReports me, ExcelPackage package, NameValueCollection queries)
    {
      var id = new Guid("574cb2fe-1acc-4e04-919c-030546b0e7bd");
      var db = me.db.Value;

      DateTime today = DateTime.Today;

      var sheet = package.Workbook.Worksheets[1];

      var members = db.GetActiveMembers(id, today, "Memberships").OrderBy(f => f.LastName).ThenBy(f => f.FirstName);

      var spartCourses = new[] { "OEC", "Avalanche I", "Avalanche II", "MT&R", "MT&R 2" };
      var courses = db.TrainingCourses.Where(f => spartCourses.Contains(f.DisplayName) || f.WacRequired > 0);
      var spartCourseGuids = new Guid[spartCourses.Length];
      for (int i = 0; i < spartCourses.Length; i++)
      {
        string courseName = spartCourses[i];
        spartCourseGuids[i] = courses.Where(f => f.DisplayName == courseName).Select(f => f.Id).SingleOrDefault();
      }

      int row = 2;
      foreach (var member in members)
      {
        var expires = CompositeTrainingStatus.Compute(member, courses, today);

        int col = 1;
        sheet.Cells[row, col++].Value = member.LastName;
        sheet.Cells[row, col++].Value = member.FirstName;

        foreach (var courseId in spartCourseGuids)
        {
          if (courseId != Guid.Empty)
          {
            var expire = expires.Expirations[(Guid)courseId];
            if (expire.CourseName.StartsWith("Avalanche") && expire.Completed.HasValue)
            {
              expire.Expires = expire.Completed.Value.AddYears(3);
            }
            sheet.Cells[row, col].Value = expire.ToString();
          }
          col++;
        }

        var now = DateTime.Now;
        var yearStart = new DateTime(now.Year, 1, 1);
        var lastYear = new DateTime(now.Year - 1, 1, 1);

        sheet.Cells[row, col++].Value = member.MissionRosters.Where(f => f.Unit.Id == id && f.TimeIn >= yearStart).Select(f => f.Mission.Id).Distinct().Count();
        sheet.Cells[row, col++].Value = member.MissionRosters.Where(f => f.Unit.Id == id && f.TimeIn >= lastYear && f.TimeIn < yearStart).Select(f => f.Mission.Id).Distinct().Count();

        row++;
      }

      sheet.Cells["A:I"].AutoFitColumns();
    }
  }
}
