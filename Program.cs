using OfficeOpenXml;
using System.Reflection;
using System.Text.RegularExpressions;

namespace DesafioTunts;

class Program
{
    static void Main(string[] args)
    {
        var workingDirectory = System.IO.Path.GetDirectoryName(Assembly.GetEntryAssembly().Location);
        var path = Path.Combine(Directory.GetParent(workingDirectory).Parent.Parent.FullName, @"engenharia-de-software-desafio-ananias-jose.xlsx");
        ManipulateXLS(path);
    }
    static void ManipulateXLS(string FilePath)
    {
        const int columnSituation = 7;
        const int columnFinalGradeForApproval = 8;
       
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        FileInfo existingFile = new FileInfo(FilePath);
        
        if (!existingFile.Exists) 
        {
            Console.WriteLine("File not found");
            return;
        }

        using (ExcelPackage package = new ExcelPackage(existingFile))
        {
            ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
            
            if (worksheet.Dimension?.End?.Row == null)
            {
                Console.WriteLine("Empty sheet");
                return;
            }

            int rowCount = worksheet.Dimension.End.Row;     
            
            var totalClassesSemester = Convert.ToString(worksheet.Cells[2, 1].Value);
            var totalClasses = Convert.ToInt32(Regex.Replace(totalClassesSemester, @"\D", ""));
            int maxAbsence = totalClasses / 4;

            for (int row = 4; row <= rowCount; row++)
            {
                int totalAbsencePerStudent = Convert.ToInt32(worksheet.Cells[row, 3].Value);
               
                double exam1 = Convert.ToDouble(worksheet.Cells[row, 4].Value);
                double exam2 = Convert.ToDouble(worksheet.Cells[row, 5].Value);
                double exam3 = Convert.ToDouble(worksheet.Cells[row, 6].Value);

                double averageGrade = Math.Round(((exam1 + exam2 + exam3) / 3) / 10, 1);

                if (totalAbsencePerStudent > maxAbsence)
                {
                    worksheet.Cells[row, columnSituation].Value = "Failed by absence";
                    worksheet.Cells[row, columnFinalGradeForApproval].Value = 0;
                    continue;
                }

                if (averageGrade < 5)
                {
                    worksheet.Cells[row, columnSituation].Value ="Failed the examination";
                    worksheet.Cells[row, columnFinalGradeForApproval].Value = 0;
                }

                if (averageGrade >= 5 && averageGrade < 7)
                {
                    double finalExam = (10 - averageGrade) * 2;
                    
                    worksheet.Cells[row, columnSituation].Value = "Final exam";
                    worksheet.Cells[row, columnFinalGradeForApproval].Value = finalExam;
                }

                if (averageGrade >= 7)
                {
                    worksheet.Cells[row, columnSituation].Value = "Approved";
                    worksheet.Cells[row, columnFinalGradeForApproval].Value = 0;
                }

            }
            package.Save();
        }

    }

}