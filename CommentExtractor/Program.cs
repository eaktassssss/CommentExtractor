
using OfficeOpenXml;
using System.Text.RegularExpressions;


try
{
    string projectPath = @"ProjetPath";  
    var files = Directory.GetFiles(projectPath, "*.cs", SearchOption.AllDirectories);

    using (var package = new ExcelPackage())
    {
        var worksheet = package.Workbook.Worksheets.Add("Yorumlar");

        int row = 1;
        foreach (var file in files)
        {
            string[] lines = File.ReadAllLines(file);
            string currentClass = "Empty"; 

            foreach (var line in lines)
            {
                
                var classMatch = Regex.Match(line, @"class\s+(\w+)");
                if (classMatch.Success)
                {
                    currentClass = classMatch.Groups[1].Value;
                }

                
                if (line.Trim().StartsWith("//"))
                {
                    worksheet.Cells[row, 1].Value = file;
                    worksheet.Cells[row, 2].Value = currentClass;  
                    worksheet.Cells[row, 3].Value = line;
                    row++;
                }
            }

            string fileContent = File.ReadAllText(file);
            var matches = Regex.Matches(fileContent, @"/\*.*?\*/", RegexOptions.Singleline);
            foreach (Match match in matches)
            {
                worksheet.Cells[row, 1].Value = file;
                worksheet.Cells[row, 2].Value = currentClass;  
                worksheet.Cells[row, 3].Value = match.Value;
                row++;
            }
        }

        package.SaveAs(new FileInfo(@"C:\Comments.xlsx"));   
    }

    Console.WriteLine("İşlem tamamlandı!");
    Console.ReadLine();

}
catch (Exception exception)
{
    throw new Exception(exception.Message);
}
