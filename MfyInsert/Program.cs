
using OfficeOpenXml;

string path = @"C:\Users\User\Desktop\mfy.xlsx";
string scrptPath = @"C:\Users\User\Desktop\IHMA\ihma_adm_backend\src\libs\Ihma.Adm.DataLayer.PgSql\Scripts\0094 ADM_MNL insert into INFO_MFY.sql";
FileInfo fileInfo = new FileInfo(path);
ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
List<string> query = new List<string>();

using (ExcelPackage package = new ExcelPackage(fileInfo))
{
    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
    int rowCount = worksheet.Dimension.Rows;
    int columnCount = worksheet.Dimension.Columns;

    for (int row = 2; row <= rowCount; row++)
    {
        string code = worksheet.Cells[row, 8].Value?.ToString();
        string name = worksheet.Cells[row, 10].Value?.ToString();
        string sector = worksheet.Cells[row, 4].Value?.ToString();
        string region = worksheet.Cells[row, 3].Value?.ToString();
        string district = worksheet.Cells[row, 2].Value?.ToString();
        string inn = string.IsNullOrEmpty(worksheet.Cells[row, 7].Value?.ToString()) ? "null" : $"'{worksheet.Cells[row, 7].Value?.ToString()}'";
        var value = @$"INSERT INTO ADM_MNL.INFO_MFY (ID, ORDER_CODE, CODE, SHORT_NAME, FULL_NAME, INN, REGION_ID, DISTRICT_ID, SECTOR_ID, STATE_ID) 
        VALUES({row - 1}, '00{row - 1}', '00{row - 1}', N'{name}', N'{name}', {inn},{region},{district},{sector},1);";

        query.Add(value);
    }
}
File.WriteAllLines(scrptPath, query.ToArray());
Console.WriteLine();

