using NPOI.SS.UserModel;

namespace GetJsonAndValidateDataFromXlsx
{
    public static class ExcelExtensions
    {
        public static string StringValue(this ICell cell)
        {
            if (cell == null)
                return "";
            switch ((int)cell.CellType)
            {
                case 0:
                    return cell.NumericCellValue.ToString().Trim();
                case 1:
                    return cell.StringCellValue.Trim();
                case 2:
                    return cell.CellFormula.ToString().Trim();
                case 3:
                    return "";
                case 4:
                    return cell.BooleanCellValue.ToString().Trim();
                default:
                    return cell.StringCellValue.Trim();
            }
        }
    }
}
