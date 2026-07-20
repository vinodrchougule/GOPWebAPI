using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using Aspose.Cells;

namespace GOPWebAPI.Helpers
{
    public class AsposeHelpers
    {

        #region Function to return workbook
        public Workbook GetWorkbook()
        {
            Workbook wb = new Workbook();
            Aspose.Cells.License l = new Aspose.Cells.License();
            l.SetLicense(HttpContext.Current.Server.MapPath("~/License.lic"));
            wb.LoadData(HttpContext.Current.Server.MapPath(@"~/Templates/Griha.xlsx"));
            return wb;
        }
        #endregion

        #region Function to return style
        public Aspose.Cells.Style GetStyle(Workbook wb, int wsNo, string styleType)
        {
            var ws = wb.Worksheets[wsNo];

            Aspose.Cells.Style style = ws.Cells[0, 0].GetStyle();
            if (styleType.ToLower() == "header")
            {
                style.IsTextWrapped = true;
                style.HorizontalAlignment = TextAlignmentType.Center;
                style.VerticalAlignment = TextAlignmentType.Center;
                style.ForegroundColor = System.Drawing.Color.FromArgb(0, 0, 240, 255);
                style.BackgroundColor = System.Drawing.Color.FromArgb(0, 0, 240, 255);
                style.Pattern = BackgroundType.VerticalStripe;
                style.Font.Color = System.Drawing.Color.Black;
                style.Font.IsBold = true;
            }
            else if (styleType.ToLower() == "center")
                style.HorizontalAlignment = TextAlignmentType.Center;
            else if (styleType.ToLower() == "left")
                style.HorizontalAlignment = TextAlignmentType.Left;

            style.Borders[BorderType.TopBorder].Color = System.Drawing.Color.Black;
            style.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
            style.Borders[BorderType.BottomBorder].Color = System.Drawing.Color.Black;
            style.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
            style.Borders[BorderType.LeftBorder].Color = System.Drawing.Color.Black;
            style.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
            style.Borders[BorderType.RightBorder].Color = System.Drawing.Color.Black;
            style.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;

            return style;
        }
        #endregion
    }
}