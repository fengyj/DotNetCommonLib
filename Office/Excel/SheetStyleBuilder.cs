using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace me.fengyj.CommonLib.Office.Excel {
    public class SheetStyleBuilder {
        public static void BuildTo(WorkbookPart workbookPart) {

            var style = new Stylesheet();

            var formats = new NumberingFormats();
            formats.Append(SheetStyle.GetNumberingStyles().Select(i => i.Format.CloneNode(true))); // the items cannot be added to different spreadsheet, so need to clone them
            style.Append(formats);

            var fonts = new Fonts();
            fonts.Append(SheetStyle.GetFontStyles().Select(i => i.Font.CloneNode(true))); // the items cannot be added to different spreadsheet, so need to clone them
            style.Append(fonts);

            var fills = new Fills();
            fills.Append(SheetStyle.GetFillStyles().Select(i => i.Fill.CloneNode(true))); // the items cannot be added to different spreadsheet, so need to clone them
            style.Append(fills);

            var borders = new Borders();
            borders.Append(SheetStyle.GetBorderStyles().Select(i => i.Border.CloneNode(true))); // the items cannot be added to different spreadsheet, so need to clone them
            style.Append(borders);

            var cells = new CellFormats();
            cells.Append(SheetStyle.GetCellStyles().Select(i => i.CellFormat.CloneNode(true))); // the items cannot be added to different spreadsheet, so need to clone them
            style.Append(cells);

            var stylePart = workbookPart.AddNewPart<WorkbookStylesPart>();
            stylePart.Stylesheet = style;
            stylePart.Stylesheet.Save();
        }
    }
}
