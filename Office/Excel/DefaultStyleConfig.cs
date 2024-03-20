using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using DocumentFormat.OpenXml.Spreadsheet;

namespace me.fengyj.CommonLib.Office.Excel {

    public class DefaultStyleConfig {

        public static FontConfig Font { get; set; } = FontConfig.Default;

        public static TableConfig Table = new();

        public class FontConfig {

            public static readonly FontConfig Default = new();
            public static readonly FontConfig Chinese = new() { H1_Name = "SimHei", H2_Name = "SimHei", H3_Name = "Microsoft YaHei UI", Default_Name = "Microsoft YaHei UI" };

            public string H1_Name { get; set; } = "Arial";
            public string H2_Name { get; set; } = "Arial";
            public string H3_Name { get; set; } = "Calibri";
            public string Default_Name { get; set; } = "Calibri";
            public string Mono_Name { get; set; } = "Consolas";

            public double H1_Size { get; set; } = 16;
            public double H2_Size { get; set; } = 14;
            public double H3_Size { get; set; } = 12;
            public double Default_Size { get; set; } = 11;
            public double Quote_Size { get; set; } = 10;

            public System.Drawing.Color Color_Default { get; set; } = System.Drawing.Color.FromArgb(57, 57, 57); // dark gray
            public System.Drawing.Color Color_Emphasize { get; set; } = System.Drawing.Color.FromArgb(47, 117, 181); // blue
            public System.Drawing.Color Color_Warning { get; set; } = System.Drawing.Color.Orange;
            public System.Drawing.Color Color_Error { get; set; } = System.Drawing.Color.Red;
            public System.Drawing.Color Color_Quote { get; set; } = System.Drawing.Color.Gray;
        }

        public class TableConfig {

            public string TableStyle_Default = "TableStyleMedium1";
            public string PivotTableStyle_Default = "PivotStyleMedium1";
        }
    }

}
