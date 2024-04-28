using System.Collections.Concurrent;
using System.Drawing;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace me.fengyj.CommonLib.Office.Excel {
    public class SheetStyle {
        public static IEnumerable<BorderStyle> GetBorderStyles() => BorderStyle.GetStyles();
        public static IEnumerable<CellStyle> GetCellStyles() => CellStyle.GetStyles();
        public static IEnumerable<FillStyle> GetFillStyles() => FillStyle.GetStyles();
        public static IEnumerable<FontStyle> GetFontStyles() => FontStyle.GetStyles();
        public static IEnumerable<NumberingStyle> GetNumberingStyles() => NumberingStyle.GetStyles();
    }

    public class AlignmentStyle {

        public static readonly AlignmentStyle Left = new() { Horizontal = HorizontalAlignmentValues.Left, Vertical = VerticalAlignmentValues.Center };
        public static readonly AlignmentStyle Center = new() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center };
        public static readonly AlignmentStyle Right = new() { Horizontal = HorizontalAlignmentValues.Right, Vertical = VerticalAlignmentValues.Center };

        public HorizontalAlignmentValues Horizontal { get; private set; }
        public VerticalAlignmentValues Vertical { get; private set; }

        public Alignment GetAlignment() => new() { Horizontal = this.Horizontal, Vertical = this.Vertical };
    }

    public class BorderStyle {

        private static volatile int Seq = -1;
        private static readonly ConcurrentBag<BorderStyle> styles = [];

        public static readonly BorderStyle None = new(new() {
            LeftBorder = new(),
            TopBorder = new(),
            RightBorder = new(),
            BottomBorder = new()
        });  // this should be the first one as the default border style
        public static readonly BorderStyle All = new(new() {
            LeftBorder = new() { Style = BorderStyleValues.Thin },
            TopBorder = new() { Style = BorderStyleValues.Thin },
            RightBorder = new() { Style = BorderStyleValues.Thin },
            BottomBorder = new() { Style = BorderStyleValues.Thin }
        });
        public static BorderStyle Bottom = new(new() { BottomBorder = new() { Style = BorderStyleValues.Thin } });

        public BorderStyle(Border border) {

            this.Border = border;
            styles.Add(this);
            TooManyStylesException.Check(styles);
        }

        public uint StyleId { get; private set; } = (uint)Interlocked.Increment(ref Seq);
        public Border Border { get; private set; }

        public static IEnumerable<BorderStyle> GetStyles() => styles.OrderBy(s => s.StyleId);
    }

    public class CellStyle {

        private static volatile int Seq = -1;
        private static readonly ConcurrentBag<CellStyle> styles = [];
        private static Dictionary<string, CellStyle>? namedStyles;

        #region default styles

        public static readonly CellStyle None = new(); // this should be the first one as the default cell style
        public static readonly CellStyle H1 = new(fontStyle: FontStyle.H1);
        public static readonly CellStyle H2 = new(fontStyle: FontStyle.H2);
        public static readonly CellStyle H3 = new(fontStyle: FontStyle.H3);
        public static readonly CellStyle Normal = new(fontStyle: FontStyle.Normal);
        public static readonly CellStyle Bold = new(fontStyle: FontStyle.Bold);
        public static readonly CellStyle Italic = new(fontStyle: FontStyle.Italic);
        public static readonly CellStyle Emphasize = new(fontStyle: FontStyle.Emphasize);
        public static readonly CellStyle Warning = new(fontStyle: FontStyle.Warning);
        public static readonly CellStyle Error = new(fontStyle: FontStyle.Error);
        public static readonly CellStyle Quote = new(fontStyle: FontStyle.Quote);

        public static readonly CellStyle TableHeader = new(fontStyle: FontStyle.TableHeader);
        public static readonly CellStyle TableFooter = new(fontStyle: FontStyle.Mono_Normal, alignmentStyle: AlignmentStyle.Right);

        public static readonly CellStyle Hyperlink = new(fontStyle: FontStyle.Hyperlink);

        public static readonly CellStyle Integer_Default = new(
            fontStyle: FontStyle.Mono_Normal,
            numberingStyle: DefaultStyleConfig.Numbering.DefaultInteger,
            alignmentStyle: AlignmentStyle.Right,
            cellValueType: CellValues.Number);
        public static readonly CellStyle Decimal_Default = new(
            fontStyle: FontStyle.Mono_Normal,
            numberingStyle: DefaultStyleConfig.Numbering.DefaultDecimal,
            alignmentStyle: AlignmentStyle.Right,
            cellValueType: CellValues.Number);

        public static readonly CellStyle DateTime_Default = new(
            fontStyle: FontStyle.Mono_Normal,
            numberingStyle: DefaultStyleConfig.Numbering.DefaultDateTime,
            alignmentStyle: AlignmentStyle.Center,
            cellValueType: CellValues.Date);
        public static readonly CellStyle Date_Default = new(
            fontStyle: FontStyle.Mono_Normal,
            numberingStyle: DefaultStyleConfig.Numbering.DefaultDate,
            alignmentStyle: AlignmentStyle.Center,
            cellValueType: CellValues.Date);
        public static readonly CellStyle Time_Default = new(
            fontStyle: FontStyle.Mono_Normal,
            numberingStyle: DefaultStyleConfig.Numbering.DefaultTime,
            alignmentStyle: AlignmentStyle.Center,
            cellValueType: CellValues.Date);

        public static readonly CellStyle Timespan_Default = new(
            fontStyle: FontStyle.Mono_Normal,
            alignmentStyle: AlignmentStyle.Right); // the value should be converted to string with the format: "[-][d:]hh:mm:ss"

        public static readonly CellStyle Bool_Default = new(
            fontStyle: FontStyle.Normal,
            alignmentStyle: AlignmentStyle.Center,
            cellValueType: CellValues.Boolean);

        #endregion

        public CellStyle(
            NumberingStyle? numberingStyle = null,
            FontStyle? fontStyle = null,
            FillStyle? fillStyle = null,
            BorderStyle? borderStyle = null,
            AlignmentStyle? alignmentStyle = null,
            CellValues? cellValueType = null) {

            this.CellFormat = new() {
                NumberFormatId = numberingStyle?.Format.NumberFormatId,
                FontId = fontStyle?.StyleId,
                FillId = fillStyle?.StyleId,
                BorderId = borderStyle?.StyleId,
                ApplyNumberFormat = numberingStyle == null ? null : BooleanValue.FromBoolean(true),
                Alignment = alignmentStyle?.GetAlignment()
            };
            this.CellValueType = cellValueType ?? CellValues.InlineString;
            this.Format = numberingStyle?.Format?.FormatCode?.Value;
            styles.Add(this);
            TooManyStylesException.Check(styles);
        }

        private CellStyle() {

            this.CellFormat = new();
            styles.Add(this);
        }
        public uint StyleId { get; private set; } = (uint)Interlocked.Increment(ref Seq);
        public CellFormat CellFormat { get; private set; }
        public CellValues CellValueType { get; private set; }
        public string? Format { get; private set; }

        public CellStyle With(
            NumberingStyle? numberingStyle = null,
            FontStyle? fontStyle = null,
            FillStyle? fillStyle = null,
            BorderStyle? borderStyle = null,
            AlignmentStyle? alignmentStyle = null,
            CellValues? cellValueType = null) {

            var nfId = numberingStyle?.Format.NumberFormatId?.Value ?? this.CellFormat.NumberFormatId?.Value;
            var fmt = numberingStyle?.Format.FormatCode?.Value ?? this.Format;
            var ftId = fontStyle?.StyleId ?? this.CellFormat.FontId?.Value;
            var flId = fillStyle?.StyleId ?? this.CellFormat.FillId?.Value;
            var bdId = borderStyle?.StyleId ?? this.CellFormat.BorderId?.Value;
            var align = alignmentStyle?.GetAlignment() ?? this.CellFormat.Alignment;
            var cvt = cellValueType ?? this.CellValueType;

            var style = styles.Where(s => nfId == s.CellFormat?.NumberFormatId?.Value)
                .Where(s => ftId == s.CellFormat?.FontId?.Value)
                .Where(s => flId == s.CellFormat?.FillId?.Value)
                .Where(s => bdId == s.CellFormat?.BorderId?.Value)
                .Where(s => align?.Horizontal?.Value == s.CellFormat?.Alignment?.Horizontal?.Value)
                .Where(s => align?.Vertical?.Value == s.CellFormat?.Alignment?.Vertical?.Value)
                .Where(s => cvt == s.CellValueType)
                .FirstOrDefault();

            if (style != null) return style;

            style = new CellStyle(cellValueType: cvt);
            if (nfId != null) style.CellFormat.NumberFormatId = UInt32Value.FromUInt32(nfId.Value);
            if (ftId != null) style.CellFormat.FontId = UInt32Value.FromUInt32(ftId.Value);
            if (flId != null) style.CellFormat.FillId = UInt32Value.FromUInt32(flId.Value);
            if (bdId != null) style.CellFormat.BorderId = UInt32Value.FromUInt32(bdId.Value);
            if (align != null) style.CellFormat.Alignment = new Alignment { Horizontal = align.Horizontal, Vertical = align.Vertical };
            if (this.CellFormat.ApplyAlignment != null) style.CellFormat.ApplyAlignment = new BooleanValue(this.CellFormat.ApplyAlignment.Value);
            if (fmt != null) style.Format = fmt;
            return style;
        }

        public static IEnumerable<CellStyle> GetStyles() => styles.OrderBy(s => s.StyleId);

        public static CellStyle? GetNamedStyle(string name) {

            if (namedStyles == null) {
                lock (typeof(CellStyle)) {
                    if (namedStyles == null) {

                        var dict = new Dictionary<string, CellStyle>();
                        var fieldds = typeof(CellStyle).GetFields(System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Static)
                            .Where(f => f.FieldType == typeof(CellStyle));
                        foreach (var f in fieldds) {
                            var o = f.GetValue(null);
                            if (o is CellStyle cs) dict.Add(f.Name, cs);
                        }
                        namedStyles = dict;
                    }
                }
            }
            if (namedStyles.TryGetValue(name, out var s)) return s;
            else return null;
        }
    }

    public class FillStyle {

        private static volatile int Seq = -1;
        private static readonly ConcurrentBag<FillStyle> styles = [];

        public static readonly FillStyle None = new(pattern: PatternValues.None); // this should be the first one as the default fill style

        public FillStyle(System.Drawing.Color? fgColor = null, System.Drawing.Color? bgColor = null, PatternValues? pattern = null) {

            if (pattern == null)
                pattern = fgColor.HasValue || bgColor.HasValue ? PatternValues.Solid : PatternValues.None;

            this.Fill = new() {
                PatternFill = new() {
                    PatternType = pattern,
                    ForegroundColor = fgColor.HasValue ? new ForegroundColor { Rgb = GetColor(fgColor.Value) } : null,
                    BackgroundColor = bgColor.HasValue ? new BackgroundColor { Rgb = GetColor(bgColor.Value) } : null
                }
            };
            styles.Add(this);
            TooManyStylesException.Check(styles);
        }

        public uint StyleId { get; private set; } = (uint)Interlocked.Increment(ref Seq);
        public Fill Fill { get; private set; }

        public static IEnumerable<FillStyle> GetStyles() => styles.OrderBy(s => s.StyleId).ToList();

        private static HexBinaryValue GetColor(System.Drawing.Color color) {
            return new HexBinaryValue { Value = ColorTranslator.ToHtml(color).Replace("#", "") };
        }
    }

    public class FontStyle {

        private static volatile int Seq = -1;
        private static readonly ConcurrentBag<FontStyle> styles = [];

        public static readonly FontStyle Normal = new(DefaultStyleConfig.Font.Default_Name, DefaultStyleConfig.Font.Default_Size, color: DefaultStyleConfig.Font.Color_Default); // this should be the first one as the default font style
        public static readonly FontStyle H1 = new(DefaultStyleConfig.Font.H1_Name, DefaultStyleConfig.Font.H1_Size, bold: true);
        public static readonly FontStyle H2 = new(DefaultStyleConfig.Font.H2_Name, DefaultStyleConfig.Font.H2_Size, bold: true);
        public static readonly FontStyle H3 = new(DefaultStyleConfig.Font.H3_Name, DefaultStyleConfig.Font.H3_Size, bold: true, color: DefaultStyleConfig.Font.Color_Default);
        public static readonly FontStyle Bold = new(DefaultStyleConfig.Font.Default_Name, DefaultStyleConfig.Font.Default_Size, bold: true, color: DefaultStyleConfig.Font.Color_Default);
        public static readonly FontStyle Italic = new(DefaultStyleConfig.Font.Default_Name, DefaultStyleConfig.Font.Default_Size, italic: true, color: DefaultStyleConfig.Font.Color_Default);
        public static readonly FontStyle Quote = new(DefaultStyleConfig.Font.Default_Name, DefaultStyleConfig.Font.Quote_Size, italic: true, color: DefaultStyleConfig.Font.Color_Quote);
        public static readonly FontStyle Mono_Normal = new(DefaultStyleConfig.Font.Mono_Name, DefaultStyleConfig.Font.Default_Size, color: DefaultStyleConfig.Font.Color_Default);
        public static readonly FontStyle Mono_Bold = new(DefaultStyleConfig.Font.Mono_Name, DefaultStyleConfig.Font.Default_Size, bold: true, color: DefaultStyleConfig.Font.Color_Default);
        public static readonly FontStyle Mono_Italic = new(DefaultStyleConfig.Font.Mono_Name, DefaultStyleConfig.Font.Default_Size, italic: true, color: DefaultStyleConfig.Font.Color_Default);

        public static readonly FontStyle Emphasize = new(DefaultStyleConfig.Font.Default_Name, DefaultStyleConfig.Font.Default_Size, color: DefaultStyleConfig.Font.Color_Emphasize);
        public static readonly FontStyle Warning = new(DefaultStyleConfig.Font.Default_Name, DefaultStyleConfig.Font.Default_Size, color: DefaultStyleConfig.Font.Color_Warning);
        public static readonly FontStyle Error = new(DefaultStyleConfig.Font.Default_Name, DefaultStyleConfig.Font.Default_Size, color: DefaultStyleConfig.Font.Color_Error);

        public static readonly FontStyle Mono_Emphasize = new(DefaultStyleConfig.Font.Mono_Name, DefaultStyleConfig.Font.Default_Size, color: DefaultStyleConfig.Font.Color_Emphasize);
        public static readonly FontStyle Mono_Warning = new(DefaultStyleConfig.Font.Mono_Name, DefaultStyleConfig.Font.Default_Size, color: DefaultStyleConfig.Font.Color_Warning);
        public static readonly FontStyle Mono_Error = new(DefaultStyleConfig.Font.Mono_Name, DefaultStyleConfig.Font.Default_Size, color: DefaultStyleConfig.Font.Color_Error);

        public static readonly FontStyle TableHeader = new(DefaultStyleConfig.Font.Default_Name, DefaultStyleConfig.Font.Default_Size, bold: true, italic: true, color: DefaultStyleConfig.Font.Color_Default);

        public static readonly FontStyle Hyperlink = new(DefaultStyleConfig.Font.Default_Name, DefaultStyleConfig.Font.Default_Size, underline: true, color: DefaultStyleConfig.Font.Color_Hyperlink);

        public FontStyle(Font font) {

            this.Font = font;
            styles.Add(this);
            TooManyStylesException.Check(styles);
        }

        public FontStyle(string fontName, double fontSize, bool bold = false, bool italic = false, bool underline = false, System.Drawing.Color? color = null)
            : this(new Font() {
                FontName = new() { Val = StringValue.FromString(fontName) },
                FontSize = new() { Val = DoubleValue.FromDouble(fontSize) },
                Color = color.HasValue ? new() { Rgb = GetColor(color.Value) } : null,
                Bold = bold ? new Bold() : null,
                Italic = italic ? new Italic() : null,
                Underline = underline ? new Underline() : null
            }) {
        }

        public uint StyleId { get; private set; } = (uint)Interlocked.Increment(ref Seq);
        public Font Font { get; private set; }

        public static IEnumerable<FontStyle> GetStyles() => styles.OrderBy(s => s.StyleId).ToList();

        private static HexBinaryValue GetColor(System.Drawing.Color color) {
            return new HexBinaryValue {
                Value = ColorTranslator.ToHtml(System.Drawing.Color.FromArgb(
                          color.A,
                          color.R,
                          color.G,
                          color.B)).Replace("#", "")
            };
        }
    }

    public class NumberingStyle {

        private const uint StartId = 164; // the values less than it are used by Excel built-in styles

        private static volatile uint Seq = StartId - 1;
        private static readonly ConcurrentDictionary<string, NumberingStyle> styles = [];

        #region default formats

        public static readonly NumberingStyle None = new();

        public static readonly NumberingStyle DateTime_Default = new("yyyy-mm-dd hh:mm:ss");
        public static readonly NumberingStyle DateTime_UK = new("dd/mm/yyyy hh:mm:ss AM/PM");
        public static readonly NumberingStyle DateTime_US = new("mm/dd/yyyy hh:mm:ss AM/PM");

        public static readonly NumberingStyle Date_Default = new("yyyy-mm-dd");
        public static readonly NumberingStyle Date_UK = new("dd/mm/yyyy");
        public static readonly NumberingStyle Date_US = new("mm/dd/yyyy");

        public static readonly NumberingStyle Time = new("hh:mm:ss");

        public static readonly NumberingStyle Integer_Default = new("0");
        public static readonly NumberingStyle Integer_Thousands = new("#,##0");

        public static readonly NumberingStyle Decimal_Default = new("0.0########");
        public static readonly NumberingStyle Decimal_2 = new("0.00");
        public static readonly NumberingStyle Decimal_3 = new("0.000");
        public static readonly NumberingStyle Decimal_4 = new("0.0000");
        public static readonly NumberingStyle Decimal_6 = new("0.000000");

        public static readonly NumberingStyle Decimal_Thousands = new("#,##0.0########");
        public static readonly NumberingStyle Decimal_Thousands_2 = new("#,##0.00");
        public static readonly NumberingStyle Decimal_Thousands_3 = new("#,##0.000");
        public static readonly NumberingStyle Decimal_Thousands_4 = new("#,##0.0000");
        public static readonly NumberingStyle Decimal_Thousands_6 = new("#,##0.000000");

        public static readonly NumberingStyle Percent_Default = new("0.0########%");
        public static readonly NumberingStyle Percent_0 = new("0%");
        public static readonly NumberingStyle Percent_1 = new("0.0%");
        public static readonly NumberingStyle Percent_2 = new("0.00%");
        public static readonly NumberingStyle Percent_3 = new("0.000%");

        #endregion

#pragma warning disable CS8618 // Non-nullable field must contain a non-null value when exiting constructor. Consider declaring as nullable.
        public NumberingStyle(string format) {

            if (string.IsNullOrWhiteSpace(format)) throw new ArgumentNullException(nameof(format), "The parameter cannot be null nor empty.");

            var existed = styles.GetOrAdd(format, f => {
                TooManyStylesException.Check(styles);
                this.Format = new NumberingFormat { FormatCode = StringValue.FromString(f), NumberFormatId = Interlocked.Increment(ref Seq) };
                return this;
            });

            if (existed != null) {
                this.Format = existed.Format;
            }
        }
#pragma warning restore CS8618 // Non-nullable field must contain a non-null value when exiting constructor. Consider declaring as nullable.

        private NumberingStyle() {
            this.Format = new NumberingFormat() { NumberFormatId = 0 };
            //styles.TryAdd(string.Empty, this); // DO NOT add it to the list.
        }

#pragma warning disable CS8604 // Possible null reference argument.
        public uint StyleId => this.Format.NumberFormatId;
#pragma warning restore CS8604 // Possible null reference argument.
        public NumberingFormat Format { get; private set; }

        public static IEnumerable<NumberingStyle> GetStyles() => styles.Values.OrderBy(s => s.StyleId).ToList();
    }

    public class TableStyle {

        public TableStyle(string? styleName = null, bool showHeader = true, bool showRowStripes = true, bool showColumnStripes = false) {
            this.TableStyleInfo = new() { Name = styleName ?? DefaultStyleConfig.Table.TableStyle_Default, ShowColumnStripes = showColumnStripes, ShowRowStripes = showRowStripes };
            this.ShowHeader = showHeader;
        }
        public TableStyleInfo TableStyleInfo { get; private set; }
        public bool ShowHeader { get; private set; }
    }

    /// <summary>
    /// Exception for remindering if adding too many styles by accident.
    /// If it's necessary, can set <see cref="TooManyStylesException.DisableAlert"/> to disable the alert.
    /// </summary>
    public class TooManyStylesException : ApplicationException {
        public TooManyStylesException() :
            base("Too many styles have been added, please confirm this is necessary.") { }

        /// <summary>
        /// 
        /// </summary>
        public static bool DisableAlert { get; set; } = false;

        public static void Check(System.Collections.ICollection styles) {

            if (!DisableAlert && styles != null && styles.Count > 256)
                throw new TooManyStylesException();
        }
    }
}
