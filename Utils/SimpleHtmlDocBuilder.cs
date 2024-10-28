using System.Security;

namespace me.fengyj.CommonLib.Utils {
    public class SimpleHtmlDocBuilder {

        private readonly List<IContentBuilder> contents = [];

        private SimpleHtmlDocBuilder(string? inlineBodyStyle = null) {
            this.InlineBodyStyle = inlineBodyStyle;
        }
        public string? InlineBodyStyle { get; private set; }

        public static SimpleHtmlDocBuilder Create(string? inlineBodyStyle = null) {

            return new SimpleHtmlDocBuilder(inlineBodyStyle);
        }

        public static string BuildWithRawHtml(String html, string? inlineBodyStyle = null) {

            var style = ContentBuilderUtil.BuildInlineStyleAttribute(inlineBodyStyle);
            return $"<!DOCTYPE html><html><head><style>{DefaultHtmlStyle.StyleSheet}</style></head><body {style}>{html}</body></html>";
        }

        public SimpleHtmlDocBuilder AppendHeader(HeaderBuilder builder) {

            this.contents.Add(builder);
            return this;
        }

        public SimpleHtmlDocBuilder AppendHeader(String text, string? inlineTextStyle = null) {

            this.contents.Add(HeaderBuilder.Create(text, inlineTextStyle));
            return this;
        }

        public SimpleHtmlDocBuilder AppendHeaderInRed(String text) {

            this.contents.Add(HeaderBuilder.Create(text, DefaultHtmlStyle.InlineStyle_Color_Red));
            return this;
        }

        public SimpleHtmlDocBuilder AppendHeaderInBlue(String text) {

            this.contents.Add(HeaderBuilder.Create(text, DefaultHtmlStyle.InlineStyle_Color_Blue));
            return this;
        }

        public SimpleHtmlDocBuilder AppendParagraph(ParagraphBuilder builder) {

            this.contents.Add(builder);
            return this;
        }

        public SimpleHtmlDocBuilder AppendParagraph(String text, string? inlineTextStyle = null) {

            this.contents.Add(ParagraphBuilder.Create(text, inlineTextStyle));
            return this;
        }

        public SimpleHtmlDocBuilder AppendParagraphInRed(String text) {

            this.contents.Add(ParagraphBuilder.Create(text, DefaultHtmlStyle.InlineStyle_Color_Red));
            return this;
        }

        public SimpleHtmlDocBuilder AppendParagraphInBlue(String text) {

            this.contents.Add(ParagraphBuilder.Create(text, DefaultHtmlStyle.InlineStyle_Color_Blue));
            return this;
        }

        public SimpleHtmlDocBuilder AppendTable(TableBuilder builder) {

            this.contents.Add(builder);
            return this;
        }

        public SimpleHtmlDocBuilder AppendList(ListBuilder builder) {

            this.contents.Add(builder);
            return this;
        }

        public String Build() {

            return BuildWithRawHtml(string.Join(string.Empty, this.contents.Select(i => i.Build())), inlineBodyStyle: this.InlineBodyStyle);
        }
    }

    public static class DefaultHtmlStyle {

        public const string Class_Text = "text";
        public const string Class_Link = "link";
        public const string Class_Table = "table";
        public const string Class_TableHeader = "table_header";
        public const string Class_Header = "header";
        public const string Class_List = "list";

        public static string InlineStyle_Color_Red = "color:#ED7D31";
        public static string InlineStyle_Color_Blue = "color:#5B9BD5";
        public static string InlineStyle_Table_Row_Strips = "background-color:#EAEAEA";

        public static string InlineStyle_Bold = "font-weight: Bold";
        public static string InlineStyle_Italic = "font-style: italic";
        public static string InlineStyle_Underline = "text-decoration: underline";
        public static string InlineStyle_Line_Through = "text-decoration: line-through";


        public static string StyleSheet = @"
body {background:#fefefe;font-family:Verdana,sans-serif;font-size:11.0pt;color:#404040}

.text {font-family:""Verdana"",sans-serif;font-size:10.0pt;color:#404040}
.header {font-family:""Verdana"",sans-serif;font-size:12.0pt;color:#404040}

.table {border-collapse:collapse}
.table_header {background-color:#C0C0C0}
";

        public static string? GetStyle(params string[] styles) {
            return styles == null ? null : string.Join(";", styles);
        }
    }

    public class ListBuilder : IContentBuilder {

        private readonly IList<IList<ParagraphBuilder>> contents = [];

        private ListBuilder(string? inlineListStyle) {
            this.InlineListStyle = inlineListStyle;
        }

        public string? InlineListStyle { get; private set; }
        public string ListClass { get; set; } = DefaultHtmlStyle.Class_List;

        public static ListBuilder Create(string content) {
            return Create(Enumerable.Repeat(content, 1));
        }

        public static ListBuilder Create(IEnumerable<string> items, string? inlineListStyle = null) {

            var build = new ListBuilder(inlineListStyle);
            foreach (var item in items) {
                build.AppendItem(item);
            }
            return build;
        }

        public static ListBuilder Create(IEnumerable<IEnumerable<string>> paragraphList, string? inlineListStyle = null) {

            var build = new ListBuilder(inlineListStyle);
            foreach (var item in paragraphList) {
                build.AppendItem(item);
            }
            return build;
        }

        public static ListBuilder Create(IEnumerable<ParagraphBuilder> paragraphBuilders, string? inlineListStyle = null) {
            return new ListBuilder(inlineListStyle).AppendItem(paragraphBuilders);
        }

        public static ListBuilder Create(IEnumerable<IEnumerable<ParagraphBuilder>> paragraphBuilders, string? inlineListStyle = null) {

            var build = new ListBuilder(inlineListStyle);
            foreach (var item in paragraphBuilders) {
                build.AppendItem(item);
            }
            return build;
        }

        public ListBuilder AppendItem(string item, string? inlineTextStyle = null) {

            this.contents.Add([ParagraphBuilder.Create(item, inlineTextStyle)]);
            return this;
        }

        public ListBuilder AppendItem(IEnumerable<string> item, string? inlineTextStyle = null) {
            return this.AppendItem(item.Select(i => ParagraphBuilder.Create(i, inlineTextStyle)));
        }

        public ListBuilder AppendItem(ParagraphBuilder paragraphBuilder) {
            this.contents.Add([paragraphBuilder]);
            return this;
        }

        public ListBuilder AppendItem(IEnumerable<ParagraphBuilder> paragraphBuilders) {
            this.contents.Add(paragraphBuilders.ToList());
            return this;
        }

        public string Build() {

            var listClass = this.BuildClassAttribute(this.ListClass); // ListClass 
            var listInlineStyle = this.BuildInlineStyleAttribute(this.InlineListStyle);

            var items = this.contents.Select(i => string.Join(string.Empty, i.Select(p => p.Build()))).Select(s => $"<li>{s}</li>");
            return $"<ul{listClass}{listInlineStyle}>{string.Join("", items)}</ul>";
        }
    }

    public class ParagraphBuilder : IContentBuilder {

        private readonly IList<TextBuilder> contents;

        private ParagraphBuilder(IList<TextBuilder> contents) {
            this.contents = contents;
        }

        public static ParagraphBuilder Create(IList<TextBuilder> contents) {
            return new ParagraphBuilder(contents);
        }

        public static ParagraphBuilder Create(string content, string? inlineTextStyle = null) {
            return Create(TextBuilder.Create(content, inlineTextStyle));
        }

        public static ParagraphBuilder CreateInRed(string content) {
            return Create(content, DefaultHtmlStyle.InlineStyle_Color_Red);
        }

        public static ParagraphBuilder CreateInBlue(string content) {
            return Create(content, DefaultHtmlStyle.InlineStyle_Color_Blue);
        }

        public static ParagraphBuilder Create(TextBuilder textBuilder) {

            var list = new List<TextBuilder> {
                textBuilder
            };
            return new ParagraphBuilder(list);
        }

        public ParagraphBuilder Append(TextBuilder textBuilder) {

            this.contents.Add(textBuilder);
            return this;
        }

        public ParagraphBuilder Append(string content, string? inlineTextStyle = null) {
            return this.Append(TextBuilder.Create(content, inlineTextStyle));
        }

        public ParagraphBuilder AppendInRed(string content) {
            return this.Append(TextBuilder.Create(content, DefaultHtmlStyle.InlineStyle_Color_Red));
        }

        public ParagraphBuilder AppendInBlue(string content) {
            return this.Append(TextBuilder.Create(content, DefaultHtmlStyle.InlineStyle_Color_Blue));
        }

        public string Build() {

            return $"<p>{string.Join("", this.contents.Select(c => c.Build()))}</p>";
        }
    }

    public class HeaderBuilder : IContentBuilder {

        private readonly IList<TextBuilder> contents;

        private HeaderBuilder(IList<TextBuilder> contents, string? inlineHeaderStyle = null) {

            this.contents = contents;
            this.InlineHeaderStyle = inlineHeaderStyle;

            foreach (var c in contents) {
                if (c.InlineTextStyle == null) c.InlineTextStyle = inlineHeaderStyle;
                c.StyleClass = this.HeaderClass;
            }
        }

        public string? InlineHeaderStyle { get; private set; }
        public string HeaderClass { get; private set; } = DefaultHtmlStyle.Class_Header;

        public static HeaderBuilder Create(string header, string? inlineHeaderStyle = null) {
            return Create(TextBuilder.Create(header), inlineHeaderStyle);
        }

        public static HeaderBuilder Create(TextBuilder textBuilder, string? inlineHeaderStyle = null) {

            var list = new List<TextBuilder> {
                textBuilder
            };
            return new HeaderBuilder(list, inlineHeaderStyle);
        }

        public HeaderBuilder Append(TextBuilder textBuilder) {

            this.contents.Add(textBuilder);

            if (textBuilder.InlineTextStyle == null) textBuilder.InlineTextStyle = this.InlineHeaderStyle;
            textBuilder.StyleClass = this.HeaderClass;

            return this;
        }

        public HeaderBuilder Append(string content, string? inlineTextStyle = null) {
            return this.Append(TextBuilder.Create(content, inlineTextStyle));
        }

        public HeaderBuilder AppendInRed(string content) {
            return this.Append(TextBuilder.Create(content, DefaultHtmlStyle.InlineStyle_Color_Red));
        }

        public HeaderBuilder AppendInBlue(string content) {
            return this.Append(TextBuilder.Create(content, DefaultHtmlStyle.InlineStyle_Color_Blue));
        }

        public string Build() {

            return $"<p><b>{string.Join("", this.contents.Select(c => c.Build()))}</b></p>";
        }
    }

    public class TableBuilder : IContentBuilder {

        private List<IList<object>> items = [];

        private TableBuilder(string? inlineTableStyle = null, string? inlineTableHeaderStyle = null, string? inlineTableHeaderTextStyle = null) {
            this.InlineTableStyle = inlineTableStyle;
            this.InlineTableHeaderStyle = inlineTableHeaderStyle;
            this.InlineTableHeaderTextStyle = inlineTableHeaderTextStyle;
        }

        public string? InlineTableStyle { get; private set; }
        public string? InlineTableHeaderStyle { get; private set; }
        public string? InlineTableHeaderTextStyle { get; private set; }

        public string TableClass { get; set; } = DefaultHtmlStyle.Class_Table;
        public string TableHeaderClass { get; set; } = DefaultHtmlStyle.Class_TableHeader;

        public static TableBuilder Create(IList<string> headers, string? inlineTableStyle = null, string? inlineTableHeaderStyle = null, string? inlineTableHeaderTextStyle = null) {

            var builder = new TableBuilder(inlineTableStyle, inlineTableHeaderStyle, inlineTableHeaderTextStyle);
            builder.items.Add(headers.Cast<object>().ToList());
            return builder;
        }

        public static TableBuilder Create(string[] headers, string? inlineTableStyle = null, string? inlineTableHeaderStyle = null, string? inlineTableHeaderTextStyle = null) {

            var builder = new TableBuilder(inlineTableStyle, inlineTableHeaderStyle, inlineTableHeaderTextStyle);
            builder.items.Add(headers.Cast<object>().ToList());
            return builder;
        }

        public TableBuilder AddRow(IList<object> row) {
            this.items.Add(row);
            return this;
        }

        public TableBuilder AddRow(params object[] row) {
            this.items.Add(row?.ToList() ?? []);
            return this;
        }

        public TableBuilder AddRows(IEnumerable<IList<object>> rows) {
            this.items.AddRange(rows);
            return this;
        }

        public TableBuilder AddRow(IList<IList<object>> rows) {
            this.items.AddRange(rows);
            return this;
        }

        public string Build() {

            var rowNum = 0;

            var tblHeaderClass = this.BuildClassAttribute(this.TableHeaderClass); // TableHeaderClass
            var tblHeaderStyle = this.BuildInlineStyleAttribute(this.InlineTableHeaderStyle);
            var tblClass = this.BuildClassAttribute(this.TableClass); // TableClass
            var tblStyle = this.BuildInlineStyleAttribute(this.InlineTableStyle);

            var header = $"<tr{tblHeaderClass}{tblHeaderStyle}>{string.Join(string.Empty, this.items[0].Select(i => this.GetHeaderCell(i?.ToString())))}</tr>";

            var aligns = this.GetColumnsAlign();

            var rows = this.items.Skip(1).Select(i => {

                var rowStyle = rowNum++ % 2 == 0 ? string.Empty : DefaultHtmlStyle.InlineStyle_Table_Row_Strips;
                var cells = i.Select((c, idx) => {

                    if (idx >= aligns.Count) return null;

                    return this.GetCell(c, aligns[idx]);
                });
                return $"<tr {this.BuildInlineStyleAttribute(rowStyle)}>{string.Join("", cells)}</tr>";
            });

            return $"<table cellpadding='6' cellspacing='0'{tblClass}{tblStyle}>{header}{string.Join("", rows)}</table>";
        }

        private string GetHeaderCell(string? header) {

            var tb = TextBuilder.Create(header ?? string.Empty, this.InlineTableHeaderTextStyle);

            return $"<td align='center'><b><i>{tb.Build()}</i></b></td>";
        }

        private string GetCell(object? value, string align = "left") {

            var tb = value is IContentBuilder ? (IContentBuilder)value : TextBuilder.Create(value?.ToString() ?? string.Empty);

            return $"<td align='{align}'>{tb.Build()}</td>";
        }

        private IList<string> GetColumnsAlign() {

            var aligns = new List<string>();

            for (var i = 0; i < this.items[0].Count; i++) {

                var align = "left";
                var idx = i;

                var cellVal = this.items
                    .Skip(1)
                    .Select(r => r.Count > idx ? r[idx] : null)
                    .Where(r => r != null)
                    .FirstOrDefault();

                if (cellVal != null) {
                    if (cellVal is decimal
                        || cellVal is int
                        || cellVal is uint
                        || cellVal is short
                        || cellVal is ushort
                        || cellVal is long
                        || cellVal is ulong
                        || cellVal is float
                        || cellVal is double) {
                        align = "right";
                    }
                }
                aligns.Add(align);
            }

            return aligns;
        }
    }

    public class LinkBuilder : TextBuilder {

        private LinkBuilder(string link, string? label, string? inlineTextStyle = null) : base(label ?? link, inlineTextStyle) {
            this.Link = link;
            this.StyleClass = DefaultHtmlStyle.Class_Link;
        }

        public static LinkBuilder Create(string link, string? label = null, string? inlineTextStyle = null) {
            return new LinkBuilder(link, label, inlineTextStyle);
        }

        public string Link { get; private set; }

        public override string Build() {

            var inlineStyle = this.BuildInlineStyleAttribute(this.InlineTextStyle);
            var styleClass = this.BuildClassAttribute(this.StyleClass);
            return $"<span{styleClass}{inlineStyle}><a href='{this.Link}'>{this.Escape(this.Value)}</a></span>";
        }
    }

    public class TextBuilder : IContentBuilder {

        protected TextBuilder(string value, string? inlineTextStyle = null) {
            this.Value = value;
            this.InlineTextStyle = inlineTextStyle;
        }

        public string? InlineTextStyle { get; set; }
        public string StyleClass { get; set; } = DefaultHtmlStyle.Class_Text;
        public string Value { get; set; }

        public static TextBuilder Create(string value, string? inlineTextStyle = null) {
            return new TextBuilder(value, inlineTextStyle);
        }

        public static TextBuilder CreateInRed(string value) {
            return new TextBuilder(value, DefaultHtmlStyle.InlineStyle_Color_Red);
        }

        public static TextBuilder CreateInBlue(string value) {
            return new TextBuilder(value, DefaultHtmlStyle.InlineStyle_Color_Blue);
        }

        public virtual string Build() {

            var inlineStyle = this.BuildInlineStyleAttribute(this.InlineTextStyle);
            var styleClass = this.BuildClassAttribute(this.StyleClass);
            return $"<span{inlineStyle}{styleClass}>{this.Escape(this.Value) ?? string.Empty}</span>";
        }
    }

    public interface IContentBuilder {

        string Build();
    }

    public static class ContentBuilderUtil {

        public static string BuildInlineStyleAttribute(params string?[] styles) {

            if (styles == null || styles.Length == 0) return string.Empty;

            var items = styles.Where(s => !string.IsNullOrWhiteSpace(s)).ToArray();

            if (items.Length == 0) return string.Empty;
            else return $" style='{string.Join(';' + string.Empty, items)}'";
        }

        public static string BuildClassAttribute(params string?[] classes) {

            if (classes == null || classes.Length == 0) return string.Empty;

            var items = classes.Where(s => !string.IsNullOrWhiteSpace(s)).ToArray();

            if (items.Length == 0) return string.Empty;
            else return $" class='{string.Join(' ', items)}'";
        }

        public static string BuildInlineStyleAttribute(this IContentBuilder builder, params string?[] styles) {
            return BuildInlineStyleAttribute(styles);
        }

        public static string BuildClassAttribute(this IContentBuilder builder, params string?[] classes) {
            return BuildClassAttribute(classes);
        }

        public static string Escape(this IContentBuilder builder, string? text) {
            return SecurityElement.Escape(text) ?? string.Empty;
        }
    }
}
