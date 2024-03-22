namespace me.fengyj.CommonLib.Office.Excel {
    public class HyperLinkValue {

        public HyperLinkValue(string link, string? text = null) {
            this.Link = link;
            this.Text = text;
        }

        public string? Text { get; private set; }
        public string Link { get; private set; }

        public string? CellReferenceId { get; set; }

        public string DisplayName {
            get {
                return this.Text ?? this.Link;
            }
        }
    }
}
