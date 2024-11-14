namespace me.fengyj.CommonLib.Office.Excel {
    public class OpenXmlValidationException : Exception {

        public string[] Errors { get; private set; }

        public OpenXmlValidationException(string[] errors) {
            this.Errors = errors;
        }

        public override string ToString() => $"Validation Failed. Errors: {string.Join(" ", this.Errors ?? [])}";
    }
}
