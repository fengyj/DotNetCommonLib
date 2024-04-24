namespace me.fengyj.CommonLib.Office.Excel {
    public class CellFormulaValue {

        public CellFormulaValue(string formula) => this.Formula = formula;

        public string Formula { get; private set; }
    }
}
