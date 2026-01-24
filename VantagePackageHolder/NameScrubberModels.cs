using Excel = Microsoft.Office.Interop.Excel;

namespace VantagePackageHolder
{
    internal sealed class NameScrubberItem
    {
        public NameScrubberItem(Excel.Name name)
        {
            Name = name;
        }

        public Excel.Name Name { get; }
        public string Label { get; set; }
        public string Text { get; set; }
        public string ParentName { get; set; }
        public string RefersTo { get; set; }
        public bool IsChecked { get; set; }

        public bool? CachedIsErroneous { get; set; }
        public bool? CachedIsLinked { get; set; }
        public bool? CachedIsLambda { get; set; }
        public bool? CachedHasDependents { get; set; }
    }

    internal sealed class NameDependentItem
    {
        public NameDependentItem(Excel.Range range, string label, string formula)
        {
            Range = range;
            Label = label;
            Formula = formula;
        }

        public Excel.Range Range { get; }
        public string Label { get; }
        public string Formula { get; }
    }
}
