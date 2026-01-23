namespace ExcelAddIn1.Audit.Check;

public enum Category
{
	FormulaErrors = 0,
	FormulaComplexity = 1,
	FormulaIntegrity = 2,
	BrandCompliance = 3,
	ModelStructure = 4,
	BestPractices = 5,
	Workbook = 6,
	Performance = 7,
	Data = 8,
	HiddenData = 9,
	PrivacySecurity = 10,
	Oddities = 11,
	Group = -1
}
