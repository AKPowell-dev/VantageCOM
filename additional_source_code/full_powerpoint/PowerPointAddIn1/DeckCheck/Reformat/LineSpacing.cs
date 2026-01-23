using Microsoft.Office.Core;

namespace PowerPointAddIn1.DeckCheck.Reformat;

public struct LineSpacing
{
	public float Before;

	public float After;

	public float Within;

	public MsoTriState LineRuleWithin;
}
