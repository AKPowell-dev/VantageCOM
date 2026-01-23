using Microsoft.Office.Core;

namespace ExcelAddIn1.Charts.GrowthArrow;

public struct ArrowOptions
{
	public ArrowType LineType;

	public CagrLabelPosition LabelPosition;

	public float Weight;

	public int Color;

	public bool LabelBold;

	public bool LabelBorder;

	public bool Invert;

	public bool MatchColor;

	public bool Rotate;

	public string Format;

	public MsoAutoShapeType Shape;
}
