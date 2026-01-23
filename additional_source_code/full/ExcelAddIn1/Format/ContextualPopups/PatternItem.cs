using System.Windows.Media.Imaging;
using System.Xml;
using A;
using Microsoft.Office.Core;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Format.ContextualPopups;

public sealed class PatternItem
{
	private string A;

	private MsoPatternType A;

	private BitmapImage A;

	public string Name
	{
		get
		{
			return this.A;
		}
		set
		{
			this.A = value;
		}
	}

	public MsoPatternType Pattern
	{
		get
		{
			return this.A;
		}
		set
		{
			this.A = value;
		}
	}

	public BitmapImage Image
	{
		get
		{
			return A;
		}
		set
		{
			A = value;
		}
	}

	public PatternItem(XmlNode nd, bool blnChecked)
	{
		Pattern = (MsoPatternType)Conversions.ToInteger(nd.Attributes[VH.A(73267)].Value);
		switch (Pattern)
		{
		case MsoPatternType.msoPattern5Percent:
			Name = VH.A(144969);
			break;
		case MsoPatternType.msoPattern10Percent:
			Name = VH.A(144974);
			break;
		case MsoPatternType.msoPattern20Percent:
			Name = VH.A(144981);
			break;
		case MsoPatternType.msoPattern25Percent:
			Name = VH.A(144988);
			break;
		case MsoPatternType.msoPattern30Percent:
			Name = VH.A(144995);
			break;
		case MsoPatternType.msoPattern40Percent:
			Name = VH.A(145002);
			break;
		case MsoPatternType.msoPattern50Percent:
			Name = VH.A(145009);
			break;
		case MsoPatternType.msoPattern60Percent:
			Name = VH.A(145016);
			break;
		case MsoPatternType.msoPattern70Percent:
			Name = VH.A(145023);
			break;
		case MsoPatternType.msoPattern75Percent:
			Name = VH.A(145030);
			break;
		case MsoPatternType.msoPattern80Percent:
			Name = VH.A(145037);
			break;
		case MsoPatternType.msoPattern90Percent:
			Name = VH.A(145044);
			break;
		case MsoPatternType.msoPatternDarkHorizontal:
			Name = VH.A(145051);
			break;
		case MsoPatternType.msoPatternDarkVertical:
			Name = VH.A(145082);
			break;
		case MsoPatternType.msoPatternDarkDownwardDiagonal:
			Name = VH.A(145109);
			break;
		case MsoPatternType.msoPatternDarkUpwardDiagonal:
			Name = VH.A(145154);
			break;
		case MsoPatternType.msoPatternSmallCheckerBoard:
			Name = VH.A(145195);
			break;
		case MsoPatternType.msoPatternTrellis:
			Name = VH.A(145234);
			break;
		case MsoPatternType.msoPatternLightHorizontal:
			Name = VH.A(145249);
			break;
		case MsoPatternType.msoPatternLightVertical:
			Name = VH.A(145282);
			break;
		case MsoPatternType.msoPatternLightDownwardDiagonal:
			Name = VH.A(145311);
			break;
		case MsoPatternType.msoPatternLightUpwardDiagonal:
			Name = VH.A(145358);
			break;
		case MsoPatternType.msoPatternSmallGrid:
			Name = VH.A(145401);
			break;
		case MsoPatternType.msoPatternDottedDiamond:
			Name = VH.A(145422);
			break;
		case MsoPatternType.msoPatternWideDownwardDiagonal:
			Name = VH.A(145451);
			break;
		case MsoPatternType.msoPatternWideUpwardDiagonal:
			Name = VH.A(145496);
			break;
		case MsoPatternType.msoPatternDashedUpwardDiagonal:
			Name = VH.A(145537);
			break;
		case MsoPatternType.msoPatternDashedDownwardDiagonal:
			Name = VH.A(145582);
			break;
		case MsoPatternType.msoPatternNarrowVertical:
			Name = VH.A(145631);
			break;
		case MsoPatternType.msoPatternNarrowHorizontal:
			Name = VH.A(145662);
			break;
		case MsoPatternType.msoPatternDashedVertical:
			Name = VH.A(145697);
			break;
		case MsoPatternType.msoPatternDashedHorizontal:
			Name = VH.A(145728);
			break;
		case MsoPatternType.msoPatternLargeConfetti:
			Name = VH.A(145763);
			break;
		case MsoPatternType.msoPatternLargeGrid:
			Name = VH.A(145792);
			break;
		case MsoPatternType.msoPatternHorizontalBrick:
			Name = VH.A(145813);
			break;
		case MsoPatternType.msoPatternLargeCheckerBoard:
			Name = VH.A(145846);
			break;
		case MsoPatternType.msoPatternSmallConfetti:
			Name = VH.A(145885);
			break;
		case MsoPatternType.msoPatternZigZag:
			Name = VH.A(145914);
			break;
		case MsoPatternType.msoPatternSolidDiamond:
			Name = VH.A(145929);
			break;
		case MsoPatternType.msoPatternDiagonalBrick:
			Name = VH.A(145956);
			break;
		case MsoPatternType.msoPatternOutlinedDiamond:
			Name = VH.A(145985);
			break;
		case MsoPatternType.msoPatternPlaid:
			Name = VH.A(146018);
			break;
		case MsoPatternType.msoPatternSphere:
			Name = VH.A(146029);
			break;
		case MsoPatternType.msoPatternWeave:
			Name = VH.A(146042);
			break;
		case MsoPatternType.msoPatternDottedGrid:
			Name = VH.A(146053);
			break;
		case MsoPatternType.msoPatternDivot:
			Name = VH.A(146076);
			break;
		case MsoPatternType.msoPatternShingle:
			Name = VH.A(146087);
			break;
		case MsoPatternType.msoPatternWave:
			Name = VH.A(146102);
			break;
		case MsoPatternType.msoPatternHorizontal:
			Name = VH.A(56669);
			break;
		case MsoPatternType.msoPatternVertical:
			Name = VH.A(146111);
			break;
		case MsoPatternType.msoPatternCross:
			Name = VH.A(146128);
			break;
		case MsoPatternType.msoPatternDownwardDiagonal:
			Name = VH.A(146139);
			break;
		case MsoPatternType.msoPatternUpwardDiagonal:
			Name = VH.A(146174);
			break;
		case MsoPatternType.msoPatternDiagonalCross:
			Name = VH.A(146205);
			break;
		}
	}
}
