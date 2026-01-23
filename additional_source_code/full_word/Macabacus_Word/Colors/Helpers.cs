using System.Drawing;
using Microsoft.Office.Interop.Word;
using Microsoft.VisualBasic;

namespace Macabacus_Word.Colors;

public sealed class Helpers
{
	public static WdColor ColorToWdColor(Color clr)
	{
		return (WdColor)Information.RGB(clr.R, clr.G, clr.B);
	}
}
