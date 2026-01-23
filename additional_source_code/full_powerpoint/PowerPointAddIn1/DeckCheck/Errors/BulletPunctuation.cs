using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class BulletPunctuation : BaseTextError
{
	[CompilerGenerated]
	private new List<bool> A;

	private List<bool> FixOptions
	{
		[CompilerGenerated]
		get
		{
			return A;
		}
		[CompilerGenerated]
		set
		{
			A = value;
		}
	}

	public BulletPunctuation(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, List<string> listLabels, string strSubtitle, List<TextRange2> listRanges, List<bool> listFixes)
		: base(ErrorType.BulletPunctuation, Main.Analysis.Options.BulletPunctuation, sld, shp, listRanges, blnHasFix: true)
	{
		//IL_0011: Unknown result type (might be due to invalid IL or missing references)
		//IL_0016: Unknown result type (might be due to invalid IL or missing references)
		((BaseError)this).DisplayText = listLabels;
		FixOptions = listFixes;
		((BaseError)this).Title = AH.A(42135);
		((BaseError)this).Subtitle = strSubtitle;
		((BaseError)this).Tooltip = AH.A(42198);
	}

	public override void FixAction(int i)
	{
		NG.A.Application.StartNewUndoEntry();
		bool flag = FixOptions[i];
		foreach (TextRange2 textRange in ((BaseError)this).TextRanges)
		{
			string text = Text.PrintableText(textRange.Text);
			if (!flag)
			{
				while (true)
				{
					switch (3)
					{
					case 0:
						continue;
					}
					break;
				}
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				if (Operators.CompareString(Conversions.ToString(text.Last()), AH.A(14417), TextCompare: false) == 0)
				{
					while (true)
					{
						switch (1)
						{
						case 0:
							continue;
						}
						break;
					}
					textRange.Text = Regex.Replace(text, AH.A(42128), "");
					continue;
				}
			}
			if (!flag)
			{
				continue;
			}
			while (true)
			{
				switch (1)
				{
				case 0:
					continue;
				}
				break;
			}
			if (Operators.CompareString(Conversions.ToString(text.Last()), AH.A(14417), TextCompare: false) == 0)
			{
				continue;
			}
			while (true)
			{
				switch (7)
				{
				case 0:
					continue;
				}
				break;
			}
			textRange.Text = text + AH.A(14417);
		}
	}
}
