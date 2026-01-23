using System.Collections.Generic;
using System.Runtime.CompilerServices;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class MultipleFontFamilies : BaseTextError
{
	[CompilerGenerated]
	private new List<string> A;

	private List<string> FixOptions
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

	public MultipleFontFamilies(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, TextRange2 rng, List<string> listLabels, string strSubtitle, List<string> listFixes)
		: base(ErrorType.MultipleFontFamilies, ((Settings)Main.Analysis.Options).MultipleFontFamilies, sld, shp, rng, blnHasFix: true)
	{
		//IL_000f: Unknown result type (might be due to invalid IL or missing references)
		((BaseError)this).DisplayText = listLabels;
		FixOptions = listFixes;
		((BaseError)this).Title = AH.A(46059);
		((BaseError)this).Subtitle = strSubtitle;
		((BaseError)this).Tooltip = AH.A(46104);
	}

	public override void FixAction(int i)
	{
		NG.A.Application.StartNewUndoEntry();
		string name = FixOptions[i];
		IEnumerator<TextRange2> enumerator = default(IEnumerator<TextRange2>);
		try
		{
			enumerator = ((BaseError)this).TextRanges.GetEnumerator();
			while (enumerator.MoveNext())
			{
				enumerator.Current.Font.Name = name;
			}
		}
		finally
		{
			if (enumerator != null)
			{
				while (true)
				{
					switch (1)
					{
					case 0:
						continue;
					}
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					enumerator.Dispose();
					break;
				}
			}
		}
	}
}
