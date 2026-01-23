using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class IllegalFont : BaseTextError
{
	[CompilerGenerated]
	private new List<string> A;

	private List<string> LegalFonts
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

	public IllegalFont(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, List<TextRange2> listRanges, List<string> listFonts)
		: base(ErrorType.IllegalFont, Main.Analysis.Options.IllegalFonts, sld, shp, listRanges, blnHasFix: true, blnCanFixMultiple: true)
	{
		//IL_0011: Unknown result type (might be due to invalid IL or missing references)
		//IL_0016: Unknown result type (might be due to invalid IL or missing references)
		LegalFonts = listFonts;
		((BaseError)this).Title = AH.A(18320);
		((BaseError)this).Subtitle = AH.A(18345);
		((BaseError)this).Tooltip = ((BaseError)this).Subtitle;
		((BaseError)this).DisplayText = listFonts.ToList();
	}

	public override void FixAction(int i)
	{
		NG.A.Application.StartNewUndoEntry();
		string name = LegalFonts[i];
		IEnumerator<TextRange2> enumerator = default(IEnumerator<TextRange2>);
		try
		{
			enumerator = ((BaseError)this).TextRanges.GetEnumerator();
			while (enumerator.MoveNext())
			{
				enumerator.Current.Font.Name = name;
			}
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
				return;
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
					enumerator.Dispose();
					break;
				}
			}
		}
	}
}
