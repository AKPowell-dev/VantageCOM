using System.Collections.Generic;
using System.Runtime.CompilerServices;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class BulletFontFamily : BaseTextError
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

	public BulletFontFamily(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, List<string> listLabels, string strSubtitle, List<TextRange2> listRanges, List<string> listFixes)
		: base(ErrorType.BulletFontFamily, Main.Analysis.Options.BulletFontFamily, sld, shp, listRanges, blnHasFix: true)
	{
		//IL_0011: Unknown result type (might be due to invalid IL or missing references)
		//IL_0016: Unknown result type (might be due to invalid IL or missing references)
		((BaseError)this).DisplayText = listLabels;
		FixOptions = listFixes;
		((BaseError)this).Title = AH.A(41814);
		((BaseError)this).Subtitle = strSubtitle;
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
				enumerator.Current.ParagraphFormat.Bullet.Font.Name = name;
			}
			while (true)
			{
				switch (6)
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
					switch (5)
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
