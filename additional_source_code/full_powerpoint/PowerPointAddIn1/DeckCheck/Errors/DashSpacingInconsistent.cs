using System.Collections.Generic;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class DashSpacingInconsistent : BaseTextError
{
	public DashSpacingInconsistent(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, List<TextRange2> listRanges)
		: base(ErrorType.Text, ((Settings)Main.Analysis.Options).PunctuationSpacingInconsistent, sld, shp, listRanges, blnHasFix: true, blnCanFixMultiple: true)
	{
		//IL_000e: Unknown result type (might be due to invalid IL or missing references)
		//IL_002f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0034: Unknown result type (might be due to invalid IL or missing references)
		BaseError val = (BaseError)(object)this;
		Errors.DashSpacingInconsistent(ref val, ((Settings)Main.Analysis.Options).DashSpacingConvention);
	}

	public override void FixAction(int i)
	{
		//IL_0021: Unknown result type (might be due to invalid IL or missing references)
		//IL_0026: Unknown result type (might be due to invalid IL or missing references)
		//IL_0029: Invalid comparison between Unknown and I4
		NG.A.Application.StartNewUndoEntry();
		IEnumerator<TextRange2> enumerator = default(IEnumerator<TextRange2>);
		if ((int)((Settings)Main.Analysis.Options).DashSpacingConvention == 1)
		{
			try
			{
				enumerator = ((BaseError)this).TextRanges.GetEnumerator();
				while (enumerator.MoveNext())
				{
					enumerator.Current.Text = AH.A(15092);
				}
				while (true)
				{
					switch (2)
					{
					case 0:
						break;
					default:
						if (1 == 0)
						{
							/*OpCode not supported: LdMemberToken*/;
						}
						return;
					}
				}
			}
			finally
			{
				if (enumerator != null)
				{
					while (true)
					{
						switch (2)
						{
						case 0:
							break;
						default:
							enumerator.Dispose();
							goto end_IL_0076;
						}
						continue;
						end_IL_0076:
						break;
					}
				}
			}
		}
		IEnumerator<TextRange2> enumerator2 = default(IEnumerator<TextRange2>);
		try
		{
			enumerator2 = ((BaseError)this).TextRanges.GetEnumerator();
			while (enumerator2.MoveNext())
			{
				enumerator2.Current.Text = AH.A(15125);
			}
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					return;
				}
			}
		}
		finally
		{
			if (enumerator2 != null)
			{
				while (true)
				{
					switch (2)
					{
					case 0:
						continue;
					}
					enumerator2.Dispose();
					break;
				}
			}
		}
	}
}
