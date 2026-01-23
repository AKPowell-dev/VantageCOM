using System.Collections.Generic;
using System.Text.RegularExpressions;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class AbbreviationBillions : BaseTextError
{
	public AbbreviationBillions(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, List<TextRange2> listRanges, string strFix)
		: base(ErrorType.Text, ((Settings)Main.Analysis.Options).MillionsBillionsAbbreviation, sld, shp, listRanges, blnHasFix: true)
	{
		//IL_0010: Unknown result type (might be due to invalid IL or missing references)
		BaseError val = (BaseError)(object)this;
		Errors.AbbreviationBillions(ref val, strFix);
	}

	public override void FixAction(int i)
	{
		NG.A.Application.StartNewUndoEntry();
		IEnumerator<TextRange2> enumerator = default(IEnumerator<TextRange2>);
		try
		{
			enumerator = ((BaseError)this).TextRanges.GetEnumerator();
			while (enumerator.MoveNext())
			{
				TextRange2 current = enumerator.Current;
				current.Text = Regex.Replace(current.Text, AH.A(41698), ((BaseError)this).ReplacementText[0]);
				_ = null;
			}
			while (true)
			{
				switch (5)
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
