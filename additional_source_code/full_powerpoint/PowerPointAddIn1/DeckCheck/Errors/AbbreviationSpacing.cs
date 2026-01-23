using System.Collections.Generic;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class AbbreviationSpacing : BaseTextError
{
	[CompilerGenerated]
	private new int A;

	private int RequiredSpaces
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

	public AbbreviationSpacing(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, List<TextRange2> listRanges, int intRequiredSpaces)
		: base(ErrorType.Text, ((Settings)Main.Analysis.Options).MillionsBillionsAbbreviation, sld, shp, listRanges, blnHasFix: true)
	{
		//IL_000e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0013: Unknown result type (might be due to invalid IL or missing references)
		BaseError val = (BaseError)(object)this;
		Errors.AbbreviationSpacing(ref val, intRequiredSpaces);
		RequiredSpaces = intRequiredSpaces;
	}

	public override void FixAction(int i)
	{
		NG.A.Application.StartNewUndoEntry();
		if (RequiredSpaces == 0)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
				{
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					using IEnumerator<TextRange2> enumerator = ((BaseError)this).TextRanges.GetEnumerator();
					while (enumerator.MoveNext())
					{
						TextRange2 current = enumerator.Current;
						current.Text = Regex.Replace(current.Text, AH.A(41772) + Constants.REGEX_ABBREV_SPACING, AH.A(41785));
						_ = null;
					}
					while (true)
					{
						switch (4)
						{
						case 0:
							break;
						default:
							return;
						}
					}
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
				TextRange2 current2 = enumerator2.Current;
				current2.Text = Regex.Replace(current2.Text, AH.A(41794) + Constants.REGEX_ABBREV_SPACING, AH.A(41803));
				_ = null;
			}
			while (true)
			{
				switch (7)
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
					switch (7)
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
