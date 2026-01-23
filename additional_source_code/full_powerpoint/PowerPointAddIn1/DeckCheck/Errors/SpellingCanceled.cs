using System.Collections.Generic;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class SpellingCanceled : BaseTextError
{
	[CompilerGenerated]
	private new CanceledSpelling A;

	private CanceledSpelling Convention
	{
		[CompilerGenerated]
		get
		{
			//IL_0001: Unknown result type (might be due to invalid IL or missing references)
			return A;
		}
		[CompilerGenerated]
		set
		{
			//IL_0001: Unknown result type (might be due to invalid IL or missing references)
			//IL_0002: Unknown result type (might be due to invalid IL or missing references)
			A = value;
		}
	}

	public SpellingCanceled(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, List<TextRange2> listRanges, CanceledSpelling conv)
		: base(ErrorType.Text, ((Settings)Main.Analysis.Options).SpellingCanceled, sld, shp, listRanges, blnHasFix: true, blnCanFixMultiple: true)
	{
		//IL_0010: Unknown result type (might be due to invalid IL or missing references)
		//IL_0015: Unknown result type (might be due to invalid IL or missing references)
		//IL_00b1: Unknown result type (might be due to invalid IL or missing references)
		string text;
		if (listRanges.Count == 1)
		{
			text = A((List<TextRange2>)((BaseError)this).TextRanges, shp);
		}
		else
		{
			text = AH.A(43409);
			foreach (TextRange2 listRange in listRanges)
			{
				text = text + listRange.Text + AH.A(14258);
			}
			text = Strings.Left(text, checked(text.Length - 2));
		}
		BaseError val = (BaseError)(object)this;
		Errors.SpellingCanceled(ref val, text);
		Convention = conv;
	}

	public override void FixAction(int i)
	{
		//IL_002e: Unknown result type (might be due to invalid IL or missing references)
		NG.A.Application.StartNewUndoEntry();
		IEnumerator<TextRange2> enumerator = default(IEnumerator<TextRange2>);
		try
		{
			enumerator = ((BaseError)this).TextRanges.GetEnumerator();
			while (enumerator.MoveNext())
			{
				TextRange2 current = enumerator.Current;
				if ((int)Convention == 0)
				{
					while (true)
					{
						switch (7)
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
					current.Text = current.Text.Replace(AH.A(46928), AH.A(46943));
				}
				else
				{
					current.Text = Regex.Replace(current.Text, AH.A(46956), AH.A(46985));
				}
				current = null;
			}
			while (true)
			{
				switch (3)
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
			if (enumerator != null)
			{
				while (true)
				{
					switch (2)
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
