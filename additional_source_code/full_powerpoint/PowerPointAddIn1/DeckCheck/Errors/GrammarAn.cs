using System;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class GrammarAn : BaseTextError
{
	[CompilerGenerated]
	private new TextRange2 A;

	private TextRange2 ParentTextRange
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

	public GrammarAn(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, List<TextRange2> listRanges, TextRange2 parent)
		: base(ErrorType.Text, (Severity)3, sld, shp, listRanges, blnHasFix: true)
	{
		string text = ((listRanges.Count != 1) ? (AH.A(42825) + listRanges.Count + AH.A(44503)) : A((List<TextRange2>)((BaseError)this).TextRanges, shp));
		BaseError val = (BaseError)(object)this;
		Errors.GrammarAn(ref val, text);
		ParentTextRange = parent;
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
				TextRange2 textRange = current;
				if (textRange.Text.Length == 1)
				{
					while (true)
					{
						switch (4)
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
					if (Operators.CompareString(textRange.Text, AH.A(8100), TextCompare: false) == 0)
					{
						while (true)
						{
							switch (2)
							{
							case 0:
								continue;
							}
							break;
						}
						textRange.Text = AH.A(44472);
					}
					else
					{
						try
						{
							if (Regex.IsMatch(ParentTextRange.get_Characters(current.Start, 4).Text, AH.A(44477)))
							{
								while (true)
								{
									switch (6)
									{
									case 0:
										continue;
									}
									textRange.Text = AH.A(44498);
									break;
								}
							}
							else
							{
								textRange.Text = AH.A(14823);
							}
						}
						catch (Exception ex)
						{
							ProjectData.SetProjectError(ex);
							Exception ex2 = ex;
							textRange.Text = AH.A(14823);
							ProjectData.ClearProjectError();
						}
					}
				}
				else
				{
					textRange.Text = textRange.Text.Replace(AH.A(44472), AH.A(8100));
					textRange.Text = textRange.Text.Replace(AH.A(14823), AH.A(7902));
					textRange.Text = textRange.Text.Replace(AH.A(44498), AH.A(7902));
				}
				textRange = null;
			}
		}
		finally
		{
			if (enumerator != null)
			{
				while (true)
				{
					switch (7)
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
