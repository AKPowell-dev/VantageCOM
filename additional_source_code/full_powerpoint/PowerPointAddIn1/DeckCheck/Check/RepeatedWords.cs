using System;
using System.Collections;
using System.Text.RegularExpressions;
using A;
using MacabacusMacros.Proofing.Check;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.DeckCheck.Errors;

namespace PowerPointAddIn1.DeckCheck.Check;

public sealed class RepeatedWords : BaseTextCheck
{
	public RepeatedWords()
	{
		base.RegexObj = Text.RegexRepeatedWords();
	}

	public override void Check(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, TextRange2 rng, string strText)
	{
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = base.RegexObj.Matches(strText).GetEnumerator();
			while (enumerator.MoveNext())
			{
				Group obj = ((Match)enumerator.Current).Groups[0];
				string left = obj.Value.ToLower();
				if (Operators.CompareString(left, AH.A(17527), TextCompare: false) != 0)
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
					if (Operators.CompareString(left, AH.A(17546), TextCompare: false) != 0)
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
						if (Operators.CompareString(left, AH.A(17561), TextCompare: false) != 0)
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
							if (Operators.CompareString(left, AH.A(17580), TextCompare: false) != 0)
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
								Main.Analysis.Errors.Add(new PowerPointAddIn1.DeckCheck.Errors.RepeatedWords(sld, shp, rng.get_Characters(checked(obj.Index + 1), obj.Length), obj.Value));
							}
						}
					}
				}
				obj = null;
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
		finally
		{
			if (enumerator is IDisposable)
			{
				while (true)
				{
					switch (7)
					{
					case 0:
						continue;
					}
					(enumerator as IDisposable).Dispose();
					break;
				}
			}
		}
	}
}
