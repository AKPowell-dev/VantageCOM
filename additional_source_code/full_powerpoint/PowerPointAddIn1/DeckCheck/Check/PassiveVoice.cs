using System;
using System.Collections;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using MacabacusMacros.Proofing.Check;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointAddIn1.DeckCheck.Errors;

namespace PowerPointAddIn1.DeckCheck.Check;

public sealed class PassiveVoice : BaseTextCheck
{
	public PassiveVoice()
	{
		base.RegexObj = Text.RegexPassiveVoice();
	}

	public override void Check(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, TextRange2 rng, string strText)
	{
		IEnumerator enumerator = default(IEnumerator);
		List<TextRange2> list;
		try
		{
			enumerator = rng.get_Sentences(-1, -1).GetEnumerator();
			IEnumerator enumerator2 = default(IEnumerator);
			while (enumerator.MoveNext())
			{
				TextRange2 textRange = (TextRange2)enumerator.Current;
				string text = textRange.Text;
				if (text.Length <= 5)
				{
					continue;
				}
				while (true)
				{
					switch (6)
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
				list = new List<TextRange2>();
				try
				{
					enumerator2 = base.RegexObj.Matches(text).GetEnumerator();
					while (enumerator2.MoveNext())
					{
						Group obj = ((Match)enumerator2.Current).Groups[1];
						list.Add(textRange.get_Characters(checked(obj.Index + 1), obj.Length));
						obj = null;
					}
					while (true)
					{
						switch (6)
						{
						case 0:
							break;
						default:
							goto end_IL_00b5;
						}
						continue;
						end_IL_00b5:
						break;
					}
				}
				finally
				{
					if (enumerator2 is IDisposable)
					{
						while (true)
						{
							switch (3)
							{
							case 0:
								continue;
							}
							(enumerator2 as IDisposable).Dispose();
							break;
						}
					}
				}
				if (list.Count <= 0)
				{
					continue;
				}
				while (true)
				{
					switch (3)
					{
					case 0:
						continue;
					}
					break;
				}
				Main.Analysis.Errors.Add(new PowerPointAddIn1.DeckCheck.Errors.PassiveVoice(sld, shp, list, text));
			}
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					goto end_IL_011f;
				}
				continue;
				end_IL_011f:
				break;
			}
		}
		finally
		{
			if (enumerator is IDisposable)
			{
				while (true)
				{
					switch (3)
					{
					case 0:
						continue;
					}
					(enumerator as IDisposable).Dispose();
					break;
				}
			}
		}
		list = null;
	}
}
