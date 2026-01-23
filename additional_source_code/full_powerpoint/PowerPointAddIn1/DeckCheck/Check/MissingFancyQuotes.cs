using System;
using System.Collections;
using System.Text.RegularExpressions;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.DeckCheck.Errors;

namespace PowerPointAddIn1.DeckCheck.Check;

public sealed class MissingFancyQuotes : BaseTextCheck
{
	public override void Check(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, TextRange2 para, string strText)
	{
		if (!strText.Contains(Constants.DOUBLE_QUOTE_OPEN))
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			if (!strText.Contains(Constants.DOUBLE_QUOTE_CLOSE))
			{
				return;
			}
			while (true)
			{
				switch (2)
				{
				case 0:
					continue;
				}
				break;
			}
		}
		string[] array = new string[2]
		{
			Constants.DOUBLE_QUOTE_OPEN,
			Constants.DOUBLE_QUOTE_CLOSE
		};
		checked
		{
			int num = array.Length - 1;
			IEnumerator enumerator = default(IEnumerator);
			for (int i = 0; i <= num; i += 2)
			{
				try
				{
					string text = strText;
					try
					{
						enumerator = Regex.Matches(strText, AH.A(17472) + array[i] + AH.A(17475) + array[i] + AH.A(17472) + array[i + 1] + AH.A(17482) + array[i + 1]).GetEnumerator();
						while (enumerator.MoveNext())
						{
							Match match = (Match)enumerator.Current;
							text = Text.MaskText(text, match);
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
					if (Strings.Split(text, array[i]).Length != Strings.Split(text, array[i + 1]).Length)
					{
						while (true)
						{
							switch (7)
							{
							case 0:
								continue;
							}
							int num2 = text.IndexOf(array[i]);
							if (num2 == -1)
							{
								num2 = text.LastIndexOf(array[i + 1]);
							}
							num2++;
							Main.Analysis.Errors.Add(new PowerPointAddIn1.DeckCheck.Errors.MissingFancyQuotes(sld, shp, para.get_Characters(num2, 1)));
							return;
						}
					}
					if (text.IndexOf(array[i]) <= text.IndexOf(array[i + 1]))
					{
						continue;
					}
					while (true)
					{
						switch (2)
						{
						case 0:
							continue;
						}
						int start = text.IndexOf(array[i + 1]) + 1;
						Main.Analysis.Errors.Add(new PowerPointAddIn1.DeckCheck.Errors.MissingFancyQuotes(sld, shp, para.get_Characters(start, 1)));
						return;
					}
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					ProjectData.ClearProjectError();
				}
			}
		}
	}
}
