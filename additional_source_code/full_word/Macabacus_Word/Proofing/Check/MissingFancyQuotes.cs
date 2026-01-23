using System;
using System.Collections;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using A;
using MacabacusMacros.Proofing;
using Macabacus_Word.Proofing.Errors;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace Macabacus_Word.Proofing.Check;

public sealed class MissingFancyQuotes : BaseTextCheck
{
	public override void Check(Range rng, string strText)
	{
		if (!strText.Contains(Constants.DOUBLE_QUOTE_OPEN))
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			if (!strText.Contains(Constants.DOUBLE_QUOTE_CLOSE))
			{
				return;
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
			int num2 = 0;
			IEnumerator enumerator = default(IEnumerator);
			MatchCollection matchCollection;
			while (true)
			{
				if (num2 <= num)
				{
					try
					{
						string text = strText;
						matchCollection = Regex.Matches(strText, XC.A(2144) + array[num2] + XC.A(25260) + array[num2] + XC.A(2144) + array[num2 + 1] + XC.A(25267) + array[num2 + 1]);
						enumerator = matchCollection.GetEnumerator();
						try
						{
							while (enumerator.MoveNext())
							{
								Match match = (Match)enumerator.Current;
								text = Text.MaskText(text, match);
							}
							while (true)
							{
								switch (1)
								{
								case 0:
									break;
								default:
									goto end_IL_00f7;
								}
								continue;
								end_IL_00f7:
								break;
							}
						}
						finally
						{
							IDisposable disposable = enumerator as IDisposable;
							if (disposable != null)
							{
								disposable.Dispose();
							}
						}
						if (Strings.Split(text, array[num2]).Length != Strings.Split(text, array[num2 + 1]).Length)
						{
							while (true)
							{
								switch (4)
								{
								case 0:
									continue;
								}
								int num3 = text.IndexOf(array[num2]);
								if (num3 == -1)
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
									num3 = text.LastIndexOf(array[num2 + 1]);
								}
								num3++;
								Range duplicate = rng.Duplicate;
								duplicate.SetRange(rng.Characters[num3].Start, rng.Characters[num3 + 1].End);
								Main.Analysis.Errors.Add(new Macabacus_Word.Proofing.Errors.MissingFancyQuotes(duplicate));
								duplicate = null;
								goto end_IL_0275;
							}
						}
						if (text.IndexOf(array[num2]) > text.IndexOf(array[num2 + 1]))
						{
							int num4 = text.IndexOf(array[num2 + 1]) + 1;
							Range duplicate = rng.Duplicate;
							duplicate.SetRange(rng.Characters[num4].Start, rng.Characters[num4 + 1].End);
							Main.Analysis.Errors.Add(new Macabacus_Word.Proofing.Errors.MissingFancyQuotes(duplicate));
							duplicate = null;
							break;
						}
					}
					catch (Exception ex)
					{
						ProjectData.SetProjectError(ex);
						Exception ex2 = ex;
						ProjectData.ClearProjectError();
					}
					num2 += 2;
					continue;
				}
				while (true)
				{
					switch (5)
					{
					case 0:
						continue;
					}
					break;
				}
				break;
				continue;
				end_IL_0275:
				break;
			}
			matchCollection = null;
		}
	}

	public override void Check(object shp, TextRange2 rng, string strText)
	{
		if (!strText.Contains(Constants.DOUBLE_QUOTE_OPEN))
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			if (!strText.Contains(Constants.DOUBLE_QUOTE_CLOSE))
			{
				return;
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
			MatchCollection matchCollection;
			for (int i = 0; i <= num; i += 2)
			{
				try
				{
					string text = strText;
					matchCollection = Regex.Matches(strText, XC.A(2144) + array[i] + XC.A(25260) + array[i] + XC.A(2144) + array[i + 1] + XC.A(25267) + array[i + 1]);
					try
					{
						enumerator = matchCollection.GetEnumerator();
						while (enumerator.MoveNext())
						{
							Match match = (Match)enumerator.Current;
							text = Text.MaskText(text, match);
						}
						while (true)
						{
							switch (4)
							{
							case 0:
								break;
							default:
								goto end_IL_00fc;
							}
							continue;
							end_IL_00fc:
							break;
						}
					}
					finally
					{
						if (enumerator is IDisposable)
						{
							while (true)
							{
								switch (5)
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
						int num2 = text.IndexOf(array[i]);
						if (num2 == -1)
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
							num2 = text.LastIndexOf(array[i + 1]);
						}
						num2++;
						Main.Analysis.Errors.Add(new Macabacus_Word.Proofing.Errors.MissingFancyQuotes(RuntimeHelpers.GetObjectValue(shp), rng.get_Characters(num2, 1)));
						break;
					}
					if (text.IndexOf(array[i]) <= text.IndexOf(array[i + 1]))
					{
						continue;
					}
					while (true)
					{
						switch (1)
						{
						case 0:
							continue;
						}
						int start = text.IndexOf(array[i + 1]) + 1;
						Main.Analysis.Errors.Add(new Macabacus_Word.Proofing.Errors.MissingFancyQuotes(RuntimeHelpers.GetObjectValue(shp), rng.get_Characters(start, 1)));
						goto end_IL_022b;
					}
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					ProjectData.ClearProjectError();
				}
				continue;
				end_IL_022b:
				break;
			}
			matchCollection = null;
		}
	}
}
