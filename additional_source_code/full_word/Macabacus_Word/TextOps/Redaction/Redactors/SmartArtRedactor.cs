using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows;
using A;
using Macabacus_Word.TextOps.Redaction.Process;
using Microsoft.Office.Core;
using Microsoft.VisualBasic.CompilerServices;

namespace Macabacus_Word.TextOps.Redaction.Redactors;

public sealed class SmartArtRedactor
{
	private static readonly object A = XC.A(19485);

	public static wpfRedactSelection WpfRedactSelection;

	public static bool ApprovedMultipleBulletsInSmartArt = false;

	public static void RedactRangeInSmartArt(TextRange2 rng)
	{
		TextRedactor.RedactWordCore(rng);
	}

	public static bool RedactAllRangesInSmartArt(SmartArt smartArt)
	{
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = smartArt.AllNodes.GetEnumerator();
			IEnumerator<TextRange2> enumerator2 = default(IEnumerator<TextRange2>);
			while (enumerator.MoveNext())
			{
				SmartArtNode smartArtNode = (SmartArtNode)enumerator.Current;
				try
				{
					IEnumerable<TextRange2> wordListCore = RedactUtilities.GetWordListCore((TextRange2)NewLateBinding.LateGet(NewLateBinding.LateGet(RuntimeHelpers.GetObjectValue(smartArtNode.Shapes.Cast<object>().ElementAtOrDefault(0)), null, XC.A(19132), new object[0], null, null, null), null, XC.A(19153), new object[0], null, null, null), trimList: true);
					try
					{
						enumerator2 = wordListCore.GetEnumerator();
						while (enumerator2.MoveNext())
						{
							TextRedactor.RedactWordCore(enumerator2.Current);
						}
					}
					finally
					{
						if (enumerator2 != null)
						{
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
								enumerator2.Dispose();
								break;
							}
						}
					}
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					if (!ApprovedMultipleBulletsInSmartArt)
					{
						if (WpfRedactSelection.ProgressBarLaunched)
						{
							while (true)
							{
								switch (6)
								{
								case 0:
									continue;
								}
								break;
							}
							WpfRedactSelection.Dispatcher.Invoke([SpecialName] () =>
							{
								ApprovedMultipleBulletsInSmartArt = RedactUtilities.ShowYesNoDialogue(Window.GetWindow(WpfRedactSelection), Conversions.ToString(A));
							});
						}
						else
						{
							ApprovedMultipleBulletsInSmartArt = RedactUtilities.ShowYesNoDialogue(null, Conversions.ToString(A));
						}
						if (!ApprovedMultipleBulletsInSmartArt)
						{
							throw new MultipleBulletsInSmartArtException();
						}
					}
					ProjectData.ClearProjectError();
				}
			}
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					goto end_IL_015c;
				}
				continue;
				end_IL_015c:
				break;
			}
		}
		finally
		{
			if (enumerator is IDisposable)
			{
				while (true)
				{
					switch (4)
					{
					case 0:
						continue;
					}
					(enumerator as IDisposable).Dispose();
					break;
				}
			}
		}
		return true;
	}
}
