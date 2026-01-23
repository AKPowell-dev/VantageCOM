using System;
using System.Reflection;
using System.Runtime.CompilerServices;
using A;
using MacabacusMacros;
using Microsoft.Office.Interop.Word;
using Microsoft.VisualBasic.CompilerServices;

namespace Macabacus_Word.TextOps;

public sealed class Symbols
{
	public static void Insert(int num)
	{
		if (!Licensing.AllowRestrictedMode())
		{
			return;
		}
		while (true)
		{
			switch (7)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			Application application = PC.A.Application;
			UndoRecord undoRecord = application.UndoRecord;
			undoRecord.StartCustomRecord(XC.A(20016));
			application.ScreenUpdating = false;
			try
			{
				Selection selection = application.ActiveWindow.Selection;
				if (selection.Type == WdSelectionType.wdSelectionIP || selection.Type == WdSelectionType.wdSelectionNormal)
				{
					if (num != 108)
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
						selection.TypeText(Conversions.ToString(clsText.ConvertUnicodeToChar(num)));
					}
					else
					{
						Selection selection2 = selection;
						object Cset = XC.A(18458);
						object Count = WdConstants.wdBackward;
						selection2.MoveEndWhile(ref Cset, ref Count);
						Selection selection3 = selection;
						Count = XC.A(20043);
						Cset = false;
						object Bias = RuntimeHelpers.GetObjectValue(Missing.Value);
						selection3.InsertSymbol(108, ref Count, ref Cset, ref Bias);
					}
				}
				selection = null;
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			application.ScreenUpdating = true;
			undoRecord.EndCustomRecord();
			undoRecord = null;
			application = null;
			clsReporting.LogActivity((ActivityApp)3, (ActivityCategory)6, XC.A(20016));
			return;
		}
	}
}
