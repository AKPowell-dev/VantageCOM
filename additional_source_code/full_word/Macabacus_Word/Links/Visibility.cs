using System;
using System.Collections;
using A;
using MacabacusMacros;
using MacabacusMacros.UI;
using Microsoft.Office.Interop.Word;
using Microsoft.VisualBasic.CompilerServices;

namespace Macabacus_Word.Links;

public sealed class Visibility
{
	public static void HideTextLinks()
	{
		A(WdContentControlAppearance.wdContentControlHidden, XC.A(14711));
		Forms.SuccessMessage(XC.A(14742));
	}

	public static void ShowTextLinks()
	{
		A(WdContentControlAppearance.wdContentControlBoundingBox, XC.A(14779));
		Forms.SuccessMessage(XC.A(14810));
	}

	private static void A(WdContentControlAppearance A, string B)
	{
		Application application = PC.A.Application;
		Document activeDocument = application.ActiveDocument;
		UndoRecord undoRecord = application.UndoRecord;
		bool flag = false;
		undoRecord.StartCustomRecord(B);
		application.ScreenUpdating = false;
		if (activeDocument.TrackRevisions)
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
			if (activeDocument.TrackFormatting)
			{
				activeDocument.TrackFormatting = false;
				flag = true;
			}
		}
		try
		{
			_ = activeDocument.Sections[1].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.StoryType;
			IEnumerator enumerator = default(IEnumerator);
			try
			{
				enumerator = activeDocument.StoryRanges.GetEnumerator();
				while (enumerator.MoveNext())
				{
					Range range = (Range)enumerator.Current;
					do
					{
						Visibility.A(range, A);
						range = range.NextStoryRange;
					}
					while (range != null);
				}
			}
			finally
			{
				if (enumerator is IDisposable)
				{
					while (true)
					{
						switch (6)
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
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			Forms.ErrorMessage(ex2.Message);
			clsReporting.LogException(ex2);
			ProjectData.ClearProjectError();
		}
		if (flag)
		{
			activeDocument.TrackFormatting = true;
		}
		application.ScreenUpdating = true;
		undoRecord.EndCustomRecord();
		Common.LogActivity(B);
		activeDocument = null;
		undoRecord = null;
		application = null;
	}

	private static void A(Range A, WdContentControlAppearance B)
	{
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = A.ContentControls.GetEnumerator();
			while (enumerator.MoveNext())
			{
				ContentControl contentControl = (ContentControl)enumerator.Current;
				if (Common.IsLinked(contentControl))
				{
					contentControl.Appearance = B;
				}
			}
		}
		finally
		{
			if (enumerator is IDisposable)
			{
				while (true)
				{
					switch (6)
					{
					case 0:
						continue;
					}
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					(enumerator as IDisposable).Dispose();
					break;
				}
			}
		}
	}
}
