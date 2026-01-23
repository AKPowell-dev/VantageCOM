using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.CompilerServices;
using A;
using MacabacusMacros;
using MacabacusMacros.Auth;
using MacabacusMacros.UI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1.Publishing;

public sealed class Pdf
{
	public static void ToFolder()
	{
		if (!Access.AllowPowerPointOperation((PlanType)4, (Restriction)1, false))
		{
			return;
		}
		while (true)
		{
			switch (4)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			Application application = NG.A.Application;
			Microsoft.Office.Interop.PowerPoint.Presentation activePresentation;
			try
			{
				activePresentation = application.ActivePresentation;
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				application = null;
				ProjectData.ClearProjectError();
				return;
			}
			if (application.ActiveWindow.Selection.Type == PpSelectionType.ppSelectionSlides)
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
				if (activePresentation.Path.Length > 0)
				{
					string text = Path.Combine(activePresentation.Path, Path.GetFileNameWithoutExtension(activePresentation.Name) + AH.A(104010));
					if (!clsPublish.CancelOverwrite(text))
					{
						try
						{
							activePresentation.ExportAsFixedFormat(text, PpFixedFormatType.ppFixedFormatTypePDF, PpFixedFormatIntent.ppFixedFormatIntentPrint, MsoTriState.msoFalse, PpPrintHandoutOrder.ppPrintHandoutVerticalFirst, PpPrintOutputType.ppPrintOutputSlides, MsoTriState.msoFalse, null, PpPrintRangeType.ppPrintSelection, "", IncludeDocProperties: false, KeepIRMSettings: true, DocStructureTags: true, BitmapMissingFonts: true, UseISO19005_1: false, RuntimeHelpers.GetObjectValue(Missing.Value));
							Process.Start(text);
							clsReporting.LogActivity((ActivityApp)2, (ActivityCategory)4, AH.A(104019));
						}
						catch (Exception ex3)
						{
							ProjectData.SetProjectError(ex3);
							Exception ex4 = ex3;
							Forms.ErrorMessage(ex4.Message);
							clsReporting.LogException(ex4);
							ProjectData.ClearProjectError();
						}
					}
				}
				else
				{
					Forms.WarningMessage(AH.A(104046));
				}
			}
			else
			{
				Forms.WarningMessage(AH.A(101507));
			}
			application = null;
			activePresentation = null;
			return;
		}
	}
}
