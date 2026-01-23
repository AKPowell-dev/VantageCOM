using System;
using System.IO;
using System.Reflection;
using System.Runtime.CompilerServices;
using A;
using MacabacusMacros;
using MacabacusMacros.Auth;
using MacabacusMacros.UI;
using Microsoft.Office.Interop.Word;
using Microsoft.VisualBasic.CompilerServices;

namespace Macabacus_Word.Publishing;

public sealed class Pdf
{
	public static void ToFolder()
	{
		if (!Access.AllowWordOperation((PlanType)4, (Restriction)1, false))
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
			Application application = PC.A.Application;
			Document activeDocument;
			try
			{
				activeDocument = application.ActiveDocument;
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				application = null;
				ProjectData.ClearProjectError();
				return;
			}
			string text;
			if (activeDocument.Path.Length > 0)
			{
				if (!activeDocument.ActiveWindow.View.ShowRevisionsAndComments)
				{
					goto IL_0148;
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
				if (activeDocument.Revisions.Count <= 0)
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
					if (activeDocument.Comments.Count <= 0)
					{
						goto IL_0148;
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
				}
				text = Path.Combine(activeDocument.Path, Path.GetFileNameWithoutExtension(activeDocument.Name) + XC.A(39335));
				if (!clsPublish.CancelOverwrite(text))
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
					A(activeDocument, text, WdExportItem.wdExportDocumentWithMarkup);
				}
				text = Path.Combine(activeDocument.Path, Path.GetFileNameWithoutExtension(activeDocument.Name) + XC.A(39358));
				if (!clsPublish.CancelOverwrite(text))
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
					A(activeDocument, text, WdExportItem.wdExportDocumentContent);
				}
				goto IL_018a;
			}
			Forms.WarningMessage(XC.A(39423));
			goto IL_01b0;
			IL_01b0:
			application = null;
			activeDocument = null;
			return;
			IL_018a:
			clsReporting.LogActivity((ActivityApp)3, (ActivityCategory)4, XC.A(39396));
			goto IL_01b0;
			IL_0148:
			text = Path.Combine(activeDocument.Path, Path.GetFileNameWithoutExtension(activeDocument.Name) + XC.A(39387));
			if (!clsPublish.CancelOverwrite(text))
			{
				A(activeDocument, text, WdExportItem.wdExportDocumentContent);
			}
			goto IL_018a;
		}
	}

	private static void A(Document A, string B, WdExportItem C)
	{
		try
		{
			object FixedFormatExtClassPtr = RuntimeHelpers.GetObjectValue(Missing.Value);
			A.ExportAsFixedFormat(B, WdExportFormat.wdExportFormatPDF, OpenAfterExport: true, WdExportOptimizeFor.wdExportOptimizeForPrint, WdExportRange.wdExportAllDocument, 1, 1, C, IncludeDocProps: false, KeepIRM: true, WdExportCreateBookmarks.wdExportCreateNoBookmarks, DocStructureTags: true, BitmapMissingFonts: true, UseISO19005_1: false, ref FixedFormatExtClassPtr);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			Forms.ErrorMessage(ex2.Message);
			clsReporting.LogException(ex2);
			ProjectData.ClearProjectError();
		}
	}
}
