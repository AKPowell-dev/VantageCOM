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

public sealed class Send
{
	public static void ShowDialog()
	{
		if (!Access.AllowWordOperation((PlanType)4, (Restriction)2, false))
		{
			return;
		}
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
			Application application = PC.A.Application;
			Document document = null;
			string text = string.Empty;
			string text2 = string.Empty;
			try
			{
				document = application.ActiveDocument;
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			WdExportRange c;
			wpfSend wpfSend2;
			if (document != null)
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
				wpfSend2 = new wpfSend(document);
				wpfSend2.ShowDialog();
				if (wpfSend2.DialogResult.HasValue)
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
					if (wpfSend2.DialogResult.Value)
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
						if (wpfSend2.chkSendFile.IsChecked == true)
						{
							string text3 = wpfSend2.txtName.Text;
							if (document.Path.Length > 0)
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
								text3 += Path.GetExtension(document.Name);
							}
							else
							{
								text3 += XC.A(39526);
							}
							if (wpfSend2.radScopeSelected.IsChecked == true)
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
								text2 = A(document, text3);
							}
							else
							{
								text2 = B(document, text3);
							}
						}
						if (wpfSend2.chkSendPdf.IsChecked == true)
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
							c = ((wpfSend2.radScopeSelected.IsChecked == true) ? WdExportRange.wdExportFromTo : WdExportRange.wdExportAllDocument);
							text = clsPublish.PdfFullName(wpfSend2.txtName.Text + XC.A(39387), document.Path, wpfSend2.chkSaveCopy.IsChecked.Value);
							bool? isChecked = wpfSend2.chkSaveCopy.IsChecked;
							if (isChecked.HasValue)
							{
								while (true)
								{
									switch (5)
									{
									case 0:
										continue;
									}
									break;
								}
								if (isChecked != true)
								{
									goto IL_026a;
								}
								while (true)
								{
									switch (7)
									{
									case 0:
										continue;
									}
									break;
								}
							}
							if (clsPublish.CancelOverwrite(text))
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
								if (isChecked.HasValue)
								{
									goto IL_032b;
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
							}
							goto IL_026a;
						}
						goto IL_02c2;
					}
				}
				goto IL_032b;
			}
			goto IL_0330;
			IL_0330:
			application = null;
			return;
			IL_032b:
			wpfSend2 = null;
			document = null;
			goto IL_0330;
			IL_02c2:
			if (wpfSend2.chkSendLink.IsChecked == true)
			{
				while (true)
				{
					switch (5)
					{
					case 0:
						continue;
					}
					break;
				}
				clsPublish.AttachToEmail(document.FullName, true);
			}
			else
			{
				clsPublish.SendAttachment(text, text2, wpfSend2.chkOpen, wpfSend2.chkSaveCopy, wpfSend2.chkCompress);
			}
			clsReporting.LogActivity((ActivityApp)2, (ActivityCategory)4, XC.A(39537));
			goto IL_032b;
			IL_026a:
			A(document, text, c, wpfSend2.chkCompress.IsChecked.Value, wpfSend2.chkOpen.IsChecked.Value, wpfSend2.chkSaveCopy.IsChecked.Value);
			goto IL_02c2;
		}
	}

	private static void A(Document A, string B, WdExportRange C, bool D, bool E, bool F)
	{
		try
		{
			object FixedFormatExtClassPtr;
			if (C == WdExportRange.wdExportAllDocument)
			{
				FixedFormatExtClassPtr = RuntimeHelpers.GetObjectValue(Missing.Value);
				A.ExportAsFixedFormat(B, WdExportFormat.wdExportFormatPDF, E, WdExportOptimizeFor.wdExportOptimizeForPrint, C, 1, 1, WdExportItem.wdExportDocumentContent, IncludeDocProps: false, KeepIRM: true, WdExportCreateBookmarks.wdExportCreateNoBookmarks, DocStructureTags: true, BitmapMissingFonts: true, UseISO19005_1: false, ref FixedFormatExtClassPtr);
				return;
			}
			Range range = A.Application.ActiveWindow.Selection.Range;
			Range duplicate = range.Duplicate;
			FixedFormatExtClassPtr = WdCollapseDirection.wdCollapseStart;
			duplicate.Collapse(ref FixedFormatExtClassPtr);
			int num = Conversions.ToInteger(duplicate.get_Information(WdInformation.wdActiveEndPageNumber));
			Range duplicate2 = range.Duplicate;
			FixedFormatExtClassPtr = WdCollapseDirection.wdCollapseEnd;
			duplicate2.Collapse(ref FixedFormatExtClassPtr);
			int to = Conversions.ToInteger(duplicate2.get_Information(WdInformation.wdActiveEndPageNumber));
			FixedFormatExtClassPtr = RuntimeHelpers.GetObjectValue(Missing.Value);
			A.ExportAsFixedFormat(B, WdExportFormat.wdExportFormatPDF, E, WdExportOptimizeFor.wdExportOptimizeForPrint, C, num, to, WdExportItem.wdExportDocumentContent, IncludeDocProps: false, KeepIRM: true, WdExportCreateBookmarks.wdExportCreateNoBookmarks, DocStructureTags: true, BitmapMissingFonts: true, UseISO19005_1: false, ref FixedFormatExtClassPtr);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			Forms.ErrorMessage(XC.A(39566) + Path.GetFileName(B) + XC.A(39603) + ex2.Message);
			ProjectData.ClearProjectError();
		}
		finally
		{
		}
	}

	private static string A(Document A, string B)
	{
		Application application = PC.A.Application;
		string empty = string.Empty;
		Range range;
		try
		{
			range = application.ActiveWindow.Selection.Range;
			Window activeWindow = application.ActiveWindow;
			Selection selection = activeWindow.Selection;
			application.ScreenUpdating = false;
			range = selection.Range;
			object Direction = WdCollapseDirection.wdCollapseStart;
			selection.Collapse(ref Direction);
			int num = Conversions.ToInteger(selection.get_Information(WdInformation.wdActiveEndPageNumber));
			range.Select();
			Selection selection2 = activeWindow.Selection;
			Direction = WdCollapseDirection.wdCollapseEnd;
			selection2.Collapse(ref Direction);
			int num2 = Conversions.ToInteger(selection2.get_Information(WdInformation.wdActiveEndPageNumber));
			range.Select();
			application.ScreenUpdating = true;
			range = null;
			Selection selection3 = application.Selection;
			Direction = WdGoToItem.wdGoToPage;
			object Which = WdGoToDirection.wdGoToFirst;
			object Count = num;
			object Name = RuntimeHelpers.GetObjectValue(Missing.Value);
			selection3.GoTo(ref Direction, ref Which, ref Count, ref Name);
			num = Conversions.ToInteger(Count);
			Range range2 = application.Selection.Range;
			Selection selection4 = application.Selection;
			Name = WdGoToItem.wdGoToPage;
			Count = WdGoToDirection.wdGoToFirst;
			Which = num2;
			Direction = RuntimeHelpers.GetObjectValue(Missing.Value);
			selection4.GoTo(ref Name, ref Count, ref Which, ref Direction);
			num2 = Conversions.ToInteger(Which);
			Bookmarks bookmarks = application.Selection.Bookmarks;
			Direction = XC.A(39779);
			range2.End = bookmarks[ref Direction].Range.End;
			range2.Select();
			range = application.ActiveWindow.Selection.Range;
			Documents documents = application.Documents;
			Direction = RuntimeHelpers.GetObjectValue(Missing.Value);
			Which = RuntimeHelpers.GetObjectValue(Missing.Value);
			Count = RuntimeHelpers.GetObjectValue(Missing.Value);
			Name = RuntimeHelpers.GetObjectValue(Missing.Value);
			Document document = documents.Add(ref Direction, ref Which, ref Count, ref Name);
			PageSetup pageSetup = document.PageSetup;
			pageSetup.PageHeight = A.PageSetup.PageHeight;
			pageSetup.PageWidth = A.PageSetup.PageWidth;
			_ = null;
			range.Copy();
			application.CommandBars.ExecuteMso(XC.A(39790));
			B = Path.Combine(L.A.FileSystem.SpecialDirectories.Temp, Path.GetFileNameWithoutExtension(A.Name));
			Name = B;
			Count = WdSaveFormat.wdFormatDocumentDefault;
			Which = RuntimeHelpers.GetObjectValue(Missing.Value);
			Direction = RuntimeHelpers.GetObjectValue(Missing.Value);
			object AddToRecentFiles = RuntimeHelpers.GetObjectValue(Missing.Value);
			object WritePassword = RuntimeHelpers.GetObjectValue(Missing.Value);
			object ReadOnlyRecommended = RuntimeHelpers.GetObjectValue(Missing.Value);
			object EmbedTrueTypeFonts = RuntimeHelpers.GetObjectValue(Missing.Value);
			object SaveNativePictureFormat = RuntimeHelpers.GetObjectValue(Missing.Value);
			object SaveFormsData = RuntimeHelpers.GetObjectValue(Missing.Value);
			object SaveAsAOCELetter = RuntimeHelpers.GetObjectValue(Missing.Value);
			object Encoding = RuntimeHelpers.GetObjectValue(Missing.Value);
			object InsertLineBreaks = RuntimeHelpers.GetObjectValue(Missing.Value);
			object AllowSubstitutions = RuntimeHelpers.GetObjectValue(Missing.Value);
			object LineEnding = RuntimeHelpers.GetObjectValue(Missing.Value);
			object AddBiDiMarks = RuntimeHelpers.GetObjectValue(Missing.Value);
			document.SaveAs(ref Name, ref Count, ref Which, ref Direction, ref AddToRecentFiles, ref WritePassword, ref ReadOnlyRecommended, ref EmbedTrueTypeFonts, ref SaveNativePictureFormat, ref SaveFormsData, ref SaveAsAOCELetter, ref Encoding, ref InsertLineBreaks, ref AllowSubstitutions, ref LineEnding, ref AddBiDiMarks);
			B = Conversions.ToString(Name);
			B = document.FullName;
			application.DisplayAlerts = WdAlertLevel.wdAlertsNone;
			document.Saved = true;
			AddBiDiMarks = RuntimeHelpers.GetObjectValue(Missing.Value);
			LineEnding = RuntimeHelpers.GetObjectValue(Missing.Value);
			AllowSubstitutions = RuntimeHelpers.GetObjectValue(Missing.Value);
			document.Close(ref AddBiDiMarks, ref LineEnding, ref AllowSubstitutions);
			application.DisplayAlerts = WdAlertLevel.wdAlertsAll;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			Forms.ErrorMessage(ex2.Message);
			ProjectData.ClearProjectError();
		}
		application = null;
		A = null;
		range = null;
		return empty;
	}

	private static string B(Document A, string B)
	{
		if (Operators.CompareString(A.Name, B, TextCompare: false) == 0)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					return A.FullName;
				}
			}
		}
		string text = Path.Combine(L.A.FileSystem.SpecialDirectories.Temp, B);
		object FileName = text;
		object FileFormat = RuntimeHelpers.GetObjectValue(Missing.Value);
		object LockComments = RuntimeHelpers.GetObjectValue(Missing.Value);
		object Password = RuntimeHelpers.GetObjectValue(Missing.Value);
		object AddToRecentFiles = RuntimeHelpers.GetObjectValue(Missing.Value);
		object WritePassword = RuntimeHelpers.GetObjectValue(Missing.Value);
		object ReadOnlyRecommended = RuntimeHelpers.GetObjectValue(Missing.Value);
		object EmbedTrueTypeFonts = RuntimeHelpers.GetObjectValue(Missing.Value);
		object SaveNativePictureFormat = RuntimeHelpers.GetObjectValue(Missing.Value);
		object SaveFormsData = RuntimeHelpers.GetObjectValue(Missing.Value);
		object SaveAsAOCELetter = RuntimeHelpers.GetObjectValue(Missing.Value);
		object Encoding = RuntimeHelpers.GetObjectValue(Missing.Value);
		object InsertLineBreaks = RuntimeHelpers.GetObjectValue(Missing.Value);
		object AllowSubstitutions = RuntimeHelpers.GetObjectValue(Missing.Value);
		object LineEnding = RuntimeHelpers.GetObjectValue(Missing.Value);
		object AddBiDiMarks = RuntimeHelpers.GetObjectValue(Missing.Value);
		object CompatibilityMode = RuntimeHelpers.GetObjectValue(Missing.Value);
		A.SaveCopyAs(ref FileName, ref FileFormat, ref LockComments, ref Password, ref AddToRecentFiles, ref WritePassword, ref ReadOnlyRecommended, ref EmbedTrueTypeFonts, ref SaveNativePictureFormat, ref SaveFormsData, ref SaveAsAOCELetter, ref Encoding, ref InsertLineBreaks, ref AllowSubstitutions, ref LineEnding, ref AddBiDiMarks, ref CompatibilityMode);
		return Conversions.ToString(FileName);
	}
}
