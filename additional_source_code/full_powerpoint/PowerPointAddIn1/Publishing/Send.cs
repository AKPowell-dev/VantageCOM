using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Windows.Forms;
using A;
using MacabacusMacros;
using MacabacusMacros.Auth;
using MacabacusMacros.UI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.Pagination;

namespace PowerPointAddIn1.Publishing;

public sealed class Send
{
	public static void ShowDialog(Microsoft.Office.Interop.PowerPoint.Presentation pres = null)
	{
		if (!Access.AllowPowerPointOperation((PlanType)4, (Restriction)2, false))
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
			Microsoft.Office.Interop.PowerPoint.Application application = NG.A.Application;
			string text = string.Empty;
			string text2 = string.Empty;
			if (pres == null)
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
				try
				{
					pres = application.ActivePresentation;
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					ProjectData.ClearProjectError();
				}
			}
			PpPrintRangeType c;
			wpfSend wpfSend2;
			if (pres != null)
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
				wpfSend2 = new wpfSend(pres);
				wpfSend2.ShowDialog();
				if (wpfSend2.DialogResult.HasValue && wpfSend2.DialogResult.Value)
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
					if (wpfSend2.chkSendFile.IsChecked == true)
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
						string text3 = wpfSend2.txtName.Text;
						if (pres.Path.Length > 0)
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
							text3 += Path.GetExtension(pres.Name);
						}
						else
						{
							text3 += AH.A(102167);
						}
						text2 = ((wpfSend2.radScopeSelected.IsChecked != true) ? B(pres, text3) : A(pres, text3));
					}
					if (wpfSend2.chkSendPdf.IsChecked != true)
					{
						goto IL_02c6;
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
					int num;
					if (wpfSend2.radScopeSelected.IsChecked != true)
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
						num = 1;
					}
					else
					{
						num = 2;
					}
					c = (PpPrintRangeType)num;
					text = clsPublish.PdfFullName(wpfSend2.txtName.Text + AH.A(104010), pres.Path, wpfSend2.chkSaveCopy.IsChecked.Value);
					bool? isChecked = wpfSend2.chkSaveCopy.IsChecked;
					if (isChecked.HasValue)
					{
						if (isChecked != true)
						{
							goto IL_0257;
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
					if (!clsPublish.CancelOverwrite(text) || !isChecked.HasValue)
					{
						goto IL_0257;
					}
				}
				goto IL_032c;
			}
			goto IL_0334;
			IL_0334:
			application = null;
			return;
			IL_032c:
			wpfSend2 = null;
			pres = null;
			goto IL_0334;
			IL_02c6:
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
				clsPublish.AttachToEmail(pres.FullName, true);
			}
			else
			{
				clsPublish.SendAttachment(text, text2, wpfSend2.chkOpen, wpfSend2.chkSaveCopy, wpfSend2.chkCompress);
			}
			clsReporting.LogActivity((ActivityApp)2, (ActivityCategory)4, AH.A(104157));
			goto IL_032c;
			IL_0257:
			A(pres, text, c, wpfSend2.chkCompress.IsChecked.Value, wpfSend2.chkOpen.IsChecked.Value, wpfSend2.chkSaveCopy.IsChecked.Value, wpfSend2.chkDuplex.IsChecked.Value);
			goto IL_02c6;
		}
	}

	private static void A(Microsoft.Office.Interop.PowerPoint.Presentation A, string B, PpPrintRangeType C, bool D, bool E, bool F, bool G)
	{
		List<int> list = null;
		if (G)
		{
			A.Application.StartNewUndoEntry();
			list = BlankSlides.InsertAsNeeded(A);
		}
		try
		{
			A.ExportAsFixedFormat(B, PpFixedFormatType.ppFixedFormatTypePDF, PpFixedFormatIntent.ppFixedFormatIntentPrint, MsoTriState.msoFalse, PpPrintHandoutOrder.ppPrintHandoutVerticalFirst, PpPrintOutputType.ppPrintOutputSlides, MsoTriState.msoFalse, null, C, "", IncludeDocProperties: false, KeepIRMSettings: true, DocStructureTags: true, BitmapMissingFonts: true, UseISO19005_1: false, RuntimeHelpers.GetObjectValue(Missing.Value));
			if (E)
			{
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
					try
					{
						Process.Start(B);
					}
					catch (Exception ex)
					{
						ProjectData.SetProjectError(ex);
						Exception ex2 = ex;
						Forms.ErrorMessage(AH.A(104194) + ex2.Message);
						ProjectData.ClearProjectError();
					}
					break;
				}
			}
		}
		catch (Exception ex3)
		{
			ProjectData.SetProjectError(ex3);
			Exception ex4 = ex3;
			if (!clsFile.IsPathUrl(B) && File.Exists(B))
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
				Forms.ErrorMessage(AH.A(104311) + Path.GetFileName(B) + AH.A(104348) + ex4.Message);
			}
			else
			{
				Forms.ErrorMessage(AH.A(104524) + ex4.Message);
			}
			ProjectData.ClearProjectError();
		}
		if (list == null)
		{
			return;
		}
		checked
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					continue;
				}
				for (int i = A.Slides.Count; i >= 1; i += -1)
				{
					if (list.Contains(A.Slides[i].SlideIndex))
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
						A.Slides[i].Delete();
					}
					else
					{
						SlideNumbers.Reset(A.Slides[i]);
					}
				}
				while (true)
				{
					switch (6)
					{
					case 0:
						continue;
					}
					if (A.Designs.Count > 1)
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
						int num = A.Designs.Count;
						while (true)
						{
							if (num >= 1)
							{
								Design design = A.Designs[num];
								if (Operators.CompareString(design.Name, BlankSlides.BLANK_NAME, TextCompare: false) == 0)
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
									if (design.SlideMaster.CustomLayouts.Count == 1)
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
										if (Operators.CompareString(design.SlideMaster.CustomLayouts[1].Name, BlankSlides.BLANK_NAME, TextCompare: false) == 0)
										{
											try
											{
												design.SlideMaster.CustomLayouts[1].Delete();
												design.Delete();
											}
											catch (Exception ex5)
											{
												ProjectData.SetProjectError(ex5);
												Exception ex6 = ex5;
												ProjectData.ClearProjectError();
											}
											break;
										}
									}
								}
								design = null;
								num += -1;
								continue;
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
							break;
						}
					}
					list = null;
					return;
				}
			}
		}
	}

	private static string A(Microsoft.Office.Interop.PowerPoint.Presentation A, string B)
	{
		Microsoft.Office.Interop.PowerPoint.Application application = NG.A.Application;
		string text = string.Empty;
		SlideRange slideRange;
		try
		{
			slideRange = application.ActiveWindow.Selection.SlideRange;
			try
			{
				Microsoft.Office.Interop.PowerPoint.Presentation presentation = application.Presentations.Add(MsoTriState.msoFalse);
				Microsoft.Office.Interop.PowerPoint.Presentation presentation2 = presentation;
				PageSetup pageSetup = presentation2.PageSetup;
				pageSetup.SlideOrientation = A.PageSetup.SlideOrientation;
				pageSetup.SlideSize = A.PageSetup.SlideSize;
				pageSetup.SlideHeight = A.PageSetup.SlideHeight;
				pageSetup.SlideWidth = A.PageSetup.SlideWidth;
				_ = null;
				if (slideRange.Count == 1)
				{
					presentation2.Designs.Clone(slideRange[1].Design);
					presentation2.Designs[1].Delete();
					slideRange.Copy();
					SlideRange slideRange2 = presentation2.Slides.Paste(1);
					slideRange2.Design = slideRange[1].Design;
					slideRange2.Delete();
					_ = null;
					presentation2.Slides.Paste();
				}
				else
				{
					slideRange.Copy();
					presentation.NewWindow().Activate();
					application.CommandBars.ExecuteMso(AH.A(58900));
					System.Windows.Forms.Application.DoEvents();
				}
				try
				{
					text = Path.Combine(Path.GetTempPath(), B);
					presentation.SaveAs(text);
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					try
					{
						text = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), AH.A(5874), Path.GetFileName(A.Name));
						presentation.SaveAs(text);
					}
					catch (Exception ex3)
					{
						ProjectData.SetProjectError(ex3);
						Exception ex4 = ex3;
						Forms.ErrorMessage(AH.A(104645) + text + AH.A(104712) + ex4.Message);
						clsReporting.LogException(ex4);
						ProjectData.ClearProjectError();
					}
					ProjectData.ClearProjectError();
				}
				presentation2.Saved = MsoTriState.msoTrue;
				presentation2.Close();
				presentation2 = null;
			}
			catch (Exception ex5)
			{
				ProjectData.SetProjectError(ex5);
				Exception ex6 = ex5;
				Forms.ErrorMessage(ex6.Message);
				clsReporting.LogException(ex6);
				ProjectData.ClearProjectError();
			}
			finally
			{
				Microsoft.Office.Interop.PowerPoint.Presentation presentation = null;
			}
		}
		catch (Exception ex7)
		{
			ProjectData.SetProjectError(ex7);
			Exception ex8 = ex7;
			Forms.ErrorMessage(AH.A(101507));
			ProjectData.ClearProjectError();
		}
		application = null;
		A = null;
		slideRange = null;
		return text;
	}

	private static string B(Microsoft.Office.Interop.PowerPoint.Presentation A, string B)
	{
		if (Operators.CompareString(A.Name, B, TextCompare: false) == 0)
		{
			while (true)
			{
				switch (5)
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
		string text = Path.Combine(NB.A.FileSystem.SpecialDirectories.Temp, B);
		A.SaveCopyAs(text);
		return text;
	}
}
