using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Forms;
using A;
using MacabacusMacros;
using MacabacusMacros.Auth;
using MacabacusMacros.UI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.Template;

namespace PowerPointAddIn1.Presentation;

public sealed class Miscellaneous
{
	public static void CloseOthers(Microsoft.Office.Interop.PowerPoint.Presentation presThis = null)
	{
		if (!Licensing.AllowRestrictedMode())
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
			Microsoft.Office.Interop.PowerPoint.Application application = NG.A.Application;
			Microsoft.Office.Interop.PowerPoint.Presentation presentation;
			try
			{
				if (presThis == null)
				{
					presThis = application.ActivePresentation;
				}
				Presentations presentations = application.Presentations;
				for (int i = presentations.Count; i >= 1; i = checked(i + -1))
				{
					try
					{
						presentation = presentations[i];
						if (presentation == presThis)
						{
							continue;
						}
						while (true)
						{
							switch (7)
							{
							case 0:
								continue;
							}
							if (presentation.Saved == MsoTriState.msoFalse)
							{
								presentation.Windows[1].Activate();
							}
							presentation.Close();
							break;
						}
					}
					catch (Exception ex)
					{
						ProjectData.SetProjectError(ex);
						Exception ex2 = ex;
						ProjectData.ClearProjectError();
					}
				}
				while (true)
				{
					switch (7)
					{
					case 0:
						continue;
					}
					presentations = null;
					break;
				}
			}
			catch (Exception ex3)
			{
				ProjectData.SetProjectError(ex3);
				Exception ex4 = ex3;
				clsReporting.LogException(ex4);
				ProjectData.ClearProjectError();
			}
			presThis.Windows[1].Activate();
			presThis = null;
			presentation = null;
			application = null;
			A(AH.A(117752));
			return;
		}
	}

	public static void Reopen(Microsoft.Office.Interop.PowerPoint.Presentation pres = null)
	{
		int try0000_dispatch = -1;
		int num2 = default(int);
		int num = default(int);
		int num3 = default(int);
		string fullName = default(string);
		Microsoft.Office.Interop.PowerPoint.Application application = default(Microsoft.Office.Interop.PowerPoint.Application);
		bool flag = default(bool);
		while (true)
		{
			try
			{
				/*Note: ILSpy has introduced the following switch to emulate a goto from catch-block to try-block*/;
				switch (try0000_dispatch)
				{
				default:
					num2 = 1;
					if (!Licensing.AllowRestrictedMode())
					{
						goto end_IL_0000;
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
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					goto IL_0021;
				case 426:
					{
						num = num2;
						switch (num3)
						{
						case 1:
							break;
						default:
							goto end_IL_0000_2;
						}
						int num4 = num + 1;
						num = 0;
						switch (num4)
						{
						case 1:
							break;
						case 3:
							goto IL_0021;
						case 4:
							goto IL_0028;
						case 5:
							goto IL_0039;
						case 6:
							goto IL_003e;
						case 7:
							goto IL_0043;
						case 8:
							goto IL_0051;
						case 9:
							goto IL_0059;
						case 10:
							goto IL_0081;
						case 11:
							goto IL_008c;
						case 12:
							goto IL_00cd;
						case 14:
							goto IL_00d9;
						case 13:
						case 15:
							goto IL_00df;
						case 16:
							goto IL_00f0;
						case 17:
							goto IL_00fd;
						case 18:
							goto IL_0106;
						case 19:
							goto IL_011c;
						case 20:
							goto IL_0130;
						case 21:
							goto end_IL_0000_3;
						default:
							goto end_IL_0000_2;
						case 2:
						case 22:
							goto end_IL_0000;
						}
						goto default;
					}
					IL_0130:
					num2 = 20;
					pres = null;
					break;
					IL_00f0:
					num2 = 16;
					fullName = pres.FullName;
					goto IL_00fd;
					IL_00fd:
					num2 = 17;
					pres.Close();
					goto IL_0106;
					IL_0106:
					num2 = 18;
					application.Presentations.Open(fullName);
					goto IL_011c;
					IL_0021:
					ProjectData.ClearProjectError();
					num3 = 1;
					goto IL_0028;
					IL_0028:
					num2 = 4;
					application = NG.A.Application;
					goto IL_0039;
					IL_0039:
					num2 = 5;
					flag = false;
					goto IL_003e;
					IL_003e:
					num2 = 6;
					if (pres == null)
					{
						goto IL_0043;
					}
					goto IL_0051;
					IL_0043:
					num2 = 7;
					pres = application.ActivePresentation;
					goto IL_0051;
					IL_0051:
					num2 = 8;
					if (pres == null)
					{
						break;
					}
					goto IL_0059;
					IL_0059:
					num2 = 9;
					if (Operators.CompareString(pres.Name, pres.FullName, TextCompare: false) != 0)
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
						goto IL_0081;
					}
					goto IL_0130;
					IL_011c:
					num2 = 19;
					A(AH.A(117866));
					goto IL_0130;
					IL_0081:
					num2 = 10;
					if (pres.Saved == MsoTriState.msoFalse)
					{
						goto IL_008c;
					}
					goto IL_00df;
					IL_008c:
					num2 = 11;
					if (System.Windows.Forms.MessageBox.Show(AH.A(117777) + pres.Name + AH.A(17524), AH.A(5874), MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation) == DialogResult.OK)
					{
						goto IL_00cd;
					}
					goto IL_00d9;
					IL_00cd:
					num2 = 12;
					pres.Saved = MsoTriState.msoTrue;
					goto IL_00df;
					IL_00d9:
					num2 = 14;
					flag = true;
					goto IL_00df;
					IL_00df:
					num2 = 15;
					if (!flag)
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
						goto IL_00f0;
					}
					goto IL_011c;
					end_IL_0000_3:
					break;
				}
				num2 = 21;
				application = null;
				break;
				end_IL_0000_2:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 426;
				continue;
			}
			throw ProjectData.CreateProjectError(-2146828237);
			continue;
			end_IL_0000:
			break;
		}
		if (num == 0)
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
			ProjectData.ClearProjectError();
			return;
		}
	}

	public static Microsoft.Office.Interop.PowerPoint.Presentation Duplicate(Microsoft.Office.Interop.PowerPoint.Presentation presOld = null)
	{
		if (!Licensing.AllowRestrictedMode())
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					return null;
				}
			}
		}
		Microsoft.Office.Interop.PowerPoint.Application application = NG.A.Application;
		if (presOld == null)
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
			presOld = application.ActivePresentation;
		}
		Microsoft.Office.Interop.PowerPoint.Presentation presentation = Create.NewBlankPresentation(application);
		Design design = presentation.Designs[1];
		PageSetup pageSetup = presentation.PageSetup;
		pageSetup.SlideHeight = presOld.PageSetup.SlideHeight;
		pageSetup.SlideWidth = presOld.PageSetup.SlideWidth;
		_ = null;
		presOld.Slides.Range(RuntimeHelpers.GetObjectValue(Missing.Value)).Copy();
		application.CommandBars.ExecuteMso(AH.A(58900));
		System.Windows.Forms.Application.DoEvents();
		if (presentation.Designs.Count > 1)
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
			try
			{
				design.Delete();
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				clsReporting.LogException(ex2);
				ProjectData.ClearProjectError();
			}
		}
		application = null;
		design = null;
		A(AH.A(117905));
		return presentation;
	}

	public static void OpenFolder(Microsoft.Office.Interop.PowerPoint.Presentation pres = null)
	{
		if (!Licensing.AllowRestrictedMode())
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
			if (pres == null)
			{
				try
				{
					pres = NG.A.Application.ActivePresentation;
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					ProjectData.ClearProjectError();
				}
			}
			if (pres != null)
			{
				while (true)
				{
					switch (5)
					{
					case 0:
						break;
					default:
						try
						{
							clsFile.OpenExplorerToFile(pres.FullName);
							A(AH.A(117950));
							return;
						}
						catch (Exception ex3)
						{
							ProjectData.SetProjectError(ex3);
							Exception ex4 = ex3;
							clsReporting.LogException(ex4);
							ProjectData.ClearProjectError();
							return;
						}
					}
				}
			}
			Forms.WarningMessage(AH.A(117979));
			return;
		}
	}

	public static void AnalyzeFileSize()
	{
		if (!Access.AllowPowerPointOperation((PlanType)4, (Restriction)1, false))
		{
			return;
		}
		while (true)
		{
			switch (1)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			Microsoft.Office.Interop.PowerPoint.Presentation activePresentation;
			try
			{
				activePresentation = NG.A.Application.ActivePresentation;
				if (activePresentation.Saved != MsoTriState.msoFalse)
				{
					goto IL_0089;
				}
				while (true)
				{
					switch (4)
					{
					case 0:
						continue;
					}
					break;
				}
				DialogResult dialogResult = System.Windows.Forms.MessageBox.Show(AH.A(118028), AH.A(5874), MessageBoxButtons.YesNoCancel, MessageBoxIcon.Exclamation);
				if (dialogResult == DialogResult.Yes)
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
					activePresentation.Save();
					goto IL_0089;
				}
				if (dialogResult != DialogResult.Cancel)
				{
					goto IL_0089;
				}
				goto end_IL_0023;
				IL_0089:
				bool flag = false;
				try
				{
					IEnumerable<wpfAnalyzeFileSize> source = System.Windows.Application.Current.Windows.OfType<wpfAnalyzeFileSize>();
					if (source.Any())
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
						source.ElementAt(0).Activate();
						flag = true;
					}
					source = null;
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					ProjectData.ClearProjectError();
				}
				if (!flag)
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
					new wpfAnalyzeFileSize().Show();
					_ = null;
				}
				A(AH.A(118270));
				end_IL_0023:;
			}
			catch (Exception ex3)
			{
				ProjectData.SetProjectError(ex3);
				Exception ex4 = ex3;
				ProjectData.ClearProjectError();
			}
			activePresentation = null;
			return;
		}
	}

	public static bool IsProtectedView(bool SuppressMessages)
	{
		bool result = false;
		try
		{
			_ = NG.A.Application.ActiveProtectedViewWindow;
			if (!SuppressMessages)
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
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				Forms.WarningMessage(AH.A(118321));
			}
			result = true;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		return result;
	}

	internal static bool A()
	{
		Microsoft.Office.Interop.PowerPoint.Presentation activePresentation = NG.A.Application.ActivePresentation;
		if (!activePresentation.Final)
		{
			return activePresentation.ReadOnly == MsoTriState.msoFalse;
		}
		return false;
	}

	public static void CopyPath(Microsoft.Office.Interop.PowerPoint.Presentation pres = null)
	{
		if (!Licensing.AllowRestrictedMode())
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
			try
			{
				if (pres == null)
				{
					pres = NG.A.Application.ActivePresentation;
				}
				if (pres.Path.Length > 0)
				{
					while (true)
					{
						switch (1)
						{
						case 0:
							continue;
						}
						clsClipboard.SetText(pres.FullName);
						A(AH.A(118424));
						break;
					}
				}
				else
				{
					Forms.WarningMessage(AH.A(118443));
				}
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			pres = null;
			return;
		}
	}

	private static void A(string A)
	{
		clsReporting.LogActivity((ActivityApp)2, (ActivityCategory)8, A);
	}
}
