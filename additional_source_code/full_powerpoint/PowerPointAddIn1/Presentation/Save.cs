using System;
using System.Collections;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using A;
using MacabacusMacros;
using MacabacusMacros.UI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1.Presentation;

public sealed class Save
{
	public static void All()
	{
		if (!Licensing.AllowRestrictedMode())
		{
			return;
		}
		IEnumerator enumerator = default(IEnumerator);
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
			bool flag = false;
			Microsoft.Office.Interop.PowerPoint.Presentation activePresentation;
			if (application.Presentations.Count > 0)
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
				activePresentation = application.ActivePresentation;
				try
				{
					enumerator = application.Presentations.GetEnumerator();
					while (enumerator.MoveNext())
					{
						Microsoft.Office.Interop.PowerPoint.Presentation presentation = (Microsoft.Office.Interop.PowerPoint.Presentation)enumerator.Current;
						try
						{
							Microsoft.Office.Interop.PowerPoint.Presentation presentation2 = presentation;
							if (presentation2.ReadOnly == MsoTriState.msoFalse)
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
								if (presentation2.Path.Length == 0)
								{
									presentation2.Windows[1].Activate();
									flag = true;
									SaveFileDialog saveFileDialog = new SaveFileDialog();
									saveFileDialog.DefaultExt = AH.A(116773);
									saveFileDialog.FileName = presentation.Name;
									saveFileDialog.Filter = AH.A(116782);
									if (saveFileDialog.ShowDialog() == DialogResult.OK)
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
										presentation.SaveAs(saveFileDialog.FileName);
										presentation.Saved = MsoTriState.msoTrue;
									}
									saveFileDialog = null;
								}
								else
								{
									presentation2.Save();
									presentation2.Saved = MsoTriState.msoTrue;
								}
							}
							presentation2 = null;
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
						switch (6)
						{
						case 0:
							break;
						default:
							goto end_IL_015f;
						}
						continue;
						end_IL_015f:
						break;
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
				if (flag)
				{
					activePresentation.Windows[1].Activate();
				}
				A(AH.A(116861));
			}
			activePresentation = null;
			application = null;
			return;
		}
	}

	public static void Up(Microsoft.Office.Interop.PowerPoint.Presentation pres)
	{
		if (!Licensing.AllowRestrictedMode())
		{
			return;
		}
		Microsoft.Office.Interop.PowerPoint.Application application = NG.A.Application;
		bool flag = true;
		try
		{
			string text3;
			if (application.Presentations.Count > 0)
			{
				if (pres.Path.Length == 0)
				{
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
						((_Application)pres.Application).get_FileDialog(MsoFileDialogType.msoFileDialogSaveAs).Show();
						break;
					}
				}
				else
				{
					string text = clsFile.BaseName(pres.Name);
					string extension = Path.GetExtension(pres.Name);
					int num = clsFile.VersionNumber(pres.Name);
					if (num == 0)
					{
						while (true)
						{
							switch (6)
							{
							case 0:
								continue;
							}
							Forms.WarningMessage(AH.A(116878));
							break;
						}
					}
					else
					{
						num = checked(num + 1);
						string text2 = text + Conversions.ToString(num) + extension;
						text3 = pres.Path + AH.A(17472) + text2;
						if (clsFile.IsPathUrl(text3))
						{
							goto IL_012f;
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
						if (!File.Exists(text3))
						{
							goto IL_012f;
						}
						while (true)
						{
							switch (7)
							{
							case 0:
								continue;
							}
							Forms.WarningMessage(AH.A(117038));
							break;
						}
					}
				}
			}
			goto end_IL_001c;
			IL_012f:
			if (clsFile.NewerVersions(text3).Any() && MessageBox.Show(AH.A(117139), AH.A(5874), MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
			{
				flag = false;
			}
			if (flag)
			{
				pres.SaveAs(text3);
				pres.Saved = MsoTriState.msoTrue;
			}
			end_IL_001c:;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			if (pres.ReadOnly == MsoTriState.msoTrue)
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
				if (clsFile.IsPathUrl(pres.FullName))
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
					if (ex2.Message.Contains(AH.A(117329)))
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
						Forms.ErrorMessage(AH.A(117362));
						goto IL_021c;
					}
				}
			}
			Forms.ErrorMessage(AH.A(117616) + ex2.Message);
			clsReporting.LogException(ex2);
			goto IL_021c;
			IL_021c:
			ProjectData.ClearProjectError();
		}
		application = null;
		A(AH.A(117737));
	}

	public static bool IsVersionInFileName(Microsoft.Office.Interop.PowerPoint.Presentation pres)
	{
		if (Operators.CompareString(pres.FullName, pres.Name, TextCompare: false) != 0)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					return clsFile.FileNameVersionRegex().IsMatch(pres.Name);
				}
			}
		}
		return false;
	}

	private static void A(string A)
	{
		clsReporting.LogActivity((ActivityApp)2, (ActivityCategory)8, A);
	}
}
