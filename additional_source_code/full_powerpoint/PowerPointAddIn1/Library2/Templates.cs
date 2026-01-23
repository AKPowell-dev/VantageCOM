using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Runtime.CompilerServices;
using System.Windows.Forms;
using System.Xml;
using A;
using MacabacusMacros;
using MacabacusMacros.Libraries;
using MacabacusMacros.Libraries.Caching;
using MacabacusMacros.UI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.Agenda;
using PowerPointAddIn1.Presentation;
using PowerPointAddIn1.Shapes;
using PowerPointAddIn1.Template;

namespace PowerPointAddIn1.Library2;

public sealed class Templates
{
	[CompilerGenerated]
	private static Dictionary<string, string> m_A;

	private static Dictionary<string, string> TemplatesDictionary
	{
		[CompilerGenerated]
		get
		{
			return Templates.m_A;
		}
		[CompilerGenerated]
		set
		{
			Templates.m_A = value;
		}
	}

	public static string BuildNewPresentationMenu()
	{
		return Ribbon.BuildNewPresentationMenu();
	}

	public static string BuildNewSlideMenu()
	{
		return Ribbon.BuildNewSlideMenu();
	}

	public static string BuildApplyTemplateMenu()
	{
		return Ribbon.BuildApplyTemplateMenu();
	}

	public static void ApplyTemplate(string strTemplatePath)
	{
		Microsoft.Office.Interop.PowerPoint.Application application = NG.A.Application;
		Microsoft.Office.Interop.PowerPoint.Presentation A = null;
		bool B = false;
		Microsoft.Office.Interop.PowerPoint.Presentation presentation;
		try
		{
			presentation = application.ActivePresentation;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			Forms.WarningMessage(AH.A(59235));
			application = null;
			ProjectData.ClearProjectError();
			return;
		}
		bool flag = false;
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = presentation.Designs.GetEnumerator();
			while (true)
			{
				if (enumerator.MoveNext())
				{
					Design design = (Design)enumerator.Current;
					if (design.Preserved != MsoTriState.msoTrue)
					{
						continue;
					}
					if (flag)
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
						design.Preserved = MsoTriState.msoFalse;
						continue;
					}
					switch (MessageBox.Show(AH.A(68053), AH.A(5874), MessageBoxButtons.YesNoCancel, MessageBoxIcon.Exclamation))
					{
					default:
						while (true)
						{
							switch (1)
							{
							case 0:
								continue;
							}
							break;
						}
						continue;
					case DialogResult.Yes:
						design.Preserved = MsoTriState.msoFalse;
						flag = true;
						continue;
					case DialogResult.Cancel:
						presentation = null;
						application = null;
						return;
					case DialogResult.No:
						break;
					}
					break;
				}
				while (true)
				{
					switch (7)
					{
					case 0:
						break;
					default:
						goto end_IL_00f8;
					}
					continue;
					end_IL_00f8:
					break;
				}
				break;
			}
		}
		finally
		{
			if (enumerator is IDisposable)
			{
				while (true)
				{
					switch (1)
					{
					case 0:
						continue;
					}
					(enumerator as IDisposable).Dispose();
					break;
				}
			}
		}
		presentation.ApplyTemplate(strTemplatePath);
		if (AirplaneMode.IsOn())
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
			AirplaneMode.HidePresentationImages(presentation);
		}
		if (Behavior.GetPresentationFlysheetStyle(presentation) == FlySheetStyle.Agenda)
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
			Behavior.SetPresentationFlysheetStyle(presentation, FlySheetStyle.Agenda);
		}
		Templates.A(ref A, ref B, strTemplatePath, application);
		if (A != null)
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
			PageSetup pageSetup = A.PageSetup;
			bool flag2 = pageSetup.SlideHeight != presentation.PageSetup.SlideHeight;
			bool flag3 = pageSetup.SlideWidth != presentation.PageSetup.SlideWidth;
			_ = null;
			Templates.A(presentation, A);
			new Settings(A).A(presentation);
			_ = null;
			Create.ApplyHeadersFooters(presentation, A);
			if (B)
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
				PowerPointAddIn1.Presentation.Helpers.CloseQuietly(A);
			}
			A = null;
			if (!flag2)
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
				if (!flag3)
				{
					goto IL_0229;
				}
			}
			try
			{
				Images.FixDistortion(presentation.SlideMaster.CustomLayouts, flag3, flag2);
			}
			catch (Exception ex3)
			{
				ProjectData.SetProjectError(ex3);
				Exception ex4 = ex3;
				ProjectData.ClearProjectError();
			}
			goto IL_0229;
		}
		Forms.ErrorMessage(AH.A(68368));
		goto IL_024e;
		IL_0229:
		clsReporting.LogActivity((ActivityApp)2, (ActivityCategory)11, AH.A(68339));
		goto IL_024e;
		IL_024e:
		application = null;
		presentation = null;
	}

	private static void A(ref Microsoft.Office.Interop.PowerPoint.Presentation A, ref bool B, string C, Microsoft.Office.Interop.PowerPoint.Application D)
	{
		try
		{
			A = D.Presentations[Path.GetFileName(C)];
			if (Operators.CompareString(A.FullName, C, TextCompare: false) != 0)
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
					A = null;
					break;
				}
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		if (A != null)
		{
			return;
		}
		while (true)
		{
			switch (3)
			{
			case 0:
				continue;
			}
			try
			{
				A = PowerPointAddIn1.Presentation.Helpers.OpenQuietly(D, C);
				B = true;
				return;
			}
			catch (Exception ex3)
			{
				ProjectData.SetProjectError(ex3);
				Exception ex4 = ex3;
				ProjectData.ClearProjectError();
				return;
			}
		}
	}

	private static void A(Microsoft.Office.Interop.PowerPoint.Presentation A, string B)
	{
		A.Tags.Add(AH.A(68407), B);
	}

	public static string GetTemplateId(Microsoft.Office.Interop.PowerPoint.Presentation pres)
	{
		string result;
		try
		{
			result = pres.Tags[AH.A(68407)].ToString();
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			result = "";
			ProjectData.ClearProjectError();
		}
		return result;
	}

	private static void A(Microsoft.Office.Interop.PowerPoint.Presentation A, Microsoft.Office.Interop.PowerPoint.Presentation B)
	{
		Templates.A(A, B.FullName, C: false);
	}

	internal static void A(Microsoft.Office.Interop.PowerPoint.Presentation A, string B, bool C = false)
	{
		if (Ribbon.MenuTemplates.Count > 0)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
				{
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					string value = string.Empty;
					if (Ribbon.MenuTemplates.TryGetValue(B, out value))
					{
						Templates.A(A, value);
					}
					return;
				}
				}
			}
		}
		if (!C)
		{
			return;
		}
		while (true)
		{
			switch (2)
			{
			case 0:
				continue;
			}
			Templates.A(A, Create.DefaultTemplateId);
			return;
		}
	}

	internal static void B(ref Microsoft.Office.Interop.PowerPoint.Presentation A, ref bool B, string C, Microsoft.Office.Interop.PowerPoint.Application D)
	{
		string text = Templates.A(C);
		if (Operators.CompareString(text, string.Empty, TextCompare: false) == 0)
		{
			return;
		}
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
			if (text.Length <= 0)
			{
				return;
			}
			try
			{
				A = D.Presentations[Path.GetFileName(text)];
				if (Operators.CompareString(A.FullName, text, TextCompare: false) != 0)
				{
					while (true)
					{
						switch (5)
						{
						case 0:
							continue;
						}
						A = null;
						break;
					}
				}
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			if (A != null)
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
				try
				{
					A = PowerPointAddIn1.Presentation.Helpers.OpenQuietly(D, text);
					B = true;
					return;
				}
				catch (Exception ex3)
				{
					ProjectData.SetProjectError(ex3);
					Exception ex4 = ex3;
					ProjectData.ClearProjectError();
					return;
				}
			}
		}
	}

	internal static string A(string A)
	{
		string value;
		if (Operators.CompareString(A, string.Empty, TextCompare: false) != 0)
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
			if (A.Length != 0)
			{
				value = "";
				XmlNode xmlNode = null;
				if (TemplatesDictionary != null)
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
					if (TemplatesDictionary.TryGetValue(A, out value))
					{
						goto IL_025f;
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
				}
				using (List<LibraryItem>.Enumerator enumerator = Base.LibraryCollection.GetEnumerator())
				{
					do
					{
						IL_0239:
						if (enumerator.MoveNext())
						{
							LibraryItem current = enumerator.Current;
							if (current.IsExternallyManaged())
							{
								goto IL_0239;
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
							bool flag = false;
							string text = current.Location;
							if (clsCaching.UseCachedLibraries(text))
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
								text = clsCaching.CachedLibraryFolder(current);
								if (!Directory.Exists(text))
								{
									text = current.Location;
								}
								else
								{
									flag = true;
								}
							}
							if (!flag)
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
								if (!Directory.Exists(text))
								{
									continue;
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
							string[] directories = Directory.GetDirectories(Path.Combine(text, Base.LIB_MASTERS_FOLDER_NAME));
							int num = 0;
							while (true)
							{
								if (num < directories.Length)
								{
									string text2 = directories[num];
									string manifestPath = Manifests.GetManifestPath(text2);
									try
									{
										XmlDocument xmlDocument = new XmlDocument();
										xmlDocument.Load(manifestPath);
										xmlNode = xmlDocument.DocumentElement.SelectSingleNode(AH.A(68428) + A + AH.A(68449));
										if (xmlNode != null)
										{
											while (true)
											{
												switch (1)
												{
												case 0:
													continue;
												}
												value = Path.Combine(text2, xmlNode.Attributes[AH.A(63355)].Value);
												xmlNode = null;
												if (TemplatesDictionary == null)
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
													TemplatesDictionary = new Dictionary<string, string>();
												}
												TemplatesDictionary.Add(A, value);
												break;
											}
											break;
										}
									}
									catch (FileNotFoundException ex)
									{
										ProjectData.SetProjectError(ex);
										FileNotFoundException ex2 = ex;
										ProjectData.ClearProjectError();
									}
									catch (Exception ex3)
									{
										ProjectData.SetProjectError(ex3);
										Exception ex4 = ex3;
										clsReporting.LogException(ex4);
										ProjectData.ClearProjectError();
									}
									num = checked(num + 1);
									continue;
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
								break;
							}
							continue;
						}
						while (true)
						{
							switch (4)
							{
							case 0:
								break;
							default:
								goto end_IL_0245;
							}
							continue;
							end_IL_0245:
							break;
						}
						break;
					}
					while (Operators.CompareString(value, string.Empty, TextCompare: false) == 0 || value.Length <= 0);
				}
				goto IL_025f;
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
		return "";
		IL_025f:
		return value;
	}
}
