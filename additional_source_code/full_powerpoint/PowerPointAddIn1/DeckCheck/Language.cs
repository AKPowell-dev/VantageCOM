using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Windows.Forms;
using A;
using MacabacusMacros;
using MacabacusMacros.Auth;
using MacabacusMacros.UI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1.DeckCheck;

public sealed class Language
{
	public static string LanguagesMenu()
	{
		StringBuilder stringBuilder = new StringBuilder(AH.A(47526));
		List<int> list = new List<int>();
		IEnumerator enumerator = InputLanguage.InstalledInputLanguages.GetEnumerator();
		try
		{
			while (enumerator.MoveNext())
			{
				CultureInfo culture = ((InputLanguage)enumerator.Current).Culture;
				if (!list.Contains(culture.LCID))
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
					stringBuilder.Append(AH.A(47664) + culture.LCID + AH.A(47705) + culture.DisplayName + AH.A(47724) + culture.LCID + AH.A(47785));
					list.Add(culture.LCID);
				}
				culture = null;
			}
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					goto end_IL_0109;
				}
				continue;
				end_IL_0109:
				break;
			}
		}
		finally
		{
			IDisposable disposable = enumerator as IDisposable;
			if (disposable != null)
			{
				disposable.Dispose();
			}
		}
		list = null;
		stringBuilder.Append(AH.A(49007));
		return stringBuilder.ToString();
	}

	public static void SetProofingLanguage(IRibbonControl control)
	{
		if (!Access.AllowPowerPointOperation((PlanType)4, (Restriction)2, false))
		{
			return;
		}
		IEnumerator enumerator = default(IEnumerator);
		IEnumerator enumerator2 = default(IEnumerator);
		IEnumerator enumerator3 = default(IEnumerator);
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
			int b = Conversions.ToInteger(control.Tag);
			Microsoft.Office.Interop.PowerPoint.Application application = NG.A.Application;
			Microsoft.Office.Interop.PowerPoint.Presentation activePresentation = application.ActivePresentation;
			application.StartNewUndoEntry();
			try
			{
				try
				{
					enumerator = activePresentation.Slides.GetEnumerator();
					while (enumerator.MoveNext())
					{
						Slide obj = (Slide)enumerator.Current;
						A(obj.Shapes, b);
						A(obj.NotesPage.Shapes, b);
					}
					while (true)
					{
						switch (5)
						{
						case 0:
							break;
						default:
							goto end_IL_0097;
						}
						continue;
						end_IL_0097:
						break;
					}
				}
				finally
				{
					if (enumerator is IDisposable)
					{
						while (true)
						{
							switch (5)
							{
							case 0:
								continue;
							}
							(enumerator as IDisposable).Dispose();
							break;
						}
					}
				}
				try
				{
					enumerator2 = activePresentation.Designs.GetEnumerator();
					while (enumerator2.MoveNext())
					{
						Design design = (Design)enumerator2.Current;
						A(design.SlideMaster.Shapes, b);
						try
						{
							enumerator3 = design.SlideMaster.CustomLayouts.GetEnumerator();
							while (enumerator3.MoveNext())
							{
								A(((CustomLayout)enumerator3.Current).Shapes, b);
							}
							while (true)
							{
								switch (5)
								{
								case 0:
									break;
								default:
									goto end_IL_0137;
								}
								continue;
								end_IL_0137:
								break;
							}
						}
						finally
						{
							if (enumerator3 is IDisposable)
							{
								while (true)
								{
									switch (5)
									{
									case 0:
										continue;
									}
									(enumerator3 as IDisposable).Dispose();
									break;
								}
							}
						}
					}
					while (true)
					{
						switch (7)
						{
						case 0:
							break;
						default:
							goto end_IL_0170;
						}
						continue;
						end_IL_0170:
						break;
					}
				}
				finally
				{
					if (enumerator2 is IDisposable)
					{
						while (true)
						{
							switch (4)
							{
							case 0:
								continue;
							}
							(enumerator2 as IDisposable).Dispose();
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
			clsReporting.LogActivity((ActivityApp)2, (ActivityCategory)12, AH.A(49022));
			activePresentation = null;
			return;
		}
	}

	private static void A(Microsoft.Office.Interop.PowerPoint.Shapes A, int B)
	{
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = A.GetEnumerator();
			while (enumerator.MoveNext())
			{
				Language.A((Microsoft.Office.Interop.PowerPoint.Shape)enumerator.Current, B);
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
				return;
			}
		}
		finally
		{
			if (enumerator is IDisposable)
			{
				while (true)
				{
					switch (5)
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

	private static void A(Microsoft.Office.Interop.PowerPoint.Shape A, int B)
	{
		try
		{
			if (A.HasTextFrame == MsoTriState.msoTrue)
			{
				while (true)
				{
					switch (3)
					{
					case 0:
						break;
					default:
						if (1 == 0)
						{
							/*OpCode not supported: LdMemberToken*/;
						}
						A.TextFrame2.TextRange.LanguageID = (MsoLanguageID)B;
						return;
					}
				}
			}
			if (A.HasTable == MsoTriState.msoTrue)
			{
				while (true)
				{
					switch (3)
					{
					case 0:
						break;
					default:
					{
						Table table = A.Table;
						int count = table.Rows.Count;
						int count2 = table.Columns.Count;
						int num = count;
						for (int i = 1; i <= num; i = checked(i + 1))
						{
							int num2 = count2;
							for (int j = 1; j <= num2; j = checked(j + 1))
							{
								Cell cell = table.Cell(i, j);
								if (cell.Shape.HasTextFrame == MsoTriState.msoTrue)
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
									cell.Shape.TextFrame2.TextRange.LanguageID = (MsoLanguageID)B;
								}
								cell = null;
							}
							while (true)
							{
								switch (2)
								{
								case 0:
									break;
								default:
									goto end_IL_00cf;
								}
								continue;
								end_IL_00cf:
								break;
							}
						}
						table = null;
						return;
					}
					}
				}
			}
			IEnumerator enumerator = default(IEnumerator);
			if (A.HasSmartArt == MsoTriState.msoTrue)
			{
				while (true)
				{
					switch (4)
					{
					case 0:
						break;
					default:
						try
						{
							enumerator = A.SmartArt.AllNodes.GetEnumerator();
							while (enumerator.MoveNext())
							{
								((SmartArtNode)enumerator.Current).TextFrame2.TextRange.LanguageID = (MsoLanguageID)B;
							}
							return;
						}
						finally
						{
							if (enumerator is IDisposable)
							{
								while (true)
								{
									switch (2)
									{
									case 0:
										break;
									default:
										(enumerator as IDisposable).Dispose();
										goto end_IL_014a;
									}
									continue;
									end_IL_014a:
									break;
								}
							}
						}
					}
				}
			}
			if (A.Type != MsoShapeType.msoGroup)
			{
				return;
			}
			IEnumerator enumerator2 = default(IEnumerator);
			while (true)
			{
				switch (4)
				{
				case 0:
					continue;
				}
				try
				{
					enumerator2 = A.GroupItems.GetEnumerator();
					while (enumerator2.MoveNext())
					{
						Language.A((Microsoft.Office.Interop.PowerPoint.Shape)enumerator2.Current, B);
					}
					while (true)
					{
						switch (7)
						{
						case 0:
							break;
						default:
							return;
						}
					}
				}
				finally
				{
					if (enumerator2 is IDisposable)
					{
						while (true)
						{
							switch (4)
							{
							case 0:
								continue;
							}
							(enumerator2 as IDisposable).Dispose();
							break;
						}
					}
				}
			}
		}
		catch (NotImplementedException ex)
		{
			ProjectData.SetProjectError(ex);
			NotImplementedException ex2 = ex;
			ProjectData.ClearProjectError();
		}
		catch (Exception ex3)
		{
			ProjectData.SetProjectError(ex3);
			Exception ex4 = ex3;
			clsReporting.LogException(ex4);
			ProjectData.ClearProjectError();
		}
	}
}
