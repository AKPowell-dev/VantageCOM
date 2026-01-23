using System;
using System.Collections;
using System.Collections.Generic;
using A;
using MacabacusMacros.Auth;
using MacabacusMacros.UI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1.Links;

public sealed class Ribbon
{
	public enum LinkSelection
	{
		Unknown,
		No,
		Yes
	}

	private enum BF
	{
		A,
		B,
		C,
		D
	}

	private static LinkSelection m_A = LinkSelection.Unknown;

	public static LinkSelection LinkSelected
	{
		get
		{
			return Ribbon.m_A;
		}
		set
		{
			Ribbon.m_A = value;
		}
	}

	public static void ResetSelectionType()
	{
		LinkSelected = LinkSelection.Unknown;
	}

	public static bool IsLinkSelected()
	{
		//IL_0007: Unknown result type (might be due to invalid IL or missing references)
		//IL_000c: Unknown result type (might be due to invalid IL or missing references)
		//IL_000f: Invalid comparison between Unknown and I4
		if ((int)Base.UserProfile.LicenseType != 2)
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
			LinkSelected = LinkSelection.No;
		}
		if (LinkSelected == LinkSelection.Unknown)
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
			LinkSelected = LinkSelection.No;
			Selection selection;
			try
			{
				selection = NG.A.Application.ActiveWindow.Selection;
				PpSelectionType type = selection.Type;
				if (type != PpSelectionType.ppSelectionShapes)
				{
					while (true)
					{
						switch (7)
						{
						case 0:
							continue;
						}
						if (type != PpSelectionType.ppSelectionText)
						{
							break;
						}
						A(selection);
						if (!KG.A.TextLinkCompatibilityMode || LinkSelected != LinkSelection.No)
						{
							break;
						}
						while (true)
						{
							switch (1)
							{
							case 0:
								continue;
							}
							if (!Text.LinkSelected(selection))
							{
								break;
							}
							while (true)
							{
								switch (5)
								{
								case 0:
									continue;
								}
								LinkSelected = LinkSelection.Yes;
								break;
							}
							break;
						}
						break;
					}
				}
				else
				{
					if (!KG.A.TextLinkCompatibilityMode)
					{
						if (selection.HasChildShapeRange)
						{
							IEnumerator enumerator = default(IEnumerator);
							try
							{
								enumerator = selection.ChildShapeRange.GetEnumerator();
								do
								{
									if (enumerator.MoveNext())
									{
										continue;
									}
									while (true)
									{
										switch (1)
										{
										case 0:
											break;
										default:
											goto end_IL_0112;
										}
										continue;
										end_IL_0112:
										break;
									}
									break;
								}
								while (!A((Microsoft.Office.Interop.PowerPoint.Shape)enumerator.Current));
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
						}
						else
						{
							IEnumerator enumerator2 = default(IEnumerator);
							try
							{
								enumerator2 = selection.ShapeRange.GetEnumerator();
								while (enumerator2.MoveNext() && !A((Microsoft.Office.Interop.PowerPoint.Shape)enumerator2.Current))
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
								}
							}
							finally
							{
								if (enumerator2 is IDisposable)
								{
									while (true)
									{
										switch (1)
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
					else
					{
						IEnumerator enumerator3 = default(IEnumerator);
						try
						{
							enumerator3 = selection.ShapeRange.GetEnumerator();
							while (true)
							{
								if (enumerator3.MoveNext())
								{
									Microsoft.Office.Interop.PowerPoint.Shape shp = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator3.Current;
									if (!Shapes.IsLinked(shp))
									{
										while (true)
										{
											switch (6)
											{
											case 0:
												break;
											default:
												goto end_IL_01d1;
											}
											continue;
											end_IL_01d1:
											break;
										}
										if (!Text.ContainsLinks(shp))
										{
											continue;
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
									LinkSelected = LinkSelection.Yes;
									break;
								}
								while (true)
								{
									switch (2)
									{
									case 0:
										break;
									default:
										goto end_IL_0203;
									}
									continue;
									end_IL_0203:
									break;
								}
								break;
							}
						}
						finally
						{
							if (enumerator3 is IDisposable)
							{
								while (true)
								{
									switch (1)
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
					if (LinkSelected == LinkSelection.No)
					{
						A(selection);
					}
				}
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				LinkSelected = LinkSelection.No;
				ProjectData.ClearProjectError();
			}
			selection = null;
		}
		if (LinkSelected == LinkSelection.No)
		{
			return false;
		}
		return true;
	}

	private static bool A(Microsoft.Office.Interop.PowerPoint.Shape A)
	{
		if (A.Type != MsoShapeType.msoGroup)
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
					if (Shapes.IsLinked(A))
					{
						while (true)
						{
							switch (2)
							{
							case 0:
								break;
							default:
								LinkSelected = LinkSelection.Yes;
								return true;
							}
						}
					}
					return false;
				}
			}
		}
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = A.GroupItems.GetEnumerator();
			while (enumerator.MoveNext())
			{
				Ribbon.A((Microsoft.Office.Interop.PowerPoint.Shape)enumerator.Current);
			}
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					goto end_IL_006c;
				}
				continue;
				end_IL_006c:
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
		bool result = default(bool);
		return result;
	}

	private static void A(Selection A)
	{
		Slide slide = A.SlideRange[1];
		if (slide.Hyperlinks.Count > 0)
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			List<Microsoft.Office.Interop.PowerPoint.Shape> listShapes = new List<Microsoft.Office.Interop.PowerPoint.Shape>();
			if (A.HasChildShapeRange)
			{
				IEnumerator enumerator = default(IEnumerator);
				try
				{
					enumerator = A.ChildShapeRange.GetEnumerator();
					while (enumerator.MoveNext())
					{
						Hyperlinks.ProcessSelectedShape((Microsoft.Office.Interop.PowerPoint.Shape)enumerator.Current, ref listShapes);
					}
					while (true)
					{
						switch (2)
						{
						case 0:
							break;
						default:
							goto end_IL_007c;
						}
						continue;
						end_IL_007c:
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
			}
			else
			{
				IEnumerator enumerator2 = default(IEnumerator);
				try
				{
					enumerator2 = A.ShapeRange.GetEnumerator();
					while (enumerator2.MoveNext())
					{
						Hyperlinks.ProcessSelectedShape((Microsoft.Office.Interop.PowerPoint.Shape)enumerator2.Current, ref listShapes);
					}
					while (true)
					{
						switch (6)
						{
						case 0:
							break;
						default:
							goto end_IL_00d4;
						}
						continue;
						end_IL_00d4:
						break;
					}
				}
				finally
				{
					if (enumerator2 is IDisposable)
					{
						while (true)
						{
							switch (1)
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
			IEnumerator enumerator3 = default(IEnumerator);
			try
			{
				enumerator3 = slide.Hyperlinks.GetEnumerator();
				while (enumerator3.MoveNext())
				{
					Hyperlink hyp = (Hyperlink)enumerator3.Current;
					if (!Hyperlinks.IsLinked(hyp))
					{
						continue;
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
					if (!Hyperlinks.SelectedShapesContainHyperlink(hyp, listShapes))
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
					if (A.Type != PpSelectionType.ppSelectionShapes)
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
						if (A.Type != PpSelectionType.ppSelectionText || !Hyperlinks.IsHyperlinkSelected(hyp, A))
						{
							continue;
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
					LinkSelected = LinkSelection.Yes;
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
			listShapes = null;
		}
		slide = null;
	}

	public static void RefreshLinks()
	{
		A(BF.A);
	}

	public static void EditLinks()
	{
		A(BF.B);
	}

	public static void ViewSource()
	{
		A(BF.C);
	}

	public static void BreakLinks()
	{
		A(BF.D);
	}

	private static void A(BF A)
	{
		Selection selection;
		try
		{
			Application application = NG.A.Application;
			Microsoft.Office.Interop.PowerPoint.Presentation activePresentation = application.ActivePresentation;
			selection = application.ActiveWindow.Selection;
			application = null;
			if (activePresentation.Final)
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
					Forms.WarningMessage(Common.A);
					selection = null;
					return;
				}
			}
			switch (selection.Type)
			{
			case PpSelectionType.ppSelectionText:
				switch (A)
				{
				case BF.A:
					Text.RefreshLinks(selection);
					break;
				case BF.C:
					Text.ViewSource(selection);
					break;
				case BF.B:
					Text.EditLinks(selection);
					break;
				case BF.D:
					Text.BreakLinks(selection, blnUpdateRibbon: true);
					break;
				}
				break;
			case PpSelectionType.ppSelectionShapes:
				switch (A)
				{
				case BF.A:
					Shapes.RefreshLinks(selection);
					break;
				case BF.C:
					Shapes.ViewSource(selection);
					break;
				case BF.B:
					Shapes.EditLinks(selection);
					break;
				case BF.D:
					Shapes.BreakLinks(selection);
					break;
				}
				break;
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		selection = null;
	}
}
