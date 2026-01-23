using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using A;
using MacabacusMacros;
using MacabacusMacros.Auth;
using MacabacusMacros.UI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1.Shapes;

public sealed class Styles
{
	private static readonly string m_A = AH.A(91675);

	private static readonly string m_B = AH.A(91698);

	private static Dictionary<string, Microsoft.Office.Interop.PowerPoint.Shape> m_A = null;

	private static Dictionary<string, Microsoft.Office.Interop.PowerPoint.Shape> A
	{
		get
		{
			return Styles.m_A;
		}
		set
		{
			Styles.m_A = value;
		}
	}

	public static string Menu()
	{
		Microsoft.Office.Interop.PowerPoint.Application application = NG.A.Application;
		StringBuilder stringBuilder = new StringBuilder(AH.A(47526));
		List<string> list;
		if (application.Windows.Count > 0)
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
			A();
			list = new List<string>();
			list.Add(AH.A(7953));
			list.Add(AH.A(7941));
			list.Add(AH.A(7914));
			using (Dictionary<string, Microsoft.Office.Interop.PowerPoint.Shape>.Enumerator enumerator = Styles.A.GetEnumerator())
			{
				while (enumerator.MoveNext())
				{
					KeyValuePair<string, Microsoft.Office.Interop.PowerPoint.Shape> current = enumerator.Current;
					string text = Regex.Replace(current.Value.Name, Styles.m_B, "").Trim();
					string text2 = clsRibbon.GenerateLabel(text, list);
					text = clsRibbon.FixAmpersand(text);
					if (current.Value.HasTable != MsoTriState.msoTrue)
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
						stringBuilder.Append(AH.A(87565) + current.Key + AH.A(47705) + text2 + AH.A(87612) + current.Key + AH.A(87673) + text + AH.A(87728));
					}
					else
					{
						stringBuilder.Append(AH.A(87565) + current.Key + AH.A(47705) + text2 + AH.A(87612) + current.Key + AH.A(87673) + text + AH.A(87876));
					}
				}
				while (true)
				{
					switch (3)
					{
					case 0:
						break;
					default:
						goto end_IL_01fe;
					}
					continue;
					end_IL_01fe:
					break;
				}
			}
			if (Styles.A.Count > 0)
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
				stringBuilder.Append(AH.A(88042));
				stringBuilder.Append(AH.A(88119));
			}
			else
			{
				Forms.WarningMessage(AH.A(89171));
			}
			stringBuilder.Append(AH.A(89301));
			string text3;
			if (Base.IsUserAdmin())
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
				text3 = AH.A(89410);
			}
			else
			{
				text3 = AH.A(89419);
			}
			stringBuilder.Append(AH.A(89430) + text3 + AH.A(82681));
			stringBuilder.Append(AH.A(90178) + text3 + AH.A(82681));
		}
		list = null;
		stringBuilder.Append(AH.A(49007));
		return stringBuilder.ToString();
	}

	public static void Apply(IRibbonControl control)
	{
		if (!Licensing.AllowStylesOperation())
		{
			return;
		}
		IEnumerator enumerator = default(IEnumerator);
		IEnumerator enumerator2 = default(IEnumerator);
		IEnumerator enumerator3 = default(IEnumerator);
		IEnumerator enumerator4 = default(IEnumerator);
		while (true)
		{
			switch (2)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			Microsoft.Office.Interop.PowerPoint.Shape value = null;
			string tag = control.Tag;
			if (!Styles.A.TryGetValue(tag, out value))
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
				Microsoft.Office.Interop.PowerPoint.Application application;
				Selection selection;
				try
				{
					application = NG.A.Application;
					selection = application.ActiveWindow.Selection;
					if (selection.Type == PpSelectionType.ppSelectionShapes)
					{
						goto IL_008c;
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
					if (selection.Type == PpSelectionType.ppSelectionText)
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
						goto IL_008c;
					}
					goto end_IL_0048;
					IL_008c:
					application.StartNewUndoEntry();
					if (value.HasTable == MsoTriState.msoTrue)
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
						if (selection.HasChildShapeRange)
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
							try
							{
								enumerator = selection.ChildShapeRange.GetEnumerator();
								while (enumerator.MoveNext())
								{
									A((Microsoft.Office.Interop.PowerPoint.Shape)enumerator.Current, value, tag);
								}
								while (true)
								{
									switch (6)
									{
									case 0:
										break;
									default:
										goto end_IL_00ef;
									}
									continue;
									end_IL_00ef:
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
						}
						else
						{
							try
							{
								enumerator2 = selection.ShapeRange.GetEnumerator();
								while (enumerator2.MoveNext())
								{
									A((Microsoft.Office.Interop.PowerPoint.Shape)enumerator2.Current, value, tag);
								}
								while (true)
								{
									switch (7)
									{
									case 0:
										break;
									default:
										goto end_IL_014e;
									}
									continue;
									end_IL_014e:
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
					}
					else
					{
						value.PickUp();
						if (selection.HasChildShapeRange)
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
							try
							{
								enumerator3 = selection.ChildShapeRange.GetEnumerator();
								while (enumerator3.MoveNext())
								{
									A((Microsoft.Office.Interop.PowerPoint.Shape)enumerator3.Current, tag);
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
						else
						{
							try
							{
								enumerator4 = selection.ShapeRange.GetEnumerator();
								while (enumerator4.MoveNext())
								{
									A((Microsoft.Office.Interop.PowerPoint.Shape)enumerator4.Current, tag);
								}
							}
							finally
							{
								if (enumerator4 is IDisposable)
								{
									while (true)
									{
										switch (7)
										{
										case 0:
											continue;
										}
										(enumerator4 as IDisposable).Dispose();
										break;
									}
								}
							}
						}
					}
					B(AH.A(90920));
					end_IL_0048:;
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception a = ex;
					A(a);
					ProjectData.ClearProjectError();
				}
				application = null;
				selection = null;
				return;
			}
		}
	}

	public static void Reset()
	{
		Microsoft.Office.Interop.PowerPoint.Application application = NG.A.Application;
		try
		{
			if (application.Presentations.Count > 0)
			{
				IEnumerator enumerator4 = default(IEnumerator);
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
					Selection selection = application.ActiveWindow.Selection;
					application.StartNewUndoEntry();
					try
					{
						PpSelectionType type = selection.Type;
						if (type == PpSelectionType.ppSelectionSlides)
						{
							foreach (Slide slide in application.ActivePresentation.Slides)
							{
								foreach (Microsoft.Office.Interop.PowerPoint.Shape shape in slide.Shapes)
								{
									A(shape);
								}
							}
						}
						else
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
							if ((uint)(type - 2) <= 1u)
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
								if (!selection.HasChildShapeRange)
								{
									foreach (Microsoft.Office.Interop.PowerPoint.Shape item in selection.ShapeRange)
									{
										A(item);
									}
								}
								else
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
										enumerator4 = selection.ChildShapeRange.GetEnumerator();
										while (enumerator4.MoveNext())
										{
											A((Microsoft.Office.Interop.PowerPoint.Shape)enumerator4.Current);
										}
									}
									finally
									{
										if (enumerator4 is IDisposable)
										{
											while (true)
											{
												switch (2)
												{
												case 0:
													continue;
												}
												(enumerator4 as IDisposable).Dispose();
												break;
											}
										}
									}
								}
							}
							else
							{
								Forms.WarningMessage(AH.A(13552));
							}
						}
						B(AH.A(90943));
					}
					catch (Exception ex)
					{
						ProjectData.SetProjectError(ex);
						Exception a = ex;
						A(a);
						ProjectData.ClearProjectError();
					}
					break;
				}
			}
		}
		catch (Exception ex2)
		{
			ProjectData.SetProjectError(ex2);
			Exception ex3 = ex2;
			ProjectData.ClearProjectError();
		}
		application = null;
	}

	public static void Remove()
	{
		Microsoft.Office.Interop.PowerPoint.Application application = NG.A.Application;
		int num = 0;
		checked
		{
			Master master;
			try
			{
				master = A(application);
				application.StartNewUndoEntry();
				try
				{
					for (int i = master.Shapes.Count; i >= 1; i += -1)
					{
						if (!A(master.Shapes[i]))
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
						if (1 == 0)
						{
							/*OpCode not supported: LdMemberToken*/;
						}
						master.Shapes[i].Delete();
						num++;
					}
					while (true)
					{
						switch (3)
						{
						case 0:
							continue;
						}
						if (num > 1)
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
							Forms.SuccessMessage(num + AH.A(90966));
						}
						else
						{
							Forms.InfoMessage(AH.A(91003));
						}
						B(AH.A(91036));
						break;
					}
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception a = ex;
					A(a);
					ProjectData.ClearProjectError();
				}
			}
			catch (Exception ex2)
			{
				ProjectData.SetProjectError(ex2);
				Exception ex3 = ex2;
				ProjectData.ClearProjectError();
			}
			application = null;
			master = null;
		}
	}

	public static void Create()
	{
		if (!Licensing.AllowStylesOperation())
		{
			return;
		}
		IEnumerator enumerator = default(IEnumerator);
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
			Microsoft.Office.Interop.PowerPoint.Application application = NG.A.Application;
			bool flag = false;
			Selection selection;
			Master master;
			try
			{
				selection = application.ActiveWindow.Selection;
				if (selection.Type == PpSelectionType.ppSelectionShapes)
				{
					goto IL_006c;
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
				if (selection.Type == PpSelectionType.ppSelectionText)
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
					goto IL_006c;
				}
				Helpers.SingleShapeRequiredError();
				goto end_IL_0031;
				IL_00bd:
				Microsoft.Office.Interop.PowerPoint.Shape shape;
				string text = Forms.InputBox(AH.A(91063), AH.A(91082), shape.Name);
				if (Operators.CompareString(text, string.Empty, TextCompare: false) != 0)
				{
					while (true)
					{
						switch (7)
						{
						case 0:
							continue;
						}
						if (text.Length <= 0)
						{
							break;
						}
						while (true)
						{
							switch (3)
							{
							case 0:
								continue;
							}
							master = A(application);
							text += AH.A(91113);
							enumerator = master.Shapes.GetEnumerator();
							try
							{
								while (true)
								{
									if (enumerator.MoveNext())
									{
										if (Operators.CompareString(((Microsoft.Office.Interop.PowerPoint.Shape)enumerator.Current).Name, text, TextCompare: false) != 0)
										{
											continue;
										}
										while (true)
										{
											switch (2)
											{
											case 0:
												continue;
											}
											Forms.WarningMessage(AH.A(91126));
											flag = true;
											break;
										}
										break;
									}
									while (true)
									{
										switch (2)
										{
										case 0:
											break;
										default:
											goto end_IL_019a;
										}
										continue;
										end_IL_019a:
										break;
									}
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
							if (!flag)
							{
								application.StartNewUndoEntry();
								try
								{
									shape.Copy();
									Microsoft.Office.Interop.PowerPoint.Shape shape2 = master.Shapes.Paste()[1];
									shape2.Visible = MsoTriState.msoFalse;
									shape2.Name = text;
									shape2.Top = shape.Top;
									shape2.Left = shape.Left;
									_ = null;
									Forms.SuccessMessage(AH.A(91256));
									B(AH.A(91063));
								}
								catch (Exception ex)
								{
									ProjectData.SetProjectError(ex);
									Exception a = ex;
									A(a);
									ProjectData.ClearProjectError();
								}
							}
							break;
						}
						break;
					}
				}
				goto end_IL_0031;
				IL_006c:
				if (selection.ShapeRange.Count == 1)
				{
					shape = selection.ShapeRange[1];
					MsoShapeType type = shape.Type;
					if (type == MsoShapeType.msoAutoShape || type == MsoShapeType.msoPicture)
					{
						goto IL_00bd;
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
					if (type == MsoShapeType.msoTextBox)
					{
						goto IL_00bd;
					}
					Forms.WarningMessage(AH.A(91293));
				}
				else
				{
					Helpers.SingleShapeRequiredError();
				}
				end_IL_0031:;
			}
			catch (Exception ex2)
			{
				ProjectData.SetProjectError(ex2);
				Exception ex3 = ex2;
				ProjectData.ClearProjectError();
			}
			application = null;
			selection = null;
			master = null;
			return;
		}
	}

	public static void Edit()
	{
		if (!Licensing.AllowStylesOperation())
		{
			return;
		}
		Microsoft.Office.Interop.PowerPoint.Application application = NG.A.Application;
		List<string> list = new List<string>();
		try
		{
			Master master = A(application);
			application.ActiveWindow.ViewType = PpViewType.ppViewMasterThumbnails;
			Microsoft.Office.Interop.PowerPoint.Shapes shapes = master.Shapes;
			IEnumerator enumerator = shapes.GetEnumerator();
			try
			{
				while (enumerator.MoveNext())
				{
					Microsoft.Office.Interop.PowerPoint.Shape shape = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator.Current;
					if (HasStyleShapeName(shape))
					{
						shape.Visible = MsoTriState.msoTrue;
						list.Add(shape.Name);
					}
				}
				while (true)
				{
					switch (2)
					{
					case 0:
						continue;
					}
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
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
			if (list.Count > 0)
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
				shapes.Range(list.ToArray()).Select();
			}
			if (!application.CommandBars.GetPressedMso(AH.A(91479)))
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
				application.CommandBars.ExecuteMso(AH.A(91479));
				System.Windows.Forms.Application.DoEvents();
			}
			Forms.WarningMessage(AH.A(91506));
			B(AH.A(91652));
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception a = ex;
			A(a);
			ProjectData.ClearProjectError();
		}
		application = null;
		list = null;
	}

	private static void A()
	{
		Styles.A = new Dictionary<string, Microsoft.Office.Interop.PowerPoint.Shape>();
		Microsoft.Office.Interop.PowerPoint.Shapes shapes = NG.A.Application.ActivePresentation.Designs[1].SlideMaster.Shapes;
		IEnumerator enumerator = shapes.GetEnumerator();
		try
		{
			while (enumerator.MoveNext())
			{
				Microsoft.Office.Interop.PowerPoint.Shape shape = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator.Current;
				if (A(shape))
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
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					Styles.A.Add(shape.Id.ToString(), shape);
				}
				_ = null;
			}
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					goto end_IL_009d;
				}
				continue;
				end_IL_009d:
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
		shapes = null;
	}

	private static void A(Microsoft.Office.Interop.PowerPoint.Shape A, string B)
	{
		if (A.Type != MsoShapeType.msoGroup)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					A.Apply();
					A.Tags.Add(Styles.m_A, B);
					return;
				}
			}
		}
		IEnumerator enumerator = A.GroupItems.GetEnumerator();
		try
		{
			while (enumerator.MoveNext())
			{
				_ = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator.Current;
				Styles.A(A, B);
			}
			while (true)
			{
				switch (3)
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
			IDisposable disposable = enumerator as IDisposable;
			if (disposable != null)
			{
				disposable.Dispose();
			}
		}
	}

	private static void A(Microsoft.Office.Interop.PowerPoint.Shape A, Microsoft.Office.Interop.PowerPoint.Shape B, string C)
	{
		if (A.Type != MsoShapeType.msoGroup)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					if (A.HasTable == MsoTriState.msoTrue)
					{
						Styles.B(B, A);
						A.Tags.Add(Styles.m_A, C);
					}
					return;
				}
			}
		}
		IEnumerator enumerator = A.GroupItems.GetEnumerator();
		try
		{
			while (enumerator.MoveNext())
			{
				_ = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator.Current;
				Styles.A(A, B, C);
			}
			while (true)
			{
				switch (5)
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
			IDisposable disposable = enumerator as IDisposable;
			if (disposable != null)
			{
				disposable.Dispose();
			}
		}
	}

	private static void A(Microsoft.Office.Interop.PowerPoint.Shape A)
	{
		if (A.Type != MsoShapeType.msoGroup)
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
					B(A);
					return;
				}
			}
		}
		IEnumerator enumerator = A.GroupItems.GetEnumerator();
		try
		{
			while (enumerator.MoveNext())
			{
				Styles.A((Microsoft.Office.Interop.PowerPoint.Shape)enumerator.Current);
			}
			while (true)
			{
				switch (4)
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
			IDisposable disposable = enumerator as IDisposable;
			if (disposable != null)
			{
				disposable.Dispose();
			}
		}
	}

	private static void B(Microsoft.Office.Interop.PowerPoint.Shape A)
	{
		Microsoft.Office.Interop.PowerPoint.Shape value = null;
		if (!Styles.A.TryGetValue(A.Tags[Styles.m_A], out value))
		{
			return;
		}
		if (value.HasTable == MsoTriState.msoTrue)
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
			B(value, A);
		}
		else
		{
			value.PickUp();
			A.Apply();
		}
		value = null;
	}

	private static Master A(Microsoft.Office.Interop.PowerPoint.Application A)
	{
		return A.ActivePresentation.Designs[1].SlideMaster;
	}

	private static bool A(Microsoft.Office.Interop.PowerPoint.Shape A)
	{
		if (A.Visible == MsoTriState.msoFalse)
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
					return HasStyleShapeName(A);
				}
			}
		}
		return false;
	}

	public static bool HasStyleShapeName(Microsoft.Office.Interop.PowerPoint.Shape shp)
	{
		return Regex.IsMatch(shp.Name, Styles.m_B);
	}

	private static void A(Exception A)
	{
		Forms.ErrorMessage(A.Message);
		clsReporting.LogException(A);
	}

	private static void B(string A)
	{
		clsReporting.LogActivity((ActivityApp)2, (ActivityCategory)1, A);
	}

	private static void B(Microsoft.Office.Interop.PowerPoint.Shape A, Microsoft.Office.Interop.PowerPoint.Shape B)
	{
		Table table = A.Table;
		bool num = Styles.A(table);
		Table table2 = B.Table;
		table2.ApplyStyle(table.Style.Id);
		Microsoft.Office.Interop.PowerPoint.ColorFormat foreColor;
		if (num)
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
			table2.FirstCol = table.FirstCol;
			table2.FirstRow = table.FirstRow;
			table2.LastCol = table.LastCol;
			table2.LastRow = table.LastRow;
			table2.HorizBanding = table.HorizBanding;
			table2.VertBanding = table.VertBanding;
			if (table.Background.Fill.Visible == MsoTriState.msoTrue)
			{
				foreColor = table.Background.Fill.ForeColor;
				switch (foreColor.Type)
				{
				case MsoColorType.msoColorTypeRGB:
					table2.Background.Fill.ForeColor.RGB = foreColor.RGB;
					break;
				case MsoColorType.msoColorTypeScheme:
				{
					Microsoft.Office.Interop.PowerPoint.ColorFormat foreColor2 = table2.Background.Fill.ForeColor;
					foreColor2.ObjectThemeColor = foreColor.ObjectThemeColor;
					foreColor2.TintAndShade = foreColor.TintAndShade;
					foreColor2.Brightness = foreColor.Brightness;
					_ = null;
					break;
				}
				}
			}
			if (table.FirstCol && table2.FirstCol)
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
				foreColor = null;
				Table table3 = table;
				if (table3.FirstRow)
				{
					Microsoft.Office.Interop.PowerPoint.FillFormat fill = table3.Cell(2, 1).Shape.Fill;
					if (fill.Visible == MsoTriState.msoTrue)
					{
						foreColor = fill.ForeColor;
					}
					fill = null;
				}
				else
				{
					Microsoft.Office.Interop.PowerPoint.FillFormat fill2 = table3.Cell(1, 1).Shape.Fill;
					if (fill2.Visible == MsoTriState.msoTrue)
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
						foreColor = fill2.ForeColor;
					}
					fill2 = null;
				}
				table3 = null;
				if (foreColor != null)
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
					if (foreColor.Type == MsoColorType.msoColorTypeScheme)
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
						foreach (Cell cell in table2.Columns[1].Cells)
						{
							Styles.A(cell, foreColor);
						}
					}
					else
					{
						string value = Conversions.ToString(foreColor.RGB);
						IEnumerator enumerator2 = default(IEnumerator);
						try
						{
							enumerator2 = table2.Columns[1].Cells.GetEnumerator();
							while (enumerator2.MoveNext())
							{
								((Cell)enumerator2.Current).Shape.Fill.ForeColor.RGB = Conversions.ToInteger(value);
							}
							while (true)
							{
								switch (2)
								{
								case 0:
									break;
								default:
									goto end_IL_02e9;
								}
								continue;
								end_IL_02e9:
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
				}
			}
			if (table.LastCol && table2.LastCol)
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
				foreColor = null;
				Table table4 = table;
				if (table4.FirstRow)
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
					Microsoft.Office.Interop.PowerPoint.FillFormat fill3 = table4.Cell(2, table4.Columns.Count).Shape.Fill;
					if (fill3.Visible == MsoTriState.msoTrue)
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
						foreColor = fill3.ForeColor;
					}
					fill3 = null;
				}
				else
				{
					Microsoft.Office.Interop.PowerPoint.FillFormat fill4 = table4.Cell(1, table4.Columns.Count).Shape.Fill;
					if (fill4.Visible == MsoTriState.msoTrue)
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
						foreColor = fill4.ForeColor;
					}
					fill4 = null;
				}
				table4 = null;
				if (foreColor != null)
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
					int count = table2.Columns.Count;
					if (foreColor.Type == MsoColorType.msoColorTypeScheme)
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
						IEnumerator enumerator3 = default(IEnumerator);
						try
						{
							enumerator3 = table2.Columns[count].Cells.GetEnumerator();
							while (enumerator3.MoveNext())
							{
								Styles.A((Cell)enumerator3.Current, foreColor);
							}
							while (true)
							{
								switch (3)
								{
								case 0:
									break;
								default:
									goto end_IL_0464;
								}
								continue;
								end_IL_0464:
								break;
							}
						}
						finally
						{
							if (enumerator3 is IDisposable)
							{
								while (true)
								{
									switch (7)
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
					else
					{
						string value = Conversions.ToString(foreColor.RGB);
						{
							IEnumerator enumerator4 = table2.Columns[count].Cells.GetEnumerator();
							try
							{
								while (enumerator4.MoveNext())
								{
									((Cell)enumerator4.Current).Shape.Fill.ForeColor.RGB = Conversions.ToInteger(value);
								}
								while (true)
								{
									switch (7)
									{
									case 0:
										break;
									default:
										goto end_IL_04f8;
									}
									continue;
									end_IL_04f8:
									break;
								}
							}
							finally
							{
								IDisposable disposable = enumerator4 as IDisposable;
								if (disposable != null)
								{
									disposable.Dispose();
								}
							}
						}
					}
				}
			}
			if (table.FirstRow)
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
				if (table2.FirstRow)
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
					if (table.Cell(1, 1).Shape.Fill.Visible == MsoTriState.msoTrue)
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
						foreColor = table.Cell(1, 1).Shape.Fill.ForeColor;
						if (foreColor.Type == MsoColorType.msoColorTypeScheme)
						{
							IEnumerator enumerator5 = default(IEnumerator);
							try
							{
								enumerator5 = table2.Rows[1].Cells.GetEnumerator();
								while (enumerator5.MoveNext())
								{
									Styles.A((Cell)enumerator5.Current, foreColor);
								}
								while (true)
								{
									switch (4)
									{
									case 0:
										break;
									default:
										goto end_IL_05dd;
									}
									continue;
									end_IL_05dd:
									break;
								}
							}
							finally
							{
								if (enumerator5 is IDisposable)
								{
									while (true)
									{
										switch (5)
										{
										case 0:
											continue;
										}
										(enumerator5 as IDisposable).Dispose();
										break;
									}
								}
							}
						}
						else
						{
							string value = Conversions.ToString(foreColor.RGB);
							IEnumerator enumerator6 = default(IEnumerator);
							try
							{
								enumerator6 = table2.Rows[1].Cells.GetEnumerator();
								while (enumerator6.MoveNext())
								{
									((Cell)enumerator6.Current).Shape.Fill.ForeColor.RGB = Conversions.ToInteger(value);
								}
								while (true)
								{
									switch (1)
									{
									case 0:
										break;
									default:
										goto end_IL_0678;
									}
									continue;
									end_IL_0678:
									break;
								}
							}
							finally
							{
								if (enumerator6 is IDisposable)
								{
									while (true)
									{
										switch (4)
										{
										case 0:
											continue;
										}
										(enumerator6 as IDisposable).Dispose();
										break;
									}
								}
							}
						}
					}
				}
			}
			if (table.LastRow)
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
				if (table2.LastRow)
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
					if (table.Cell(table.Rows.Count, 1).Shape.Fill.Visible == MsoTriState.msoTrue)
					{
						foreColor = table.Cell(table.Rows.Count, 1).Shape.Fill.ForeColor;
						int count2 = table2.Rows.Count;
						if (foreColor.Type == MsoColorType.msoColorTypeScheme)
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
							{
								IEnumerator enumerator7 = table2.Rows[count2].Cells.GetEnumerator();
								try
								{
									while (enumerator7.MoveNext())
									{
										Styles.A((Cell)enumerator7.Current, foreColor);
									}
									while (true)
									{
										switch (5)
										{
										case 0:
											break;
										default:
											goto end_IL_0792;
										}
										continue;
										end_IL_0792:
										break;
									}
								}
								finally
								{
									IDisposable disposable2 = enumerator7 as IDisposable;
									if (disposable2 != null)
									{
										disposable2.Dispose();
									}
								}
							}
						}
						else
						{
							string value = Conversions.ToString(foreColor.RGB);
							IEnumerator enumerator8 = default(IEnumerator);
							try
							{
								enumerator8 = table2.Rows[count2].Cells.GetEnumerator();
								while (enumerator8.MoveNext())
								{
									((Cell)enumerator8.Current).Shape.Fill.ForeColor.RGB = Conversions.ToInteger(value);
								}
								while (true)
								{
									switch (2)
									{
									case 0:
										break;
									default:
										goto end_IL_0821;
									}
									continue;
									end_IL_0821:
									break;
								}
							}
							finally
							{
								if (enumerator8 is IDisposable)
								{
									while (true)
									{
										switch (5)
										{
										case 0:
											continue;
										}
										(enumerator8 as IDisposable).Dispose();
										break;
									}
								}
							}
						}
					}
				}
			}
		}
		table2 = null;
		table = null;
		foreColor = null;
	}

	private static bool A(Table A)
	{
		return true;
	}

	private static void A(Cell A, Microsoft.Office.Interop.PowerPoint.ColorFormat B)
	{
		Microsoft.Office.Interop.PowerPoint.ColorFormat foreColor = A.Shape.Fill.ForeColor;
		foreColor.ObjectThemeColor = B.ObjectThemeColor;
		foreColor.TintAndShade = B.TintAndShade;
		foreColor.Brightness = B.Brightness;
		_ = null;
	}
}
