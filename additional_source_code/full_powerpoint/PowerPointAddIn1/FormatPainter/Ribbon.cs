using System;
using System.Collections;
using A;
using MacabacusMacros;
using MacabacusMacros.Auth;
using MacabacusMacros.UI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Tools;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.Shapes;

namespace PowerPointAddIn1.FormatPainter;

public sealed class Ribbon
{
	public static void Copy()
	{
		DocumentWindow activeWindow;
		Selection selection;
		try
		{
			activeWindow = NG.A.Application.ActiveWindow;
			selection = activeWindow.Selection;
			if (Pane.IsSingleShapeSelected(selection))
			{
				Pane.CopiedProperties = new Properties(Base.SelectedShapes(selection)[1]);
				CustomTaskPane paneByHwnd = Pane.GetPaneByHwnd(activeWindow.HWND);
				if (paneByHwnd != null && paneByHwnd.Visible)
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
						((ctpFormatPainter)paneByHwnd.Control).A.PopulateProperties();
						break;
					}
				}
			}
			else
			{
				Helpers.SingleShapeRequiredError();
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			Helpers.SingleShapeRequiredError();
			ProjectData.ClearProjectError();
		}
		activeWindow = null;
		selection = null;
	}

	public static void Size()
	{
		A(A, AH.A(138018));
	}

	private static void A(Microsoft.Office.Interop.PowerPoint.Shape A, Properties B)
	{
		A.Height = B.Layout.Height;
		A.Width = B.Layout.Width;
	}

	public static void Height()
	{
		A(B, AH.A(138039));
	}

	private static void B(Microsoft.Office.Interop.PowerPoint.Shape A, Properties B)
	{
		A.Height = B.Layout.Height;
	}

	public static void Width()
	{
		A(C, AH.A(138064));
	}

	private static void C(Microsoft.Office.Interop.PowerPoint.Shape A, Properties B)
	{
		A.Width = B.Layout.Width;
	}

	public static void Position()
	{
		A(D, AH.A(138087));
	}

	private static void D(Microsoft.Office.Interop.PowerPoint.Shape A, Properties B)
	{
		A.Top = B.Layout.Top;
		A.Left = B.Layout.Left;
	}

	public static void Top()
	{
		A(E, AH.A(138116));
	}

	private static void E(Microsoft.Office.Interop.PowerPoint.Shape A, Properties B)
	{
		A.Top = B.Layout.Top;
	}

	public static void Bottom()
	{
		A(F, AH.A(138135));
	}

	private static void F(Microsoft.Office.Interop.PowerPoint.Shape A, Properties B)
	{
		A.Top = B.Layout.Bottom - A.Height;
	}

	public static void MidpointY()
	{
		if (!A())
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
			A(G, AH.A(138160));
			return;
		}
	}

	private static void G(Microsoft.Office.Interop.PowerPoint.Shape A, Properties B)
	{
		A.Top = B.Layout.MidpointY - A.Height / 2f;
	}

	public static void Left()
	{
		A(H, AH.A(138193));
	}

	private static void H(Microsoft.Office.Interop.PowerPoint.Shape A, Properties B)
	{
		A.Left = B.Layout.Left;
	}

	public static void Right()
	{
		A(I, AH.A(138214));
	}

	private static void I(Microsoft.Office.Interop.PowerPoint.Shape A, Properties B)
	{
		A.Left = B.Layout.Right - A.Width;
	}

	public static void MidpointX()
	{
		if (!Access.AllowPowerPointOperation((PlanType)4, (Restriction)1, false))
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			A(J, AH.A(138237));
			return;
		}
	}

	private static void J(Microsoft.Office.Interop.PowerPoint.Shape A, Properties B)
	{
		A.Left = B.Layout.MidpointX - A.Width / 2f;
	}

	public static void LockAspectRatio()
	{
		A(K, AH.A(138270));
	}

	private static void K(Microsoft.Office.Interop.PowerPoint.Shape A, Properties B)
	{
		A.LockAspectRatio = B.Layout.LockAspectRatio;
	}

	public static void Rotation()
	{
		if (Access.AllowPowerPointOperation((PlanType)4, (Restriction)1, false))
		{
			A(L, AH.A(138317));
		}
	}

	private static void L(Microsoft.Office.Interop.PowerPoint.Shape A, Properties B)
	{
		try
		{
			A.Rotation = B.Layout.Rotation;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	public static void Bullets()
	{
		if (Access.AllowPowerPointOperation((PlanType)4, (Restriction)1, false))
		{
			A(M, AH.A(138346));
		}
	}

	private static void M(Microsoft.Office.Interop.PowerPoint.Shape A, Properties B)
	{
		Ribbon.A(Apply.ApplyBullets, A, B);
	}

	public static void Indents()
	{
		if (!A())
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			A(N, AH.A(138373));
			return;
		}
	}

	private static void N(Microsoft.Office.Interop.PowerPoint.Shape A, Properties B)
	{
		Ribbon.A(Apply.ApplyIndents, A, B);
	}

	public static void LineSpacing()
	{
		if (!A())
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
			A(O, AH.A(138400));
			return;
		}
	}

	private static void O(Microsoft.Office.Interop.PowerPoint.Shape A, Properties B)
	{
		Ribbon.A(Apply.ApplyLineSpacing, A, B);
	}

	public static void Margins()
	{
		if (!A())
		{
			return;
		}
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
			A(P, AH.A(138437));
			return;
		}
	}

	private static void P(Microsoft.Office.Interop.PowerPoint.Shape A, Properties B)
	{
		if (A.HasTextFrame == MsoTriState.msoTrue)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
				{
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					Properties.TextBoxProperties textBox = B.TextBox;
					A.TextFrame2.MarginTop = textBox.MarginTop;
					A.TextFrame2.MarginBottom = textBox.MarginBottom;
					A.TextFrame2.MarginLeft = textBox.MarginLeft;
					A.TextFrame2.MarginRight = textBox.MarginRight;
					textBox = default(Properties.TextBoxProperties);
					return;
				}
				}
			}
		}
		checked
		{
			if (A.HasTable == MsoTriState.msoTrue)
			{
				while (true)
				{
					switch (1)
					{
					case 0:
						break;
					default:
					{
						Table table = A.Table;
						int count = table.Rows.Count;
						int count2 = table.Columns.Count;
						int num = count;
						for (int i = 1; i <= num; i++)
						{
							int num2 = count2;
							for (int j = 1; j <= num2; j++)
							{
								Cell cell = table.Cell(i, j);
								if (cell.Selected)
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
									if (cell.Shape.HasTextFrame == MsoTriState.msoTrue)
									{
										Microsoft.Office.Interop.PowerPoint.TextFrame2 textFrame = cell.Shape.TextFrame2;
										textFrame.MarginTop = B.TextBox.MarginTop;
										textFrame.MarginBottom = B.TextBox.MarginBottom;
										textFrame.MarginLeft = B.TextBox.MarginLeft;
										textFrame.MarginRight = B.TextBox.MarginRight;
										_ = null;
									}
								}
								cell = null;
							}
							while (true)
							{
								switch (7)
								{
								case 0:
									break;
								default:
									goto end_IL_0171;
								}
								continue;
								end_IL_0171:
								break;
							}
						}
						table = null;
						return;
					}
					}
				}
			}
			if (A.HasSmartArt != MsoTriState.msoTrue)
			{
				return;
			}
			IEnumerator enumerator = default(IEnumerator);
			while (true)
			{
				switch (1)
				{
				case 0:
					continue;
				}
				try
				{
					enumerator = A.SmartArt.AllNodes.GetEnumerator();
					while (enumerator.MoveNext())
					{
						SmartArtNode obj = (SmartArtNode)enumerator.Current;
						Properties.TextBoxProperties textBox2 = B.TextBox;
						obj.TextFrame2.MarginTop = textBox2.MarginTop;
						obj.TextFrame2.MarginBottom = textBox2.MarginBottom;
						obj.TextFrame2.MarginLeft = textBox2.MarginLeft;
						obj.TextFrame2.MarginRight = textBox2.MarginRight;
						textBox2 = default(Properties.TextBoxProperties);
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
					if (enumerator is IDisposable)
					{
						while (true)
						{
							switch (4)
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
		}
	}

	public static void TextWrap()
	{
		if (A())
		{
			A(Q, AH.A(138464));
		}
	}

	private static void Q(Microsoft.Office.Interop.PowerPoint.Shape A, Properties B)
	{
		if (A.HasTextFrame == MsoTriState.msoTrue)
		{
			A.TextFrame2.WordWrap = B.TextBox.WordWrap;
		}
	}

	public static void AutoSize()
	{
		if (!A())
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			A(R, AH.A(138495));
			return;
		}
	}

	private static void R(Microsoft.Office.Interop.PowerPoint.Shape A, Properties B)
	{
		if (A.HasTextFrame != MsoTriState.msoTrue)
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			A.TextFrame2.AutoSize = B.TextBox.AutoSize;
			return;
		}
	}

	public static void HorizontalAlignment()
	{
		A(S, AH.A(138526));
	}

	private static void S(Microsoft.Office.Interop.PowerPoint.Shape A, Properties B)
	{
		if (A.HasTextFrame != MsoTriState.msoTrue)
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
			A.TextFrame2.HorizontalAnchor = B.TextBox.HorizontalAnchor;
			A.TextFrame2.TextRange.ParagraphFormat.Alignment = B.TextBox.HorizontalAlignment;
			return;
		}
	}

	public static void VerticalAlignment()
	{
		A(T, AH.A(138579));
	}

	private static void T(Microsoft.Office.Interop.PowerPoint.Shape A, Properties B)
	{
		if (A.HasTextFrame == MsoTriState.msoTrue)
		{
			A.TextFrame2.VerticalAnchor = B.TextBox.VerticalAnchor;
		}
	}

	public static void Adjustments()
	{
		if (A())
		{
			A(U, AH.A(138628));
		}
	}

	private static void U(Microsoft.Office.Interop.PowerPoint.Shape A, Properties B)
	{
		checked
		{
			if (A.Type == MsoShapeType.msoAutoShape)
			{
				int num = B.AutoShape.Adjustments.Count - 1;
				for (int i = 0; i <= num; i++)
				{
					A.Adjustments[i + 1] = B.AutoShape.Adjustments[i];
				}
			}
		}
	}

	public static void AutoShapeType()
	{
		if (!A())
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			A(V, AH.A(138663));
			return;
		}
	}

	private static void V(Microsoft.Office.Interop.PowerPoint.Shape A, Properties B)
	{
		if (A.Type != MsoShapeType.msoAutoShape)
		{
			return;
		}
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
			A.AutoShapeType = B.AutoShape.Type;
			return;
		}
	}

	private static void A(Action<Microsoft.Office.Interop.PowerPoint.Shape, Properties> A, string B)
	{
		Application application = NG.A.Application;
		Selection selection;
		if (Pane.CopiedProperties != null)
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
			try
			{
				selection = application.ActiveWindow.Selection;
				if (selection.Type == PpSelectionType.ppSelectionShapes)
				{
					goto IL_005d;
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
					goto IL_005d;
				}
				Ribbon.A();
				goto end_IL_002c;
				IL_005d:
				application.StartNewUndoEntry();
				try
				{
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
						{
							IEnumerator enumerator = selection.ChildShapeRange.GetEnumerator();
							try
							{
								while (enumerator.MoveNext())
								{
									Microsoft.Office.Interop.PowerPoint.Shape arg = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator.Current;
									A(arg, Pane.CopiedProperties);
								}
								while (true)
								{
									switch (7)
									{
									case 0:
										break;
									default:
										goto end_IL_00ad;
									}
									continue;
									end_IL_00ad:
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
						}
					}
					else
					{
						IEnumerator enumerator2 = default(IEnumerator);
						try
						{
							enumerator2 = selection.ShapeRange.GetEnumerator();
							while (enumerator2.MoveNext())
							{
								Microsoft.Office.Interop.PowerPoint.Shape arg2 = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator2.Current;
								A(arg2, Pane.CopiedProperties);
							}
							while (true)
							{
								switch (2)
								{
								case 0:
									break;
								default:
									goto end_IL_010a;
								}
								continue;
								end_IL_010a:
								break;
							}
						}
						finally
						{
							if (enumerator2 is IDisposable)
							{
								while (true)
								{
									switch (3)
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
					clsReporting.LogActivity((ActivityApp)2, (ActivityCategory)1, B);
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					Forms.ErrorMessage(ex2.Message);
					clsReporting.LogException(ex2);
					ProjectData.ClearProjectError();
				}
				end_IL_002c:;
			}
			catch (Exception ex3)
			{
				ProjectData.SetProjectError(ex3);
				Exception ex4 = ex3;
				Ribbon.A();
				ProjectData.ClearProjectError();
			}
		}
		else
		{
			Forms.WarningMessage(AH.A(138706));
		}
		application = null;
		selection = null;
	}

	private static void A(Action<TextRange2, Properties.TextBoxProperties> A, Microsoft.Office.Interop.PowerPoint.Shape B, Properties C)
	{
		if (B.HasTextFrame == MsoTriState.msoTrue)
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
					A(B.TextFrame2.TextRange, C.TextBox);
					return;
				}
			}
		}
		checked
		{
			if (B.HasTable == MsoTriState.msoTrue)
			{
				while (true)
				{
					switch (7)
					{
					case 0:
						break;
					default:
					{
						Table table = B.Table;
						int count = table.Rows.Count;
						int count2 = table.Columns.Count;
						int num = count;
						for (int i = 1; i <= num; i++)
						{
							int num2 = count2;
							for (int j = 1; j <= num2; j++)
							{
								Cell cell = table.Cell(i, j);
								if (cell.Selected)
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
									if (cell.Shape.HasTextFrame == MsoTriState.msoTrue)
									{
										A(cell.Shape.TextFrame2.TextRange, C.TextBox);
									}
								}
								cell = null;
							}
						}
						table = null;
						return;
					}
					}
				}
			}
			if (B.HasSmartArt != MsoTriState.msoTrue)
			{
				return;
			}
			IEnumerator enumerator = default(IEnumerator);
			while (true)
			{
				switch (4)
				{
				case 0:
					continue;
				}
				try
				{
					enumerator = B.SmartArt.AllNodes.GetEnumerator();
					while (enumerator.MoveNext())
					{
						SmartArtNode smartArtNode = (SmartArtNode)enumerator.Current;
						A(smartArtNode.TextFrame2.TextRange, C.TextBox);
					}
					return;
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
		}
	}

	private static void A()
	{
		Forms.WarningMessage(AH.A(73308));
	}

	private static bool A()
	{
		return Access.AllowPowerPointOperation((PlanType)4, (Restriction)1, false);
	}
}
