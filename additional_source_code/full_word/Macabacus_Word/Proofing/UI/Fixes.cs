using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Controls.Primitives;
using System.Windows.Forms;
using System.Windows.Media;
using A;
using MacabacusMacros;
using MacabacusMacros.Proofing;
using MacabacusMacros.Proofing.UI;
using MacabacusMacros.UI;
using Macabacus_Word.Proofing.Errors;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using Microsoft.VisualBasic.CompilerServices;

namespace Macabacus_Word.Proofing.UI;

public sealed class Fixes
{
	private struct GC
	{
		public double A;

		public double B;

		public double C;

		public double D;

		public List<Rect> A;
	}

	[CompilerGenerated]
	internal sealed class HC
	{
		public BaseError A;

		public HC(HC A)
		{
			if (A != null)
			{
				this.A = A.A;
			}
		}

		[SpecialName]
		internal bool A(BaseError A)
		{
			if ((object)((object)A).GetType() == ((object)this.A).GetType())
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
						return A != this.A;
					}
				}
			}
			return false;
		}
	}

	private static GC m_A;

	private static System.Windows.Media.Color m_A = System.Windows.Media.Colors.Transparent;

	private static List<Type> m_A;

	private static System.Windows.Media.Color A
	{
		get
		{
			return Fixes.m_A;
		}
		set
		{
			Fixes.m_A = value;
		}
	}

	private static List<Type> A
	{
		get
		{
			return Fixes.m_A;
		}
		set
		{
			Fixes.m_A = value;
		}
	}

	public static void DefaultFixButtonClicked(BaseError err)
	{
		if (!((BaseError)err).IsFixed)
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
					if (A(err))
					{
						while (true)
						{
							switch (7)
							{
							case 0:
								break;
							default:
							{
								System.Drawing.Color color = Fixes.FindNearestColor(((BaseError)err).NonconformingColor);
								A(color);
								A(Fixes.ConvertToWpfColor(color), B: true);
								return;
							}
							}
						}
					}
					FixItem(0);
					return;
				}
			}
		}
		D(err);
	}

	public static void ShowOptions(BaseError err, ToggleButton btnFix)
	{
		//IL_006f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0074: Unknown result type (might be due to invalid IL or missing references)
		//IL_0086: Expected O, but got Unknown
		//IL_008b: Expected O, but got Unknown
		//IL_003f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0044: Unknown result type (might be due to invalid IL or missing references)
		//IL_0056: Expected O, but got Unknown
		//IL_005b: Expected O, but got Unknown
		Callout.DoNotClose = true;
		if (A(err))
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
				{
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					wpfFixPalette val = new wpfFixPalette(btnFix, (Visual)Pane.TaskPane, ((Conventions)Main.Analysis.Conventions).ColorPalette, false);
					((System.Windows.Window)val).Closed += B;
					((System.Windows.Window)val).Show();
					_ = null;
					return;
				}
				}
			}
		}
		wpfFixMenu val2 = new wpfFixMenu(btnFix, ((BaseError)err).DisplayText, (Visual)Pane.TaskPane, false);
		((System.Windows.Window)val2).Closed += A;
		((System.Windows.Window)val2).Show();
		_ = null;
	}

	private static void A(object A, EventArgs B)
	{
		//IL_0001: Unknown result type (might be due to invalid IL or missing references)
		//IL_0007: Expected O, but got Unknown
		wpfFixMenu val = (wpfFixMenu)A;
		if (val.Index > -1)
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			FixItem(val.Index);
		}
		val = null;
		Pane.RefocusActiveListBoxItem();
		Callout.DoNotClose = false;
	}

	private static void B(object A, EventArgs B)
	{
		//IL_0001: Unknown result type (might be due to invalid IL or missing references)
		//IL_0007: Expected O, but got Unknown
		wpfFixPalette val = (wpfFixPalette)A;
		if (val.SelectedColor != System.Windows.Media.Colors.Transparent)
		{
			Fixes.B(val.SelectedColor);
		}
		val = null;
		Pane.RefocusActiveListBoxItem();
		Callout.DoNotClose = false;
	}

	private static bool A(BaseError A)
	{
		return ((BaseError)A).HasColorFix;
	}

	public static void FixItem(int i)
	{
		A(Pane.ActiveItem, i);
	}

	private static void B(System.Windows.Media.Color A)
	{
		Fixes.A(clsColors.RGB2Color(Colors.Color2RGB(A)));
		Fixes.A(A, B: true);
	}

	private static void A(System.Drawing.Color A)
	{
		Fixes.A();
		((BaseError)Pane.ActiveItem).FixAction(A);
		((BaseError)Pane.ActiveItem).IsFixed = true;
	}

	private static void A(System.Windows.Media.Color A, bool B)
	{
		bool flag = false;
		System.Windows.Media.Color a = Fixes.A;
		Conventions conventions;
		List<PaletteColor> colorPalette = ((Conventions)(conventions = Main.Analysis.Conventions)).ColorPalette;
		Fixes.UpdateColorPaletteUsage(A, B, a, ref colorPalette, ref flag);
		((Conventions)conventions).ColorPalette = colorPalette;
		if (B)
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
			if (!flag)
			{
				while (true)
				{
					switch (7)
					{
					case 0:
						break;
					default:
						Fixes.A = A;
						return;
					}
				}
			}
		}
		Fixes.A = System.Windows.Media.Colors.Transparent;
	}

	private static void A(BaseError A, int B)
	{
		HC a = default(HC);
		HC CS_0024_003C_003E8__locals25 = new HC(a);
		CS_0024_003C_003E8__locals25.A = A;
		Fixes.A();
		try
		{
			ErrorType type = CS_0024_003C_003E8__locals25.A.Type;
			if (type <= ErrorType.ImageDistortion)
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
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				switch (type)
				{
				case ErrorType.Text:
				case ErrorType.BulletPunctuation:
				case ErrorType.BulletSize:
				case ErrorType.BulletFontFamily:
				case ErrorType.BulletIndent:
				case ErrorType.MultipleFontFamilies:
				case ErrorType.LineSpacing:
					((BaseError)CS_0024_003C_003E8__locals25.A).FixAction(B);
					C(CS_0024_003C_003E8__locals25.A);
					break;
				case ErrorType.ImageDistortion:
					((BaseError)CS_0024_003C_003E8__locals25.A).FixAction(B);
					Fixes.A(CS_0024_003C_003E8__locals25.A);
					break;
				case ErrorType.FillTransparency:
					((BaseError)CS_0024_003C_003E8__locals25.A).FixAction(B);
					break;
				}
			}
			else if (type <= ErrorType.ProofingLanguage)
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
				if (type != ErrorType.TableCellMargins)
				{
					if (type != ErrorType.ProofingLanguage)
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
					else
					{
						((BaseError)CS_0024_003C_003E8__locals25.A).FixAction();
					}
				}
				else
				{
					((BaseError)CS_0024_003C_003E8__locals25.A).FixAction(B);
					Fixes.B(CS_0024_003C_003E8__locals25.A);
				}
			}
			else
			{
				switch (type)
				{
				case ErrorType.ShapeOutsideMargins:
					((BaseError)CS_0024_003C_003E8__locals25.A).FixAction();
					Fixes.A(CS_0024_003C_003E8__locals25.A);
					break;
				case ErrorType.ChartLegendEntryMissing:
					((BaseError)CS_0024_003C_003E8__locals25.A).FixAction();
					break;
				case ErrorType.ChartDataLabelMissing:
					((BaseError)CS_0024_003C_003E8__locals25.A).FixAction();
					break;
				case ErrorType.ChartDataLabelsInconsistent:
					((BaseError)CS_0024_003C_003E8__locals25.A).FixAction();
					break;
				case ErrorType.ChartDataLabelNumberFormats:
					((BaseError)CS_0024_003C_003E8__locals25.A).FixAction(B);
					break;
				}
			}
			((BaseError)CS_0024_003C_003E8__locals25.A).IsFixed = true;
			if (CS_0024_003C_003E8__locals25.A.Type == ErrorType.Text && !(CS_0024_003C_003E8__locals25.A is HyphenWordsInconsistent))
			{
				double top = default(double);
				double top2 = default(double);
				while (true)
				{
					switch (3)
					{
					case 0:
						continue;
					}
					if (Fixes.A == null)
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
						Fixes.A = new List<Type>();
					}
					if (Fixes.A.Contains(((object)CS_0024_003C_003E8__locals25.A).GetType()))
					{
						break;
					}
					try
					{
						List<BaseError> list = Main.Analysis.Errors.Where([SpecialName] (BaseError baseError) =>
						{
							if ((object)((object)baseError).GetType() == ((object)CS_0024_003C_003E8__locals25.A).GetType())
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
										return baseError != CS_0024_003C_003E8__locals25.A;
									}
								}
							}
							return false;
						}).ToList();
						int count = list.Count;
						if (count <= 0)
						{
							break;
						}
						Callout.DoNotClose = true;
						if (Callout.Dialog != null)
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
							top = Callout.Dialog.Top;
							Callout.Dialog.Top = -10000.0;
						}
						if (Callout.MarchingAnts != null)
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
							top2 = ((System.Windows.Window)(object)Callout.MarchingAnts).Top;
							((System.Windows.Window)(object)Callout.MarchingAnts).Top = -10000.0;
						}
						if (System.Windows.Forms.MessageBox.Show(XC.A(38195) + count + XC.A(38262), XC.A(2438), MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) == DialogResult.Yes)
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
							try
							{
								using (List<BaseError>.Enumerator enumerator = list.GetEnumerator())
								{
									while (enumerator.MoveNext())
									{
										BaseError current = enumerator.Current;
										((BaseError)current).FixAction(B);
										((BaseError)current).IsFixed = true;
										Pane.TaskPane.RemoveItem(current, blnAnimate: false);
									}
									while (true)
									{
										switch (7)
										{
										case 0:
											break;
										default:
											goto end_IL_033d;
										}
										continue;
										end_IL_033d:
										break;
									}
								}
								Pane.TaskPane.RemoveItemAndNavigate(CS_0024_003C_003E8__locals25.A);
							}
							catch (Exception ex)
							{
								ProjectData.SetProjectError(ex);
								Exception ex2 = ex;
								System.Windows.Forms.MessageBox.Show(XC.A(38299) + ex2.Message, XC.A(2438), MessageBoxButtons.OK, MessageBoxIcon.Hand);
								ProjectData.ClearProjectError();
							}
						}
						else
						{
							Fixes.A.Add(((object)CS_0024_003C_003E8__locals25.A).GetType());
						}
						Callout.Dialog.Top = top;
						((System.Windows.Window)(object)Callout.MarchingAnts).Top = top2;
						Pane.RefocusActiveListBoxItem();
						Callout.DoNotClose = false;
						break;
					}
					catch (Exception ex3)
					{
						ProjectData.SetProjectError(ex3);
						Exception ex4 = ex3;
						ProjectData.ClearProjectError();
					}
					finally
					{
						List<BaseError> list = null;
					}
					break;
				}
			}
		}
		catch (Exception ex5)
		{
			ProjectData.SetProjectError(ex5);
			Exception ex6 = ex5;
			System.Windows.Forms.MessageBox.Show(XC.A(38414) + ex6.Message, XC.A(2438), MessageBoxButtons.OK, MessageBoxIcon.Hand);
			ProjectData.ClearProjectError();
		}
		CS_0024_003C_003E8__locals25.A = null;
	}

	private static void A(BaseError A)
	{
		List<Rect> list = new List<Rect>();
		Rect objectRectangle;
		try
		{
			objectRectangle = Callout.GetObjectRectangle(A.Shape);
			list.Add(objectRectangle);
			Fixes.A(objectRectangle.X, objectRectangle.Y, list);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		list = null;
		objectRectangle = default(Rect);
	}

	private static void B(BaseError A)
	{
		double num = 10000.0;
		double num2 = 10000.0;
		List<Rect> list = new List<Rect>();
		checked
		{
			try
			{
				int num3 = A.Shapes.Count - 1;
				for (int i = 0; i <= num3; i++)
				{
					Microsoft.Office.Interop.Word.Shapes shapes = A.Shapes;
					object Index = i;
					Rect objectRectangle = Callout.GetObjectRectangle(shapes[ref Index]);
					list.Add(objectRectangle);
					if (objectRectangle.X < num2)
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
						num2 = objectRectangle.X;
					}
					if (objectRectangle.Y < num)
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
						num = objectRectangle.Y;
					}
					objectRectangle = default(Rect);
				}
				while (true)
				{
					switch (3)
					{
					case 0:
						continue;
					}
					Fixes.A(num2, num, list);
					list = null;
					return;
				}
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
		}
	}

	private static void C(BaseError A)
	{
		try
		{
			if (Callout.UseRelativePosition(A))
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
				_ = A.Shape.Left;
				_ = A.Shape.Top;
			}
			List<Rect> list = new List<Rect>();
			double num = 10000.0;
			double num2 = 10000.0;
			using (List<Range>.Enumerator enumerator = A.Ranges.GetEnumerator())
			{
				while (enumerator.MoveNext())
				{
					Rect objectRectangle = Callout.GetObjectRectangle(enumerator.Current);
					list.Add(objectRectangle);
					if (objectRectangle.X < num2)
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
						num2 = objectRectangle.X;
					}
					if (objectRectangle.Y < num)
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
						num = objectRectangle.Y;
					}
					objectRectangle = default(Rect);
				}
				while (true)
				{
					switch (6)
					{
					case 0:
						break;
					default:
						goto end_IL_00cb;
					}
					continue;
					end_IL_00cb:
					break;
				}
			}
			Fixes.A(num2, num, list);
			list = null;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	private static void A(double A, double B, List<Rect> C)
	{
		Callout.Reposition(Callout.Dialog, A, B);
		Fixes.B(C);
	}

	private static void B(List<Rect> A)
	{
		wpfMarchingAnts marchingAnts = Callout.MarchingAnts;
		marchingAnts.ClearMarchingAnts();
		marchingAnts.AddMarchingAnts(A);
		_ = null;
	}

	private static void D(BaseError A)
	{
		PC.A.Application.CommandBars.ExecuteMso(XC.A(38463));
		System.Windows.Forms.Application.DoEvents();
		BaseError activeItem = Pane.ActiveItem;
		BaseError baseError = activeItem;
		ErrorType type = baseError.Type;
		if (type <= ErrorType.MultipleFontFamilies)
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			if (type != ErrorType.BulletPunctuation)
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
				if (type != ErrorType.MultipleFontFamilies)
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
					goto IL_017b;
				}
				if (baseError.Type == ErrorType.MultipleFontFamilies)
				{
					try
					{
						_ = ((BaseError)baseError).TextRanges[0].BoundTop;
					}
					catch (Exception ex)
					{
						ProjectData.SetProjectError(ex);
						Exception ex2 = ex;
						((BaseError)baseError).TextRanges[0] = baseError.Shape.TextFrame2.TextRange;
						ProjectData.ClearProjectError();
					}
				}
				goto IL_019a;
			}
		}
		else if (type != ErrorType.LineSpacing)
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
			if (type != ErrorType.ProofingLanguage)
			{
				goto IL_017b;
			}
			while (true)
			{
				switch (1)
				{
				case 0:
					continue;
				}
				break;
			}
		}
		checked
		{
			try
			{
				_ = ((BaseError)baseError).TextRanges[((BaseError)baseError).TextRanges.Count - 1].BoundTop;
			}
			catch (Exception ex3)
			{
				ProjectData.SetProjectError(ex3);
				Exception ex4 = ex3;
				TextRange2 textRange = activeItem.Shape.TextFrame2.TextRange.get_Paragraphs(-1, -1);
				((BaseError)activeItem).TextRanges[((BaseError)activeItem).TextRanges.Count - 1] = textRange.Item(textRange.Count);
				textRange = null;
				ProjectData.ClearProjectError();
			}
			goto IL_019a;
		}
		IL_019a:
		baseError = null;
		activeItem = null;
		B();
		((BaseError)A).IsFixed = false;
		return;
		IL_017b:
		if (activeItem is BaseColorError)
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
			Fixes.A(Fixes.A, B: false);
		}
		goto IL_019a;
	}

	private static void A()
	{
		try
		{
			Fixes.m_A = default(GC);
			Fixes.m_A.A = Callout.Dialog.Left;
			Fixes.m_A.B = Callout.Dialog.Top;
			Fixes.m_A.C = ((System.Windows.Window)(object)Callout.MarchingAnts).Left;
			Fixes.m_A.D = ((System.Windows.Window)(object)Callout.MarchingAnts).Top;
			Fixes.m_A.A = Callout.DashBoxes.ToList();
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	private static void B()
	{
		try
		{
			Callout.Dialog.Left = Fixes.m_A.A;
			Callout.Dialog.Top = Fixes.m_A.B;
			((System.Windows.Window)(object)Callout.MarchingAnts).Left = Fixes.m_A.C;
			((System.Windows.Window)(object)Callout.MarchingAnts).Top = Fixes.m_A.D;
			B(Fixes.m_A.A);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}
}
