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
using MacabacusMacros.Auth;
using MacabacusMacros.Proofing;
using MacabacusMacros.Proofing.UI;
using MacabacusMacros.UI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.DeckCheck.Errors;

namespace PowerPointAddIn1.DeckCheck.UI;

public sealed class Fixes
{
	private struct RC
	{
		public double A;

		public double B;

		public double C;

		public double D;

		public List<Rect> A;
	}

	[CompilerGenerated]
	internal sealed class SC
	{
		public BaseError A;

		public SC(SC A)
		{
			if (A == null)
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
				this.A = A.A;
				return;
			}
		}

		[SpecialName]
		internal bool A(BaseError A)
		{
			if ((object)((object)A).GetType() == ((object)this.A).GetType())
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
						return A != this.A;
					}
				}
			}
			return false;
		}
	}

	[CompilerGenerated]
	internal sealed class TC
	{
		public string A;

		public SC A;

		public TC(TC A)
		{
			if (A == null)
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
				this.A = A.A;
				return;
			}
		}

		[SpecialName]
		internal bool A(BaseError A)
		{
			if ((object)((object)A).GetType() == ((object)this.A.A).GetType())
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
				if (A != this.A.A)
				{
					while (true)
					{
						switch (1)
						{
						case 0:
							break;
						default:
							return Operators.CompareString(((CorporateDictionary)A).RuleId, this.A, TextCompare: false) == 0;
						}
					}
				}
			}
			return false;
		}
	}

	private static RC m_A;

	private static System.Windows.Media.Color m_A = System.Windows.Media.Colors.Transparent;

	[CompilerGenerated]
	private static List<Type> m_A;

	[CompilerGenerated]
	private static bool m_A;

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

	private static List<Type> DismissedFixAllTypes
	{
		[CompilerGenerated]
		get
		{
			return Fixes.m_A;
		}
		[CompilerGenerated]
		set
		{
			Fixes.m_A = value;
		}
	}

	private static bool RefocusTaskPane
	{
		[CompilerGenerated]
		get
		{
			return Fixes.m_A;
		}
		[CompilerGenerated]
		set
		{
			Fixes.m_A = value;
		}
	}

	public static void DefaultFixButtonClicked(BaseError err, bool suppressMsgs = false)
	{
		if (!((BaseError)err).IsFixed)
		{
			try
			{
				if (((BaseError)err).FixHasProgrammaticIssue)
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
							A(AH.A(55727), suppressMsgs);
							return;
						}
					}
				}
				if (A(err))
				{
					while (true)
					{
						switch (2)
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
				if (((BaseError)err).IsFixable())
				{
					while (true)
					{
						switch (1)
						{
						case 0:
							break;
						default:
							A(0);
							return;
						}
					}
				}
				return;
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception a = ex;
				A(a, err, suppressMsgs);
				ProjectData.ClearProjectError();
				return;
			}
		}
		I(err);
	}

	public static void ShowOptions(BaseError err, ToggleButton btnFix, bool blnRefocusPane)
	{
		//IL_0079: Unknown result type (might be due to invalid IL or missing references)
		//IL_007e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0090: Expected O, but got Unknown
		//IL_0095: Expected O, but got Unknown
		//IL_0046: Unknown result type (might be due to invalid IL or missing references)
		//IL_004b: Unknown result type (might be due to invalid IL or missing references)
		//IL_005d: Expected O, but got Unknown
		//IL_0062: Expected O, but got Unknown
		RefocusTaskPane = blnRefocusPane;
		Callout.DoNotClose = true;
		if (A(err))
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
				{
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					wpfFixPalette val = new wpfFixPalette(btnFix, (Visual)Pane.TaskPane, ((Conventions)Main.Analysis.Conventions).ColorPalette, !blnRefocusPane);
					((Window)val).Closed += B;
					((Window)val).Show();
					_ = null;
					return;
				}
				}
			}
		}
		wpfFixMenu val2 = new wpfFixMenu(btnFix, ((BaseError)err).DisplayText, (Visual)Pane.TaskPane, !blnRefocusPane);
		((Window)val2).Closed += A;
		((Window)val2).Show();
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
			try
			{
				Fixes.A(val.Index);
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception a = ex;
				Fixes.A(a, Pane.ActiveItem, val.SuppressMsgs);
				ProjectData.ClearProjectError();
			}
		}
		val = null;
		if (RefocusTaskPane)
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
			Pane.F();
		}
		Callout.DoNotClose = false;
	}

	private static void B(object A, EventArgs B)
	{
		//IL_0001: Unknown result type (might be due to invalid IL or missing references)
		//IL_0007: Expected O, but got Unknown
		wpfFixPalette val = (wpfFixPalette)A;
		if (val.SelectedColor != System.Windows.Media.Colors.Transparent)
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
			try
			{
				Fixes.B(val.SelectedColor);
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception a = ex;
				Fixes.A(a, Pane.ActiveItem, val.SuppressMsgs);
				ProjectData.ClearProjectError();
			}
		}
		val = null;
		if (RefocusTaskPane)
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
			Pane.F();
		}
		Callout.DoNotClose = false;
	}

	private static void A(Exception A, BaseError B, bool C)
	{
		string a;
		if (!((BaseError)B).ProgrammaticFixFailIsLikely)
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
			a = AH.A(55889);
		}
		else
		{
			a = AH.A(56049);
		}
		Fixes.A(a, C);
		((BaseError)B).HasFix = false;
		((BaseError)B).HasColorFix = false;
		((BaseError)B).FixHasProgrammaticIssue = true;
		if (((BaseError)B).ProgrammaticFixFailIsLikely)
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
			clsReporting.LogException((Exception)new TimeZoneNotFoundException(string.Format(AH.A(56275), A.Message), A));
			return;
		}
	}

	private static bool A(BaseError A)
	{
		return ((BaseError)A).HasColorFix;
	}

	private static void A(int A)
	{
		if (Pane.ActiveItem == null)
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
			Callout.DoNotClose = true;
			if (!Access.AllowPowerPointOperation((PlanType)4, (Restriction)1, false))
			{
				Callout.DoNotClose = false;
				return;
			}
			Callout.DoNotClose = false;
			Fixes.A(Pane.ActiveItem, A);
			return;
		}
	}

	private static void B(System.Windows.Media.Color A)
	{
		Fixes.A(clsColors.RGB2Color(Colors.Color2RGB(A)));
		Fixes.A(A, B: true);
	}

	private static void A(System.Drawing.Color A)
	{
		Callout.DoNotClose = true;
		if (!Access.AllowPowerPointOperation((PlanType)4, (Restriction)1, false))
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
					Callout.DoNotClose = false;
					return;
				}
			}
		}
		Callout.DoNotClose = false;
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
			if (!flag)
			{
				while (true)
				{
					switch (1)
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
		SC a = default(SC);
		SC CS_0024_003C_003E8__locals65 = new SC(a);
		CS_0024_003C_003E8__locals65.A = A;
		Fixes.A();
		try
		{
			ErrorType type = CS_0024_003C_003E8__locals65.A.Type;
			if (type <= ErrorType.StrikethroughFont)
			{
				if (type <= ErrorType.CrookedLine)
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
					if (type != ErrorType.Text)
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
						switch (type)
						{
						case ErrorType.BulletPunctuation:
						case ErrorType.BulletSize:
						case ErrorType.BulletFontFamily:
						case ErrorType.BulletIndent:
						case ErrorType.MultipleFontFamilies:
						case ErrorType.LineSpacing:
							break;
						case ErrorType.ShapeOutOfBounds:
							((BaseError)CS_0024_003C_003E8__locals65.A).FixAction();
							Callout.DoNotClose = true;
							CS_0024_003C_003E8__locals65.A.Slide.Application.ActivePresentation.Slides[1].Select();
							CS_0024_003C_003E8__locals65.A.Slide.Select();
							Callout.DoNotClose = false;
							Pane.F();
							Fixes.A(CS_0024_003C_003E8__locals65.A);
							goto IL_053a;
						case ErrorType.MasterShapePosition:
						case ErrorType.MisalignedShape:
						case ErrorType.RotatedShape:
						case ErrorType.CrookedLine:
							goto IL_02b9;
						case ErrorType.ShrinkTextOnOverflow:
							goto IL_03cc;
						case ErrorType.FillTransparency:
							((BaseError)CS_0024_003C_003E8__locals65.A).FixAction(B);
							goto IL_053a;
						default:
							goto IL_053a;
						}
					}
				}
				else
				{
					switch (type)
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
						switch (type)
						{
						case ErrorType.FractionalFontSize:
							break;
						case ErrorType.StrikethroughFont:
							goto IL_0203;
						case ErrorType.MaxMinFontSize:
							goto IL_0409;
						default:
							goto IL_053a;
						}
						break;
					case ErrorType.IllegalFont:
						break;
					case ErrorType.PlaceholderFillMismatch:
					case ErrorType.PlaceholderFontColorMismatch:
					case ErrorType.PlaceholderMarginsMismatch:
						goto IL_0203;
					case ErrorType.PlaceholderFontStyleMismatch:
					case ErrorType.PlaceholderBulletMismatch:
					case ErrorType.PlaceholderIndentMismatch:
						((BaseError)CS_0024_003C_003E8__locals65.A).FixAction();
						D(CS_0024_003C_003E8__locals65.A);
						goto IL_053a;
					case ErrorType.PlaceholderLayoutMismatch:
						((BaseError)CS_0024_003C_003E8__locals65.A).FixAction();
						Fixes.A(CS_0024_003C_003E8__locals65.A);
						goto IL_053a;
					case ErrorType.HiddenSlide:
					case ErrorType.HiddenShape:
					case ErrorType.Hyperlinks:
						goto IL_02d4;
					case ErrorType.TableCellMargins:
						((BaseError)CS_0024_003C_003E8__locals65.A).FixAction(B);
						C(CS_0024_003C_003E8__locals65.A);
						goto IL_053a;
					case ErrorType.ImageDistortion:
						goto IL_03cc;
					case ErrorType.LinkedPicture:
						((BaseError)CS_0024_003C_003E8__locals65.A).FixAction();
						goto IL_053a;
					case ErrorType.ProofingLanguage:
						((BaseError)CS_0024_003C_003E8__locals65.A).FixAction();
						goto IL_053a;
					case ErrorType.AgendaNotUpdated:
						((BaseError)CS_0024_003C_003E8__locals65.A).FixAction();
						goto IL_053a;
					case ErrorType.AirplaneMode:
					case ErrorType.AgendaMissing:
					case ErrorType.LinkBroken:
					case ErrorType.LinkNewerVersionAvailable:
					case (ErrorType)67:
					case (ErrorType)68:
					case ErrorType.MissingSlideNumber:
					case ErrorType.ExcessSlideNumber:
					case ErrorType.FootnoteMissing:
					case ErrorType.FootnotesSequence:
						goto IL_053a;
						IL_0203:
						((BaseError)CS_0024_003C_003E8__locals65.A).FixAction();
						goto IL_053a;
					}
				}
				((BaseError)CS_0024_003C_003E8__locals65.A).FixAction(B);
				D(CS_0024_003C_003E8__locals65.A);
			}
			else if (type <= ErrorType.Ink)
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
				if ((uint)(type - 120) <= 2u)
				{
					goto IL_02d4;
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
				if (type != ErrorType.Ink)
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
					((BaseError)CS_0024_003C_003E8__locals65.A).FixAction(B);
				}
			}
			else if (type != ErrorType.CorporateDictionary)
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
				if (type == ErrorType.ShapeOutsideMargins)
				{
					goto IL_02b9;
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
				switch (type)
				{
				case ErrorType.ChartLegendEntryMissing:
				{
					((BaseError)CS_0024_003C_003E8__locals65.A).FixAction();
					Rect legendRectangle;
					try
					{
						legendRectangle = MarchingAnts.GetLegendRectangle(CS_0024_003C_003E8__locals65.A.Shape.Chart.Legend, MarchingAnts.ChartLeftOffset(CS_0024_003C_003E8__locals65.A.Shape), MarchingAnts.ChartTopOffset(CS_0024_003C_003E8__locals65.A.Shape));
						Fixes.A(legendRectangle.X, legendRectangle.Y, new List<Rect>(new Rect[1] { legendRectangle }));
					}
					catch (Exception ex)
					{
						ProjectData.SetProjectError(ex);
						Exception ex2 = ex;
						ProjectData.ClearProjectError();
					}
					legendRectangle = default(Rect);
					break;
				}
				case ErrorType.ChartDataLabelMissing:
					((BaseError)CS_0024_003C_003E8__locals65.A).FixAction();
					break;
				case ErrorType.ChartDataLabelsInconsistent:
					((BaseError)CS_0024_003C_003E8__locals65.A).FixAction();
					break;
				case ErrorType.ChartDataLabelNumberFormats:
					((BaseError)CS_0024_003C_003E8__locals65.A).FixAction(B);
					break;
				}
			}
			else
			{
				((BaseError)CS_0024_003C_003E8__locals65.A).FixAction(B);
				D(CS_0024_003C_003E8__locals65.A);
			}
			goto IL_053a;
			IL_0409:
			((BaseError)CS_0024_003C_003E8__locals65.A).FixAction();
			if (CS_0024_003C_003E8__locals65.A is MinMaxFontSize)
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
				D(CS_0024_003C_003E8__locals65.A);
			}
			else if (CS_0024_003C_003E8__locals65.A is ChartTitleFontSize)
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
				E(CS_0024_003C_003E8__locals65.A);
			}
			else if (CS_0024_003C_003E8__locals65.A is ChartLegendFontSize)
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
				F(CS_0024_003C_003E8__locals65.A);
			}
			else if (!(CS_0024_003C_003E8__locals65.A is ChartDataTableFontSize))
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
				if (CS_0024_003C_003E8__locals65.A is ChartAxisTitleFontSize)
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
					G(CS_0024_003C_003E8__locals65.A);
				}
				else if (CS_0024_003C_003E8__locals65.A is ChartTickLabelsFontSize)
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
					H(CS_0024_003C_003E8__locals65.A);
				}
				else if (!(CS_0024_003C_003E8__locals65.A is ChartDataLabelsFontSize) && CS_0024_003C_003E8__locals65.A is ChartDataLabelFontSize)
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
				}
			}
			goto IL_053a;
			IL_053a:
			((BaseError)CS_0024_003C_003E8__locals65.A).IsFixed = true;
			if (((BaseError)CS_0024_003C_003E8__locals65.A).CanFixMultiple)
			{
				TC a2 = default(TC);
				double top = default(double);
				double top2 = default(double);
				while (true)
				{
					switch (6)
					{
					case 0:
						continue;
					}
					if (DismissedFixAllTypes == null)
					{
						DismissedFixAllTypes = new List<Type>();
					}
					if (DismissedFixAllTypes.Contains(((object)CS_0024_003C_003E8__locals65.A).GetType()))
					{
						break;
					}
					try
					{
						ErrorType type2 = CS_0024_003C_003E8__locals65.A.Type;
						List<BaseError> list;
						if (type2 == ErrorType.CorporateDictionary)
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
							TC CS_0024_003C_003E8__locals63 = new TC(a2);
							CS_0024_003C_003E8__locals63.A = CS_0024_003C_003E8__locals65;
							CS_0024_003C_003E8__locals63.A = ((CorporateDictionary)CS_0024_003C_003E8__locals63.A.A).RuleId;
							list = Main.Analysis.Errors.Where([SpecialName] (BaseError baseError) =>
							{
								if ((object)((object)baseError).GetType() == ((object)CS_0024_003C_003E8__locals63.A.A).GetType())
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
									if (baseError != CS_0024_003C_003E8__locals63.A.A)
									{
										while (true)
										{
											switch (1)
											{
											case 0:
												break;
											default:
												return Operators.CompareString(((CorporateDictionary)baseError).RuleId, CS_0024_003C_003E8__locals63.A, TextCompare: false) == 0;
											}
										}
									}
								}
								return false;
							}).ToList();
						}
						else
						{
							list = Main.Analysis.Errors.Where([SpecialName] (BaseError baseError) =>
							{
								if ((object)((object)baseError).GetType() == ((object)CS_0024_003C_003E8__locals65.A).GetType())
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
											return baseError != CS_0024_003C_003E8__locals65.A;
										}
									}
								}
								return false;
							}).ToList();
						}
						int count = list.Count;
						if (count <= 0)
						{
							break;
						}
						while (true)
						{
							switch (4)
							{
							case 0:
								continue;
							}
							Callout.DoNotClose = true;
							if (Callout.Dialog != null)
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
								top = Callout.Dialog.Top;
								Callout.Dialog.Top = -10000.0;
							}
							if (Callout.MarchingAnts != null)
							{
								top2 = ((Window)(object)Callout.MarchingAnts).Top;
								((Window)(object)Callout.MarchingAnts).Top = -10000.0;
							}
							if (System.Windows.Forms.MessageBox.Show(AH.A(56352) + count + AH.A(56419), AH.A(5874), MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) == DialogResult.Yes)
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
								try
								{
									foreach (BaseError item in list)
									{
										((BaseError)item).FixAction();
										((BaseError)item).FixAction(B);
										((BaseError)item).IsFixed = true;
										Pane.TaskPane.WarningsView.RemoveItem(item, blnAnimate: false);
									}
									Pane.TaskPane.WarningsView.RemoveItemAndNavigate(CS_0024_003C_003E8__locals65.A);
								}
								catch (Exception ex3)
								{
									ProjectData.SetProjectError(ex3);
									Exception ex4 = ex3;
									System.Windows.Forms.MessageBox.Show(AH.A(56456) + ex4.Message, AH.A(5874), MessageBoxButtons.OK, MessageBoxIcon.Hand);
									ProjectData.ClearProjectError();
								}
							}
							else
							{
								DismissedFixAllTypes.Add(((object)CS_0024_003C_003E8__locals65.A).GetType());
							}
							if (Callout.Dialog != null)
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
								Callout.Dialog.Top = top;
							}
							if (Callout.MarchingAnts != null)
							{
								((Window)(object)Callout.MarchingAnts).Top = top2;
							}
							Pane.F();
							Callout.DoNotClose = false;
							break;
						}
						break;
					}
					catch (Exception ex5)
					{
						ProjectData.SetProjectError(ex5);
						Exception ex6 = ex5;
						ProjectData.ClearProjectError();
					}
					finally
					{
						List<BaseError> list = null;
					}
					break;
				}
			}
			goto end_IL_0013;
			IL_03cc:
			((BaseError)CS_0024_003C_003E8__locals65.A).FixAction(B);
			Fixes.A(CS_0024_003C_003E8__locals65.A);
			goto IL_053a;
			IL_02d4:
			((BaseError)CS_0024_003C_003E8__locals65.A).FixAction();
			goto IL_053a;
			IL_02b9:
			((BaseError)CS_0024_003C_003E8__locals65.A).FixAction();
			Fixes.A(CS_0024_003C_003E8__locals65.A);
			goto IL_053a;
			end_IL_0013:;
		}
		catch (Exception ex7)
		{
			ProjectData.SetProjectError(ex7);
			Exception ex8 = ex7;
			System.Windows.Forms.MessageBox.Show(AH.A(56571) + ex8.Message, AH.A(5874), MessageBoxButtons.OK, MessageBoxIcon.Hand);
			ProjectData.ClearProjectError();
		}
		CS_0024_003C_003E8__locals65.A = null;
	}

	private static void A(string A, bool B)
	{
		if (!B)
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
					Forms.WarningMessage(A);
					return;
				}
			}
		}
		Interaction.Beep();
	}

	private static void A(BaseError A)
	{
		List<Rect> list = new List<Rect>();
		Rect shapeRectangle;
		try
		{
			shapeRectangle = MarchingAnts.GetShapeRectangle(A.Shape);
			list.Add(shapeRectangle);
			Fixes.A(shapeRectangle.X, shapeRectangle.Y, list);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		list = null;
		shapeRectangle = default(Rect);
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
					Rect shapeRectangle = MarchingAnts.GetShapeRectangle(A.Shapes[i]);
					list.Add(shapeRectangle);
					if (shapeRectangle.X < num2)
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
						num2 = shapeRectangle.X;
					}
					if (shapeRectangle.Y < num)
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
						num = shapeRectangle.Y;
					}
					shapeRectangle = default(Rect);
				}
				while (true)
				{
					switch (4)
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
					Rect textFrameRectangle = MarchingAnts.GetTextFrameRectangle(A.Shapes[i]);
					list.Add(textFrameRectangle);
					if (textFrameRectangle.X < num2)
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
						num2 = textFrameRectangle.X;
					}
					if (textFrameRectangle.Y < num)
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
						num = textFrameRectangle.Y;
					}
					textFrameRectangle = default(Rect);
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
	}

	private static void D(BaseError A)
	{
		float num = 0f;
		float num2 = 0f;
		try
		{
			if (MarchingAnts.UseRelativePosition(A.Shape))
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
				num = A.Shape.Left;
				num2 = A.Shape.Top;
			}
			List<Rect> list = new List<Rect>();
			double num3 = 10000.0;
			double num4 = 10000.0;
			IEnumerator<TextRange2> enumerator = default(IEnumerator<TextRange2>);
			Rect textRangeRectangle;
			try
			{
				enumerator = ((BaseError)A).TextRanges.GetEnumerator();
				while (enumerator.MoveNext())
				{
					textRangeRectangle = MarchingAnts.GetTextRangeRectangle(enumerator.Current, num, num2);
					list.Add(textRangeRectangle);
					if (textRangeRectangle.X < num4)
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
						num4 = textRangeRectangle.X;
					}
					if (textRangeRectangle.Y < num3)
					{
						num3 = textRangeRectangle.Y;
					}
					textRangeRectangle = default(Rect);
				}
				while (true)
				{
					switch (4)
					{
					case 0:
						break;
					default:
						goto end_IL_00e2;
					}
					continue;
					end_IL_00e2:
					break;
				}
			}
			finally
			{
				if (enumerator != null)
				{
					while (true)
					{
						switch (6)
						{
						case 0:
							continue;
						}
						enumerator.Dispose();
						break;
					}
				}
			}
			textRangeRectangle = MarchingAnts.TextRangesTopLeft(((BaseError)A).TextRanges, num, num2);
			Callout.Dialog.XOffset = textRangeRectangle.X - num4;
			num4 = textRangeRectangle.X;
			num3 = textRangeRectangle.Y;
			Fixes.A(num4, num3, list);
			list = null;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	private static void E(BaseError A)
	{
		List<Rect> list = new List<Rect>();
		Rect chartTitleRectangle;
		try
		{
			chartTitleRectangle = MarchingAnts.GetChartTitleRectangle(A.ChartTitle, MarchingAnts.ChartLeftOffset(A.Shape), MarchingAnts.ChartTopOffset(A.Shape));
			list.Add(chartTitleRectangle);
			Fixes.A(chartTitleRectangle.X, chartTitleRectangle.Y, list);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		list = null;
		chartTitleRectangle = default(Rect);
	}

	private static void F(BaseError A)
	{
		List<Rect> list = new List<Rect>();
		Rect legendRectangle;
		try
		{
			legendRectangle = MarchingAnts.GetLegendRectangle(A.Legend, MarchingAnts.ChartLeftOffset(A.Shape), MarchingAnts.ChartTopOffset(A.Shape));
			list.Add(legendRectangle);
			Fixes.A(legendRectangle.X, legendRectangle.Y, list);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		list = null;
		legendRectangle = default(Rect);
	}

	private static void G(BaseError A)
	{
		List<Rect> list = new List<Rect>();
		Rect axisTitleRectangle;
		try
		{
			axisTitleRectangle = MarchingAnts.GetAxisTitleRectangle(A.AxisTitle, MarchingAnts.ChartLeftOffset(A.Shape), MarchingAnts.ChartTopOffset(A.Shape));
			list.Add(axisTitleRectangle);
			Fixes.A(axisTitleRectangle.X, axisTitleRectangle.Y, list);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		list = null;
		axisTitleRectangle = default(Rect);
	}

	private static void H(BaseError A)
	{
		List<Rect> list = new List<Rect>();
		Rect axisRectangle;
		try
		{
			axisRectangle = MarchingAnts.GetAxisRectangle(A.Axis, MarchingAnts.ChartLeftOffset(A.Shape), MarchingAnts.ChartTopOffset(A.Shape));
			list.Add(axisRectangle);
			Fixes.A(axisRectangle.X, axisRectangle.Y, list);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		list = null;
		axisRectangle = default(Rect);
	}

	private static void A(double A, double B, List<Rect> C)
	{
		Callout.A(Callout.Dialog, A, B);
		Fixes.B(C);
	}

	private static void B(List<Rect> A)
	{
		wpfMarchingAnts marchingAnts = Callout.MarchingAnts;
		marchingAnts.ClearMarchingAnts();
		marchingAnts.AddMarchingAnts(A);
		_ = null;
	}

	private static void I(BaseError A)
	{
		Microsoft.Office.Interop.PowerPoint.Application application = NG.A.Application;
		application.CommandBars.ExecuteMso(AH.A(40491));
		System.Windows.Forms.Application.DoEvents();
		BaseError baseError;
		BaseError activeItem;
		checked
		{
			if (A.Shape != null)
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
				if (A.Shape.HasChart == MsoTriState.msoTrue)
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
					int num = 0;
					if (A.Type == ErrorType.ChartLegendEntryMissing)
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
						num = ((ChartLegendEntryMissing)A).RequiredUndoSteps - 1;
					}
					else if (A.Type == ErrorType.ChartDataLabelNumberFormats)
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
						num = ((ChartDataLabelNumberFormats)A).RequiredUndoSteps - 1;
					}
					else if (A.Type == ErrorType.ChartDataLabelsInconsistent)
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
						num = ((ChartDataLabelsInconsistent)A).RequiredUndoSteps - 1;
					}
					int num2 = num;
					for (int i = 1; i <= num2; i++)
					{
						application.CommandBars.ExecuteMso(AH.A(40491));
						System.Windows.Forms.Application.DoEvents();
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
			}
			application = null;
			activeItem = Pane.ActiveItem;
			baseError = activeItem;
			ErrorType type = baseError.Type;
			if (type <= ErrorType.MultipleFontFamilies)
			{
				if (type == ErrorType.BulletPunctuation)
				{
					goto IL_0167;
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
				if (type == ErrorType.MultipleFontFamilies)
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
					goto IL_0262;
				}
			}
			else if (type == ErrorType.LineSpacing || type == ErrorType.ProofingLanguage)
			{
				goto IL_0167;
			}
			if (activeItem is BaseColorError)
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
				Fixes.A(Fixes.A, B: false);
			}
			goto IL_0262;
		}
		IL_0262:
		baseError = null;
		activeItem = null;
		B();
		((BaseError)A).IsFixed = false;
		return;
		IL_0167:
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
			goto IL_0262;
		}
	}

	private static void A()
	{
		try
		{
			Fixes.m_A = default(RC);
			Fixes.m_A.A = Callout.Dialog.Left;
			Fixes.m_A.B = Callout.Dialog.Top;
			Fixes.m_A.C = ((Window)(object)Callout.MarchingAnts).Left;
			Fixes.m_A.D = ((Window)(object)Callout.MarchingAnts).Top;
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
			((Window)(object)Callout.MarchingAnts).Left = Fixes.m_A.C;
			((Window)(object)Callout.MarchingAnts).Top = Fixes.m_A.D;
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
