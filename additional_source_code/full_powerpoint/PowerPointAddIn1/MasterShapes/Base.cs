using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using A;
using MacabacusMacros;
using MacabacusMacros.UI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1.MasterShapes;

public sealed class Base
{
	[Serializable]
	[CompilerGenerated]
	internal sealed class _Closure_0024__
	{
		public static readonly _Closure_0024__ A;

		public static Func<Microsoft.Office.Interop.PowerPoint.Shape, Behavior> A;

		public static Func<Microsoft.Office.Interop.PowerPoint.Shape, int> A;

		static _Closure_0024__()
		{
			_Closure_0024__.A = new _Closure_0024__();
		}

		[SpecialName]
		internal Behavior A(Microsoft.Office.Interop.PowerPoint.Shape A)
		{
			return Base.A(A.Name);
		}

		[SpecialName]
		internal int A(Microsoft.Office.Interop.PowerPoint.Shape A)
		{
			return A.ZOrderPosition;
		}
	}

	[CompilerGenerated]
	private static string m_A = AH.A(151454);

	[CompilerGenerated]
	private static string m_B = AH.A(151483);

	[CompilerGenerated]
	private static Dictionary<string, Microsoft.Office.Interop.PowerPoint.Shape> m_A;

	internal static string TAG_ID
	{
		[CompilerGenerated]
		get
		{
			return Base.m_A;
		}
	}

	internal static string REGEX_PATTERN
	{
		[CompilerGenerated]
		get
		{
			return Base.m_B;
		}
	}

	internal static Dictionary<string, Microsoft.Office.Interop.PowerPoint.Shape> MyMasterShapes
	{
		[CompilerGenerated]
		get
		{
			return Base.m_A;
		}
		[CompilerGenerated]
		set
		{
			Base.m_A = value;
		}
	} = null;

	internal static void A()
	{
	}

	internal static void B()
	{
		Microsoft.Office.Interop.PowerPoint.Application application = NG.A.Application;
		List<string> list = new List<string>();
		try
		{
			application.ActiveWindow.ViewType = PpViewType.ppViewMasterThumbnails;
			Microsoft.Office.Interop.PowerPoint.Shapes shapes = application.ActivePresentation.Designs[1].SlideMaster.Shapes;
			IEnumerator enumerator = default(IEnumerator);
			try
			{
				enumerator = shapes.GetEnumerator();
				while (enumerator.MoveNext())
				{
					Microsoft.Office.Interop.PowerPoint.Shape shape = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator.Current;
					if (A(shape))
					{
						shape.Visible = MsoTriState.msoTrue;
						list.Add(shape.Name);
					}
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
						if (1 == 0)
						{
							/*OpCode not supported: LdMemberToken*/;
						}
						(enumerator as IDisposable).Dispose();
						break;
					}
				}
			}
			if (list.Any())
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
				application.CommandBars.ExecuteMso(AH.A(91479));
				System.Windows.Forms.Application.DoEvents();
			}
			Forms.WarningMessage(AH.A(151055));
			Base.A(AH.A(151201));
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

	internal static void C()
	{
		List<Microsoft.Office.Interop.PowerPoint.Shape> list = new List<Microsoft.Office.Interop.PowerPoint.Shape>();
		MyMasterShapes = new Dictionary<string, Microsoft.Office.Interop.PowerPoint.Shape>();
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = NG.A.Application.ActivePresentation.Designs[1].SlideMaster.Shapes.GetEnumerator();
			while (enumerator.MoveNext())
			{
				Microsoft.Office.Interop.PowerPoint.Shape shape = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator.Current;
				Microsoft.Office.Interop.PowerPoint.Shape shape2 = shape;
				if (shape2.Visible == MsoTriState.msoFalse)
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
					if (A(shape) && shape2.Type != MsoShapeType.msoPlaceholder && Operators.CompareString(shape2.Name, TextBox.TEXTBOX_NAME, TextCompare: false) != 0)
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
						list.Add(shape);
					}
				}
				shape2 = null;
			}
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					goto end_IL_00b8;
				}
				continue;
				end_IL_00b8:
				break;
			}
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
						continue;
					}
					(enumerator as IDisposable).Dispose();
					break;
				}
			}
		}
		if (list.Any())
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
			List<Microsoft.Office.Interop.PowerPoint.Shape> source = list;
			Func<Microsoft.Office.Interop.PowerPoint.Shape, Behavior> keySelector;
			if (_Closure_0024__.A == null)
			{
				keySelector = (_Closure_0024__.A = [SpecialName] (Microsoft.Office.Interop.PowerPoint.Shape A) => Base.A(A.Name));
			}
			else
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
				keySelector = _Closure_0024__.A;
			}
			list = source.OrderBy(keySelector).ThenBy([SpecialName] (Microsoft.Office.Interop.PowerPoint.Shape A) => A.ZOrderPosition).ToList();
			using List<Microsoft.Office.Interop.PowerPoint.Shape>.Enumerator enumerator2 = list.GetEnumerator();
			while (enumerator2.MoveNext())
			{
				Microsoft.Office.Interop.PowerPoint.Shape current = enumerator2.Current;
				MyMasterShapes.Add(current.Id.ToString(), current);
			}
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					goto end_IL_01a2;
				}
				continue;
				end_IL_01a2:
				break;
			}
		}
		list = null;
	}

	internal static Behavior A(string A)
	{
		Behavior result = Behavior.SelectedSlides;
		Match match = Regex.Match(A, REGEX_PATTERN);
		if (match.Success)
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
			string text = match.Groups[1].ToString();
			uint num = YG.A(text);
			if (num <= 1436297548)
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
				if (num <= 769487330)
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
					if (num <= 551378283)
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
						if (num != 517278592)
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
							if (num == 551378283)
							{
								if (Operators.CompareString(text, AH.A(151261), TextCompare: false) != 0)
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
								}
								else
								{
									result = Behavior.LayoutsShowingBackgroundGraphics;
								}
							}
						}
						else if (Operators.CompareString(text, AH.A(151241), TextCompare: false) != 0)
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
						}
						else
						{
							result = Behavior.AllLayouts;
						}
					}
					else if (num != 620754425)
					{
						if (num != 634280640)
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
							if (num != 769487330)
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
							}
							else if (Operators.CompareString(text, AH.A(151256), TextCompare: false) == 0)
							{
								result = Behavior.SlidesShowingBackgroundGraphics;
							}
						}
						else if (Operators.CompareString(text, AH.A(151266), TextCompare: false) != 0)
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
							result = Behavior.DynamicSlides;
						}
					}
					else if (Operators.CompareString(text, AH.A(151281), TextCompare: false) != 0)
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
					else
					{
						result = Behavior.SpecialLayouts;
					}
				}
				else if (num <= 1037384781)
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
					if (num != 892207104)
					{
						if (num != 1037384781)
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
						}
						else if (Operators.CompareString(text, AH.A(151236), TextCompare: false) != 0)
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
						}
						else
						{
							result = Behavior.AllSlides;
						}
					}
					else if (Operators.CompareString(text, AH.A(151342), TextCompare: false) != 0)
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
					}
					else
					{
						result = Behavior.CenterInShape;
					}
				}
				else if (num != 1040194900)
				{
					if (num != 1154386829)
					{
						if (num != 1436297548)
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
						}
						else if (Operators.CompareString(text, AH.A(151286), TextCompare: false) != 0)
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
						}
						else
						{
							result = Behavior.AboveTopLeft;
						}
					}
					else if (Operators.CompareString(text, AH.A(151271), TextCompare: false) != 0)
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
					}
					else
					{
						result = Behavior.DynamicLayouts;
					}
				}
				else if (Operators.CompareString(text, AH.A(151276), TextCompare: false) != 0)
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
				}
				else
				{
					result = Behavior.SpecialSlides;
				}
			}
			else if (num <= 2675227691u)
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
				if (num <= 1711446755)
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
					if (num != 1708783731)
					{
						if (num != 1711446755)
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
						else if (Operators.CompareString(text, AH.A(8070), TextCompare: false) == 0)
						{
							result = Behavior.SelectedSlides;
						}
					}
					else if (Operators.CompareString(text, AH.A(151246), TextCompare: false) != 0)
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
					}
					else
					{
						result = Behavior.ContentSlides;
					}
				}
				else if (num != 1805405166)
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
					if (num != 2195334682u)
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
						if (num != 2675227691u)
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
						}
						else if (Operators.CompareString(text, AH.A(151307), TextCompare: false) == 0)
						{
							result = Behavior.BelowBottomLeft;
						}
					}
					else if (Operators.CompareString(text, AH.A(151251), TextCompare: false) != 0)
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
					}
					else
					{
						result = Behavior.ContentLayouts;
					}
				}
				else if (Operators.CompareString(text, AH.A(151293), TextCompare: false) != 0)
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
				}
				else
				{
					result = Behavior.AboveTopRight;
				}
			}
			else if (num <= 3266897280u)
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
				if (num != 2910114357u)
				{
					if (num != 3266897280u)
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
					else if (Operators.CompareString(text, AH.A(151328), TextCompare: false) != 0)
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
					}
					else
					{
						result = Behavior.InsideBottomRight;
					}
				}
				else if (Operators.CompareString(text, AH.A(151300), TextCompare: false) != 0)
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
				}
				else
				{
					result = Behavior.BelowBottomRight;
				}
			}
			else if (num != 3268574566u)
			{
				if (num != 3704792660u)
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
					if (num != 3770225850u)
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
					}
					else if (Operators.CompareString(text, AH.A(151335), TextCompare: false) != 0)
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
					}
					else
					{
						result = Behavior.InsideBottomLeft;
					}
				}
				else if (Operators.CompareString(text, AH.A(151314), TextCompare: false) == 0)
				{
					result = Behavior.InsideTopLeft;
				}
			}
			else if (Operators.CompareString(text, AH.A(151321), TextCompare: false) != 0)
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
			}
			else
			{
				result = Behavior.InsideTopRight;
			}
			match = null;
		}
		return result;
	}

	internal static void A(Microsoft.Office.Interop.PowerPoint.Application A)
	{
		DocumentWindow activeWindow = A.ActiveWindow;
		if (activeWindow.Selection.Type == PpSelectionType.ppSelectionNone)
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
				if (activeWindow.ViewType != PpViewType.ppViewSlideSorter)
				{
					while (true)
					{
						switch (4)
						{
						case 0:
							continue;
						}
						if (activeWindow.Panes.Count <= 2)
						{
							break;
						}
						while (true)
						{
							switch (2)
							{
							case 0:
								continue;
							}
							if (activeWindow.Panes[3].ViewType != PpViewType.ppViewNotesPage)
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
								activeWindow.Panes[3].Activate();
								break;
							}
							break;
						}
						break;
					}
				}
				else
				{
					((Slide)activeWindow.View.Slide).Select();
				}
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
		}
		activeWindow = null;
	}

	internal static bool A(Microsoft.Office.Interop.PowerPoint.Shape A, string B)
	{
		return Operators.CompareString(A.Tags[TAG_ID], B, TextCompare: false) == 0;
	}

	internal static bool A(Microsoft.Office.Interop.PowerPoint.Shape A)
	{
		return Regex.IsMatch(A.Name, REGEX_PATTERN);
	}

	internal static bool A(Microsoft.Office.Interop.PowerPoint.Application A, bool B)
	{
		bool result;
		try
		{
			if (A.ActiveWindow.Panes[2].ViewType == PpViewType.ppViewSlideMaster)
			{
				if (B)
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
					Forms.WarningMessage(AH.A(151349));
				}
				result = true;
			}
			else
			{
				result = false;
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			result = false;
			ProjectData.ClearProjectError();
		}
		return result;
	}

	internal static void A(Exception A)
	{
		Forms.ErrorMessage(A.Message);
		clsReporting.LogException(A);
	}

	internal static void A(string A)
	{
		clsReporting.LogActivity((ActivityApp)2, (ActivityCategory)11, A);
	}
}
