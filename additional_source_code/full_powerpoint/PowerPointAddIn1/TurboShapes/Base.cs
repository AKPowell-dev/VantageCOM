using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Interop;
using System.Windows.Media.Imaging;
using System.Xml;
using A;
using MacabacusMacros;
using MacabacusMacros.UI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.Shapes;
using PowerPointAddIn1.Utilities;

namespace PowerPointAddIn1.TurboShapes;

public sealed class Base
{
	public enum TurboShapeType
	{
		Custom,
		HarveyBall,
		ProgressBar,
		RatingBar,
		CheckBox,
		NoticeIcon,
		TrafficLight,
		Thermometer,
		ToggleSwitch,
		Arrow,
		Dial,
		SliderBar,
		Tachometer
	}

	public enum TurboShapeColor
	{
		Primary = 1,
		Secondary,
		Red,
		Yellow,
		Green,
		Blue
	}

	public enum CalloutPosition
	{
		TopCenter = 1,
		TopLeft
	}

	private enum DG
	{
		A = 88,
		B = 90,
		C = 10,
		D = 8,
		E = 117,
		F = 118
	}

	public static readonly string TAG_TYPE = AH.A(158535);

	public static readonly string TAG_VALUE = AH.A(158564);

	public static readonly int TIMER_CHK_POSN_INTERVAL = 200;

	public static readonly int NOTCH_OFFSET = 19;

	public static readonly int TOP_OFFSET = 2;

	[CompilerGenerated]
	private static Dictionary<TurboShapeColor, int> m_A;

	[CompilerGenerated]
	private static Window m_A;

	private static Dictionary<TurboShapeColor, int> ColorPalette
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

	public static Window ActiveCallout
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

	private static void A()
	{
		XmlDocument settingsXml = KG.A.SettingsXml;
		ColorPalette = new Dictionary<TurboShapeColor, int>();
		Dictionary<TurboShapeColor, int> colorPalette = ColorPalette;
		colorPalette.Add(TurboShapeColor.Primary, clsColors.RGB2Ole(settingsXml.SelectSingleNode(AH.A(158200)).InnerText));
		colorPalette.Add(TurboShapeColor.Secondary, clsColors.RGB2Ole(settingsXml.SelectSingleNode(AH.A(158253)).InnerText));
		colorPalette.Add(TurboShapeColor.Red, clsColors.RGB2Ole(settingsXml.SelectSingleNode(AH.A(158310)).InnerText));
		colorPalette.Add(TurboShapeColor.Yellow, clsColors.RGB2Ole(settingsXml.SelectSingleNode(AH.A(158355)).InnerText));
		colorPalette.Add(TurboShapeColor.Green, clsColors.RGB2Ole(settingsXml.SelectSingleNode(AH.A(158406)).InnerText));
		colorPalette.Add(TurboShapeColor.Blue, clsColors.RGB2Ole(settingsXml.SelectSingleNode(AH.A(158455)).InnerText));
		_ = null;
		settingsXml = null;
	}

	public static int GetColor(TurboShapeColor clr)
	{
		if (ColorPalette == null)
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
			A();
		}
		return ColorPalette[clr];
	}

	public static void ResetColors()
	{
		ColorPalette = null;
	}

	public static void ImportPrimaryColor(Microsoft.Office.Interop.PowerPoint.Shape shpNew, Microsoft.Office.Interop.PowerPoint.Shape shpOld)
	{
		int num = -1;
		int color = GetColor(TurboShapeColor.Secondary);
		Microsoft.Office.Interop.PowerPoint.Shape shape = shpOld;
		try
		{
			if (shape.Type == MsoShapeType.msoGroup)
			{
				IEnumerator enumerator = default(IEnumerator);
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
					try
					{
						enumerator = shape.GroupItems.GetEnumerator();
						while (true)
						{
							if (enumerator.MoveNext())
							{
								Microsoft.Office.Interop.PowerPoint.ColorFormat foreColor = ((Microsoft.Office.Interop.PowerPoint.Shape)enumerator.Current).Fill.ForeColor;
								if (foreColor.RGB != color)
								{
									while (true)
									{
										switch (2)
										{
										case 0:
											continue;
										}
										num = foreColor.RGB;
										break;
									}
									break;
								}
								foreColor = null;
								continue;
							}
							while (true)
							{
								switch (2)
								{
								case 0:
									break;
								default:
									goto end_IL_0080;
								}
								continue;
								end_IL_0080:
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
								switch (3)
								{
								case 0:
									continue;
								}
								(enumerator as IDisposable).Dispose();
								break;
							}
						}
					}
					break;
				}
			}
			else
			{
				Microsoft.Office.Interop.PowerPoint.ColorFormat foreColor2 = shape.Fill.ForeColor;
				if (foreColor2.RGB != color)
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
					num = foreColor2.RGB;
				}
				foreColor2 = null;
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		if (num > 0)
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
			if (shpNew.Type == MsoShapeType.msoGroup)
			{
				IEnumerator enumerator2 = default(IEnumerator);
				try
				{
					enumerator2 = shpNew.GroupItems.GetEnumerator();
					while (true)
					{
						if (enumerator2.MoveNext())
						{
							Microsoft.Office.Interop.PowerPoint.ColorFormat foreColor3 = ((Microsoft.Office.Interop.PowerPoint.Shape)enumerator2.Current).Fill.ForeColor;
							if (foreColor3.RGB != color)
							{
								while (true)
								{
									switch (7)
									{
									case 0:
										continue;
									}
									foreColor3.RGB = num;
									break;
								}
								break;
							}
							foreColor3 = null;
							continue;
						}
						while (true)
						{
							switch (2)
							{
							case 0:
								break;
							default:
								goto end_IL_016a;
							}
							continue;
							end_IL_016a:
							break;
						}
						break;
					}
				}
				finally
				{
					if (enumerator2 is IDisposable)
					{
						while (true)
						{
							switch (7)
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
			else if (shpNew.Fill.ForeColor.RGB != color)
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
				shpNew.Fill.ForeColor.RGB = num;
				shpNew.Fill.BackColor.RGB = num;
			}
		}
		shape = null;
	}

	[DllImport("gdi32", CharSet = CharSet.Ansi, EntryPoint = "GetDeviceCaps", ExactSpelling = true, SetLastError = true)]
	private static extern int A(IntPtr A, int B);

	[DllImport("user32", CharSet = CharSet.Ansi, EntryPoint = "GetDC", ExactSpelling = true, SetLastError = true)]
	private static extern IntPtr A(IntPtr A);

	[DllImport("user32", CharSet = CharSet.Ansi, EntryPoint = "ReleaseDC", ExactSpelling = true, SetLastError = true)]
	private static extern bool A(IntPtr A, IntPtr B);

	private static Rectangle A(Microsoft.Office.Interop.PowerPoint.Shape A)
	{
		Graphics graphics = Graphics.FromHwnd(IntPtr.Zero);
		try
		{
			DocumentWindow activeWindow = NG.A.Application.ActiveWindow;
			int num = activeWindow.PointsToScreenPixelsX(0f);
			int num2 = activeWindow.PointsToScreenPixelsY(0f);
			double num3 = (double)activeWindow.View.Zoom / 100.0;
			IntPtr hdc = graphics.GetHdc();
			double num4 = (double)Base.A(hdc, 88) / 72.0;
			double num5 = (double)Base.A(hdc, 90) / 72.0;
			graphics.ReleaseHdc(hdc);
			double num6 = 96f / graphics.DpiY;
			return new Rectangle
			{
				X = Convert.ToInt32(((double)num + (double)A.Left * num4 * num3) * num6),
				Y = Convert.ToInt32(((double)num2 + (double)A.Top * num5 * num3) * num6),
				Width = Convert.ToInt32(((double)num + (double)A.Width * num4 * num3) * num6),
				Height = Convert.ToInt32(((double)num2 + (double)A.Height * num5 * num3) * num6)
			};
		}
		finally
		{
			if (graphics != null)
			{
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
					((IDisposable)graphics).Dispose();
					break;
				}
			}
		}
	}

	public static void TransformFromShape(Microsoft.Office.Interop.PowerPoint.Shape shp, CalloutPosition pos, ref double unitX, ref double unitY)
	{
		Rectangle rectangle = Dialogs.A(shp);
		int num = rectangle.Left;
		checked
		{
			if (pos == CalloutPosition.TopCenter)
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
				num += (int)Math.Round((double)rectangle.Width / 2.0);
			}
			Dialogs.B(ref unitX, ref unitY, num, rectangle.Top + TOP_OFFSET);
			if (pos == CalloutPosition.TopCenter)
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
				unitX -= NOTCH_OFFSET;
			}
			rectangle = default(Rectangle);
		}
	}

	public static void AddSelectionChangedEvent()
	{
		new ComAwareEventInfo(typeof(EApplication_Event), AH.A(12762)).AddEventHandler(NG.A.Application, new EApplication_WindowSelectionChangeEventHandler(A));
	}

	public static void RemoveSelectionChangedEvent()
	{
		new ComAwareEventInfo(typeof(EApplication_Event), AH.A(12762)).RemoveEventHandler(NG.A.Application, new EApplication_WindowSelectionChangeEventHandler(A));
	}

	private static void A(Selection A)
	{
		if (ActiveCallout != null)
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
			if (ActiveCallout.IsActive)
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
				break;
			}
		}
		ActiveCallout = null;
		checked
		{
			Microsoft.Office.Interop.PowerPoint.ShapeRange shapeRange;
			Microsoft.Office.Interop.PowerPoint.Shape shape;
			try
			{
				if (A.Type == PpSelectionType.ppSelectionShapes)
				{
					shapeRange = PowerPointAddIn1.Shapes.Base.SelectedShapes(A);
					if (shapeRange.Count == 1)
					{
						while (true)
						{
							switch (5)
							{
							case 0:
								continue;
							}
							shape = shapeRange[1];
							string text = shape.Tags[TAG_TYPE];
							if (text.Length > 0)
							{
								float num = Conversions.ToSingle(shape.Tags[TAG_VALUE].ToString());
								switch (unchecked((TurboShapeType)Conversions.ToInteger(text)))
								{
								case TurboShapeType.HarveyBall:
									HarveyBall.Edit(shape, (int)Math.Round(num));
									break;
								case TurboShapeType.ProgressBar:
									ProgressBar.Edit(shape, (int)Math.Round(num));
									break;
								case TurboShapeType.RatingBar:
									RatingBar.Edit(shape, num);
									break;
								case TurboShapeType.CheckBox:
									CheckBox.Edit(shape, (int)Math.Round(num));
									break;
								case TurboShapeType.NoticeIcon:
									NoticeIcon.Edit(shape, (int)Math.Round(num));
									break;
								case TurboShapeType.Thermometer:
									Thermometer.Edit(shape, (int)Math.Round(num));
									break;
								case TurboShapeType.Tachometer:
									Tachometer.Edit(shape, (int)Math.Round(num));
									break;
								case TurboShapeType.TrafficLight:
									TrafficLight.Edit(shape, (int)Math.Round(num));
									break;
								case TurboShapeType.ToggleSwitch:
									ToggleSwitch.Edit(shape, (int)Math.Round(num));
									break;
								case TurboShapeType.Arrow:
									Arrow.Edit(shape, (int)Math.Round(num));
									break;
								case TurboShapeType.SliderBar:
									SliderBar.Edit(shape, (int)Math.Round(num));
									break;
								case TurboShapeType.Custom:
									Custom.Edit(shape, (int)Math.Round(num));
									break;
								case TurboShapeType.Dial:
									break;
								}
							}
							break;
						}
					}
				}
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			shapeRange = null;
			shape = null;
		}
	}

	public static void AddTurboShape(Action<Slide, PageSetup> a)
	{
		if (!Licensing.AllowRestrictedMode())
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
			try
			{
				Microsoft.Office.Interop.PowerPoint.Application application = NG.A.Application;
				Slide arg = application.ActiveWindow.Selection.SlideRange[1];
				PageSetup pageSetup = application.ActivePresentation.PageSetup;
				application.StartNewUndoEntry();
				_ = null;
				try
				{
					a(arg, pageSetup);
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					Forms.ErrorMessage(ex2.Message);
					LogException(ex2);
					ProjectData.ClearProjectError();
				}
				arg = null;
				pageSetup = null;
				return;
			}
			catch (Exception ex3)
			{
				ProjectData.SetProjectError(ex3);
				Exception ex4 = ex3;
				Forms.WarningMessage(AH.A(158502));
				ProjectData.ClearProjectError();
				return;
			}
		}
	}

	public static Microsoft.Office.Interop.PowerPoint.Shape MergeShapes(Slide sld, List<string> listShapes)
	{
		sld.Shapes.Range(listShapes.ToArray()).MergeShapes(MsoMergeCmd.msoMergeUnion);
		return sld.Shapes[sld.Shapes.Count];
	}

	public static Microsoft.Office.Interop.PowerPoint.Shape CombineShapes(Slide sld, List<string> listShapes)
	{
		sld.Shapes.Range(listShapes.ToArray()).MergeShapes(MsoMergeCmd.msoMergeCombine);
		return sld.Shapes[sld.Shapes.Count];
	}

	public static Microsoft.Office.Interop.PowerPoint.Shape SubtractShapes(Slide sld, List<string> listShapes)
	{
		sld.Shapes.Range(listShapes.ToArray()).MergeShapes(MsoMergeCmd.msoMergeSubtract);
		return sld.Shapes[sld.Shapes.Count];
	}

	public static void FinalizeShape(Microsoft.Office.Interop.PowerPoint.Shape shp, TurboShapeType type, float val, string strName)
	{
		shp.LockAspectRatio = MsoTriState.msoTrue;
		Tags tags = shp.Tags;
		string tAG_TYPE = TAG_TYPE;
		int num = (int)type;
		tags.Add(tAG_TYPE, num.ToString());
		shp.Tags.Add(TAG_VALUE, val.ToString());
		shp.Name = strName;
		shp.Select();
		_ = null;
	}

	public static void LogException(Exception ex)
	{
		clsReporting.LogException(ex);
	}

	public static void LogActivity(string strActivity)
	{
		clsReporting.LogActivity((ActivityApp)2, (ActivityCategory)6, strActivity);
	}

	public static void SetImageSource(System.Windows.Controls.Image img, Bitmap bmp)
	{
		img.Source = Imaging.CreateBitmapSourceFromHBitmap(bmp.GetHbitmap(), IntPtr.Zero, Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions());
	}

	public static Slide GetSlide(Microsoft.Office.Interop.PowerPoint.Shape shp)
	{
		return ((Microsoft.Office.Interop.PowerPoint.Application)shp.Application).ActiveWindow.Selection.SlideRange[1];
	}
}
