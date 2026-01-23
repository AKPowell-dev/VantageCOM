using System;
using System.Collections;
using System.IO;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Timers;
using System.Windows.Forms;
using A;
using MacabacusMacros;
using MacabacusMacros.Auth;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1.Shapes;

public sealed class AirplaneMode
{
	[CompilerGenerated]
	internal sealed class VD
	{
		public Microsoft.Office.Interop.PowerPoint.ShapeRange A;

		public System.Timers.Timer A;

		public VD(VD A)
		{
			if (A != null)
			{
				this.A = A.A;
				this.A = A.A;
			}
		}

		[SpecialName]
		internal void A(object A, ElapsedEventArgs B)
		{
			IEnumerator enumerator = default(IEnumerator);
			try
			{
				enumerator = this.A.GetEnumerator();
				while (enumerator.MoveNext())
				{
					AirplaneMode.A((Microsoft.Office.Interop.PowerPoint.Shape)RuntimeHelpers.GetObjectValue(enumerator.Current));
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
			this.A = null;
			this.A = null;
		}
	}

	private static readonly string m_A = AH.A(70027);

	private static readonly string m_B = AH.A(70072);

	private static readonly string C = AH.A(70113);

	private static readonly string D = AH.A(70160);

	private static readonly string E = AH.A(70203);

	private static readonly string F = AH.A(70246);

	private static readonly string G = AH.A(70291);

	private static readonly string H = AH.A(70330);

	public static void Startup()
	{
		if (IsOn())
		{
			AddEvents(NG.A.Application);
		}
	}

	public static void Toggle(bool blnOn)
	{
		if (blnOn && !Access.AllowPowerPointOperation((PlanType)4, (Restriction)1, false))
		{
			KG.A.InvalidateControl(AH.A(69674));
			return;
		}
		Microsoft.Office.Interop.PowerPoint.Application application = NG.A.Application;
		if (blnOn)
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
			IEnumerator enumerator = default(IEnumerator);
			try
			{
				enumerator = application.Presentations.GetEnumerator();
				while (enumerator.MoveNext())
				{
					Microsoft.Office.Interop.PowerPoint.Presentation pres = (Microsoft.Office.Interop.PowerPoint.Presentation)enumerator.Current;
					try
					{
						HidePresentationImages(pres);
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
					switch (2)
					{
					case 0:
						break;
					default:
						goto end_IL_008a;
					}
					continue;
					end_IL_008a:
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
			AddEvents(application);
			clsReporting.LogActivity((ActivityApp)2, (ActivityCategory)14, AH.A(32975));
		}
		else
		{
			IEnumerator enumerator2 = default(IEnumerator);
			try
			{
				enumerator2 = application.Presentations.GetEnumerator();
				while (enumerator2.MoveNext())
				{
					Microsoft.Office.Interop.PowerPoint.Presentation a = (Microsoft.Office.Interop.PowerPoint.Presentation)enumerator2.Current;
					try
					{
						A(a);
					}
					catch (Exception ex3)
					{
						ProjectData.SetProjectError(ex3);
						Exception ex4 = ex3;
						ProjectData.ClearProjectError();
					}
				}
				while (true)
				{
					switch (2)
					{
					case 0:
						break;
					default:
						goto end_IL_0116;
					}
					continue;
					end_IL_0116:
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
			RemoveEvents(application);
		}
		application = null;
		PB.Settings.AirplaneMode = blnOn;
	}

	public static bool IsOn()
	{
		return PB.Settings.AirplaneMode;
	}

	public static void AddEvents(Microsoft.Office.Interop.PowerPoint.Application ppApp)
	{
		new ComAwareEventInfo(typeof(EApplication_Event), AH.A(69711)).AddEventHandler(ppApp, new EApplication_PresentationOpenEventHandler(HidePresentationImages));
		new ComAwareEventInfo(typeof(EApplication_Event), AH.A(58943)).AddEventHandler(ppApp, new EApplication_PresentationNewSlideEventHandler(HideSlideImages));
		new ComAwareEventInfo(typeof(EApplication_Event), AH.A(56688)).AddEventHandler(ppApp, new EApplication_PresentationBeforeCloseEventHandler(A));
	}

	public static void RemoveEvents(Microsoft.Office.Interop.PowerPoint.Application ppApp)
	{
		new ComAwareEventInfo(typeof(EApplication_Event), AH.A(69711)).RemoveEventHandler(ppApp, new EApplication_PresentationOpenEventHandler(HidePresentationImages));
		new ComAwareEventInfo(typeof(EApplication_Event), AH.A(58943)).RemoveEventHandler(ppApp, new EApplication_PresentationNewSlideEventHandler(HideSlideImages));
		new ComAwareEventInfo(typeof(EApplication_Event), AH.A(56688)).RemoveEventHandler(ppApp, new EApplication_PresentationBeforeCloseEventHandler(A));
	}

	private static void A(Microsoft.Office.Interop.PowerPoint.Presentation A, ref bool B)
	{
		if (!AirplaneMode.B(A))
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
			if (MessageBox.Show(AH.A(69744), AH.A(5874), MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation) == DialogResult.Cancel)
			{
				B = true;
			}
			return;
		}
	}

	public static void HidePresentationImages(Microsoft.Office.Interop.PowerPoint.Presentation pres)
	{
		A(pres, A);
	}

	private static void A(Microsoft.Office.Interop.PowerPoint.Presentation A)
	{
		AirplaneMode.A(A, Show);
	}

	private static void A(Microsoft.Office.Interop.PowerPoint.Presentation A, Action<Microsoft.Office.Interop.PowerPoint.Shape> B)
	{
		if (!AirplaneMode.B(A))
		{
			return;
		}
		IEnumerator enumerator = default(IEnumerator);
		IEnumerator enumerator2 = default(IEnumerator);
		IEnumerator enumerator3 = default(IEnumerator);
		IEnumerator enumerator4 = default(IEnumerator);
		IEnumerator enumerator5 = default(IEnumerator);
		IEnumerator enumerator6 = default(IEnumerator);
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
			try
			{
				enumerator = A.Slides.GetEnumerator();
				while (enumerator.MoveNext())
				{
					Slide slide = (Slide)enumerator.Current;
					try
					{
						enumerator2 = slide.Shapes.GetEnumerator();
						while (enumerator2.MoveNext())
						{
							AirplaneMode.A((Microsoft.Office.Interop.PowerPoint.Shape)enumerator2.Current, B);
						}
						while (true)
						{
							switch (6)
							{
							case 0:
								break;
							default:
								goto end_IL_006d;
							}
							continue;
							end_IL_006d:
							break;
						}
					}
					finally
					{
						if (enumerator2 is IDisposable)
						{
							while (true)
							{
								switch (2)
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
				while (true)
				{
					switch (5)
					{
					case 0:
						break;
					default:
						goto end_IL_00a1;
					}
					continue;
					end_IL_00a1:
					break;
				}
			}
			finally
			{
				if (enumerator is IDisposable)
				{
					while (true)
					{
						switch (6)
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
				enumerator3 = A.Designs.GetEnumerator();
				while (enumerator3.MoveNext())
				{
					Design design = (Design)enumerator3.Current;
					try
					{
						enumerator4 = design.SlideMaster.Shapes.GetEnumerator();
						while (enumerator4.MoveNext())
						{
							AirplaneMode.A((Microsoft.Office.Interop.PowerPoint.Shape)enumerator4.Current, B);
						}
						while (true)
						{
							switch (3)
							{
							case 0:
								break;
							default:
								goto end_IL_0122;
							}
							continue;
							end_IL_0122:
							break;
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
					try
					{
						enumerator5 = design.SlideMaster.CustomLayouts.GetEnumerator();
						while (enumerator5.MoveNext())
						{
							CustomLayout customLayout = (CustomLayout)enumerator5.Current;
							try
							{
								enumerator6 = customLayout.Shapes.GetEnumerator();
								while (enumerator6.MoveNext())
								{
									AirplaneMode.A((Microsoft.Office.Interop.PowerPoint.Shape)enumerator6.Current, B);
								}
							}
							finally
							{
								if (enumerator6 is IDisposable)
								{
									while (true)
									{
										switch (5)
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
						while (true)
						{
							switch (6)
							{
							case 0:
								break;
							default:
								goto end_IL_01d6;
							}
							continue;
							end_IL_01d6:
							break;
						}
					}
					finally
					{
						if (enumerator5 is IDisposable)
						{
							while (true)
							{
								switch (3)
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
				while (true)
				{
					switch (1)
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
				if (enumerator3 is IDisposable)
				{
					while (true)
					{
						switch (3)
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
	}

	private static void A(Microsoft.Office.Interop.PowerPoint.Shape A, Action<Microsoft.Office.Interop.PowerPoint.Shape> B)
	{
		if (A.Type != MsoShapeType.msoGroup)
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
					B(A);
					return;
				}
			}
		}
		foreach (Microsoft.Office.Interop.PowerPoint.Shape groupItem in A.GroupItems)
		{
			AirplaneMode.A(groupItem, B);
		}
	}

	public static void HideSlideImages(Slide sld)
	{
		if (!B((Microsoft.Office.Interop.PowerPoint.Presentation)sld.Parent))
		{
			return;
		}
		IEnumerator enumerator = default(IEnumerator);
		IEnumerator enumerator2 = default(IEnumerator);
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
			try
			{
				enumerator = sld.Shapes.GetEnumerator();
				while (enumerator.MoveNext())
				{
					A((Microsoft.Office.Interop.PowerPoint.Shape)enumerator.Current);
				}
				while (true)
				{
					switch (5)
					{
					case 0:
						break;
					default:
						goto end_IL_0054;
					}
					continue;
					end_IL_0054:
					break;
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
			try
			{
				enumerator2 = sld.CustomLayout.Shapes.GetEnumerator();
				while (enumerator2.MoveNext())
				{
					A((Microsoft.Office.Interop.PowerPoint.Shape)enumerator2.Current);
				}
				while (true)
				{
					switch (2)
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
						switch (5)
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

	private static void A(Microsoft.Office.Interop.PowerPoint.Shape A)
	{
		if (!Images.HasPictureOrGraphic(A))
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
			try
			{
				if (IsHidden(A))
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
					Microsoft.Office.Interop.PowerPoint.Shape shape = A;
					if (shape.Tags[G].Length == 0)
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
						if (shape.Tags[H].Length == 0)
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
							if (shape.Width / shape.Height < 15f)
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
								if (shape.Height / shape.Width < 15f)
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
									float num = clsPublish.PointsToInches(shape.Width);
									float num2 = clsPublish.PointsToInches(shape.Height);
									if (num < PB.Settings.AirplaneModeMaxWidth)
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
										if ((double)num2 < Conversions.ToDouble(PB.Settings.AirplaneModeMaxHeight))
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
											if (num > PB.Settings.AirplaneModeMinWidth)
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
												if (num2 > PB.Settings.AirplaneModeMinHeight)
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
													B(A);
												}
											}
										}
									}
								}
							}
							goto IL_019d;
						}
					}
					if (shape.Tags[H].Length > 0)
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
						B(A);
					}
					goto IL_019d;
					IL_019d:
					shape = null;
					return;
				}
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
				return;
			}
		}
	}

	private static void B(Microsoft.Office.Interop.PowerPoint.Shape A)
	{
		PictureFormat pictureFormat = A.PictureFormat;
		A.Tags.Add(AirplaneMode.m_A, pictureFormat.Brightness.ToString());
		A.Tags.Add(AirplaneMode.m_B, pictureFormat.Contrast.ToString());
		pictureFormat.Brightness = 0f;
		pictureFormat.Contrast = 0f;
		pictureFormat = null;
		if (!Images.HasGraphic(A))
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
					PictureFormat pictureFormat2 = A.PictureFormat;
					A.Tags.Add(C, ((int)pictureFormat2.TransparentBackground).ToString());
					pictureFormat2.TransparentBackground = MsoTriState.msoFalse;
					pictureFormat2 = null;
					Microsoft.Office.Interop.PowerPoint.FillFormat fill = A.Fill;
					if (fill.Visible == MsoTriState.msoTrue)
					{
						Microsoft.Office.Interop.PowerPoint.ColorFormat foreColor = fill.ForeColor;
						if (foreColor.RGB != 0)
						{
							A.Tags.Add(D, foreColor.RGB.ToString());
							foreColor.RGB = 0;
						}
						foreColor = null;
					}
					fill = null;
					return;
				}
				}
			}
		}
		Microsoft.Office.Interop.PowerPoint.LineFormat line = A.Line;
		if (line.Visible == MsoTriState.msoTrue)
		{
			A.Tags.Add(E, line.ForeColor.RGB.ToString());
			A.Tags.Add(F, line.Weight.ToString());
		}
		line.Visible = MsoTriState.msoTrue;
		line.ForeColor.RGB = 0;
		line.Weight = 500f;
		line = null;
	}

	public static void Show(Microsoft.Office.Interop.PowerPoint.Shape shp)
	{
		if (!Images.HasPictureOrGraphic(shp))
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
			try
			{
				if (!IsHidden(shp))
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
					Tags tags = shp.Tags;
					shp.PictureFormat.Brightness = Conversions.ToSingle(tags[AirplaneMode.m_A]);
					shp.PictureFormat.Contrast = Conversions.ToSingle(tags[AirplaneMode.m_B]);
					tags.Delete(AirplaneMode.m_A);
					tags.Delete(AirplaneMode.m_B);
					if (!Images.HasGraphic(shp))
					{
						shp.PictureFormat.TransparentBackground = (MsoTriState)Conversions.ToInteger(tags[C]);
						tags.Delete(C);
						if (tags[D].Length > 0)
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
							shp.Fill.ForeColor.RGB = Conversions.ToInteger(tags[D]);
							tags.Delete(D);
						}
					}
					else if (tags[F].Length > 0)
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
						shp.Line.Weight = Conversions.ToSingle(tags[F]);
						tags.Delete(F);
						if (tags[E].Length > 0)
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
							shp.Line.ForeColor.RGB = Conversions.ToInteger(tags[E]);
							tags.Delete(E);
						}
					}
					else
					{
						shp.Line.Weight = 0f;
						shp.Line.Visible = MsoTriState.msoFalse;
					}
					tags = null;
					return;
				}
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
				return;
			}
		}
	}

	public static bool IsHidden(Microsoft.Office.Interop.PowerPoint.Shape shp)
	{
		return shp.Tags[AirplaneMode.m_A].Length > 0;
	}

	private static bool A(Microsoft.Office.Interop.PowerPoint.Presentation A)
	{
		string extension = Path.GetExtension(A.Name);
		if (Operators.CompareString(extension, AH.A(69996), TextCompare: false) != 0)
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
			if (Operators.CompareString(extension, AH.A(70007), TextCompare: false) != 0)
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
				if (Operators.CompareString(extension, AH.A(70018), TextCompare: false) != 0)
				{
					return false;
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
		}
		return true;
	}

	private static bool B(Microsoft.Office.Interop.PowerPoint.Presentation A)
	{
		if (A.Windows.Count > 0)
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
					return !AirplaneMode.A(A);
				}
			}
		}
		return false;
	}

	public static void Peek()
	{
		VD a = default(VD);
		VD CS_0024_003C_003E8__locals9 = new VD(a);
		CS_0024_003C_003E8__locals9.A = Base.SelectedShapes();
		IEnumerator enumerator = CS_0024_003C_003E8__locals9.A.GetEnumerator();
		try
		{
			while (enumerator.MoveNext())
			{
				Show((Microsoft.Office.Interop.PowerPoint.Shape)enumerator.Current);
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
		CS_0024_003C_003E8__locals9.A = new System.Timers.Timer(PB.Settings.AirplaneModePeek);
		CS_0024_003C_003E8__locals9.A.Elapsed += [SpecialName] (object A, ElapsedEventArgs B) =>
		{
			IEnumerator enumerator2 = default(IEnumerator);
			try
			{
				enumerator2 = CS_0024_003C_003E8__locals9.A.GetEnumerator();
				while (enumerator2.MoveNext())
				{
					AirplaneMode.A((Microsoft.Office.Interop.PowerPoint.Shape)RuntimeHelpers.GetObjectValue(enumerator2.Current));
				}
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
						goto end_IL_002f;
					}
					continue;
					end_IL_002f:
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
							break;
						default:
							(enumerator2 as IDisposable).Dispose();
							goto end_IL_004c;
						}
						continue;
						end_IL_004c:
						break;
					}
				}
			}
			CS_0024_003C_003E8__locals9.A = null;
			CS_0024_003C_003E8__locals9.A = null;
		};
		CS_0024_003C_003E8__locals9.A.AutoReset = false;
		CS_0024_003C_003E8__locals9.A.Start();
	}

	public static void Exclude()
	{
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = Base.SelectedShapes().GetEnumerator();
			while (enumerator.MoveNext())
			{
				Microsoft.Office.Interop.PowerPoint.Shape obj = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator.Current;
				obj.Tags.Delete(H);
				obj.Tags.Add(G, AH.A(9078));
				Show(obj);
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
				return;
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
	}

	public static void Include()
	{
		IEnumerator enumerator = Base.SelectedShapes().GetEnumerator();
		try
		{
			while (enumerator.MoveNext())
			{
				Microsoft.Office.Interop.PowerPoint.Shape obj = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator.Current;
				obj.Tags.Delete(G);
				obj.Tags.Add(H, AH.A(9078));
				A(obj);
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
				return;
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
