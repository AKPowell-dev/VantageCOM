using System;
using System.Collections;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Threading;
using A;
using MacabacusMacros;
using MacabacusMacros.UI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.Shapes;
using PowerPointAddIn1.Slides;

namespace PowerPointAddIn1.MasterShapes;

public sealed class AddRemove
{
	public struct Master
	{
		public string Id;

		public string Name;

		public string PlaceholderText;

		public float Left;

		public float Top;
	}

	[CompilerGenerated]
	internal sealed class TF
	{
		public Microsoft.Office.Interop.PowerPoint.Shape A;

		public TF(TF A)
		{
			if (A == null)
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
				this.A = A.A;
				return;
			}
		}

		[SpecialName]
		internal void A()
		{
			this.A.Copy();
		}
	}

	[CompilerGenerated]
	internal sealed class UF
	{
		public Microsoft.Office.Interop.PowerPoint.Shape A;

		public UF(UF A)
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
		internal void A()
		{
			this.A.Copy();
		}
	}

	public static void Insert(IRibbonControl control)
	{
		TF a = default(TF);
		TF CS_0024_003C_003E8__locals7 = new TF(a);
		if (!Licensing.AllowMasterShapesOperation())
		{
			return;
		}
		CS_0024_003C_003E8__locals7.A = null;
		string tag = control.Tag;
		Application application = NG.A.Application;
		if (application.ActiveWindow.Selection.Type != PpSelectionType.ppSelectionShapes)
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
					Forms.WarningMessage(AH.A(139883));
					application = null;
					return;
				}
			}
		}
		if (!Base.MyMasterShapes.TryGetValue(tag, out CS_0024_003C_003E8__locals7.A))
		{
			return;
		}
		IEnumerator enumerator = default(IEnumerator);
		while (true)
		{
			switch (2)
			{
			case 0:
				continue;
			}
			Microsoft.Office.Interop.PowerPoint.Presentation activePresentation;
			Microsoft.Office.Interop.PowerPoint.ShapeRange shapeRange;
			Microsoft.Office.Interop.PowerPoint.Shape b;
			object objectValue;
			List<int> list;
			try
			{
				Master d = MasterShapeProperties(CS_0024_003C_003E8__locals7.A);
				Behavior g = Base.A(d.Name);
				activePresentation = application.ActivePresentation;
				shapeRange = application.ActiveWindow.Selection.ShapeRange;
				objectValue = RuntimeHelpers.GetObjectValue(shapeRange[1].Parent);
				list = new List<int>();
				application.StartNewUndoEntry();
				DateTime now = DateTime.Now;
				if (clsClipboard.CopyWithWait((Action)([SpecialName] () =>
				{
					CS_0024_003C_003E8__locals7.A.Copy();
				}), 4000))
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
					while (true)
					{
						if (objectValue is Slide)
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
							Slide slide = (Slide)objectValue;
							try
							{
								enumerator = shapeRange.GetEnumerator();
								while (enumerator.MoveNext())
								{
									Microsoft.Office.Interop.PowerPoint.Shape c = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator.Current;
									b = A(slide, CS_0024_003C_003E8__locals7.A, c, d, activePresentation, now, g);
									list.Add(PowerPointAddIn1.Shapes.Helpers.A(slide, b));
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
							slide.Shapes.Range(list.ToArray()).Select();
							slide = null;
							break;
						}
						if (objectValue is CustomLayout)
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
							CustomLayout customLayout = (CustomLayout)objectValue;
							foreach (Microsoft.Office.Interop.PowerPoint.Shape item in shapeRange)
							{
								b = A(customLayout, CS_0024_003C_003E8__locals7.A, item, d, activePresentation, now, g);
								list.Add(PowerPointAddIn1.Shapes.Helpers.A(customLayout, b));
							}
							customLayout.Shapes.Range(list.ToArray()).Select();
							customLayout = null;
							break;
						}
						objectValue = RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue, null, AH.A(28234), new object[0], null, null, null));
						if (!(objectValue is Microsoft.Office.Interop.PowerPoint.Presentation))
						{
							continue;
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
						break;
					}
					clsClipboard.ClearClipboard();
					Base.A(AH.A(139940));
				}
				else
				{
					Forms.ErrorMessage(AH.A(139977));
				}
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception a2 = ex;
				Base.A(a2);
				ProjectData.ClearProjectError();
			}
			application = null;
			activePresentation = null;
			shapeRange = null;
			CS_0024_003C_003E8__locals7.A = null;
			b = null;
			objectValue = null;
			list = null;
			return;
		}
	}

	private static Microsoft.Office.Interop.PowerPoint.Shape A(Slide A, Microsoft.Office.Interop.PowerPoint.Shape B, Microsoft.Office.Interop.PowerPoint.Shape C, Master D, Microsoft.Office.Interop.PowerPoint.Presentation E, DateTime F, Behavior G)
	{
		Microsoft.Office.Interop.PowerPoint.Shape shape = TryPaste(B, A.Shapes);
		PowerPointAddIn1.MasterShapes.AddRemove.A(shape, C, D, E, F, G);
		return shape;
	}

	private static Microsoft.Office.Interop.PowerPoint.Shape A(CustomLayout A, Microsoft.Office.Interop.PowerPoint.Shape B, Microsoft.Office.Interop.PowerPoint.Shape C, Master D, Microsoft.Office.Interop.PowerPoint.Presentation E, DateTime F, Behavior G)
	{
		Microsoft.Office.Interop.PowerPoint.Shape shape = TryPaste(B, A.Shapes);
		PowerPointAddIn1.MasterShapes.AddRemove.A(shape, C, D, E, F, G);
		return shape;
	}

	public static Microsoft.Office.Interop.PowerPoint.Shape TryPaste(Microsoft.Office.Interop.PowerPoint.Shape shpMaster, Microsoft.Office.Interop.PowerPoint.Shapes shps)
	{
		int num = 1;
		do
		{
			try
			{
				return shps.Paste()[1];
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				Thread.Sleep(100);
				if (num == 10)
				{
					throw;
				}
				ProjectData.ClearProjectError();
			}
			num = checked(num + 1);
		}
		while (num <= 10);
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
			return null;
		}
	}

	private static void A(Microsoft.Office.Interop.PowerPoint.Shape A, Microsoft.Office.Interop.PowerPoint.Shape B, Master C, Microsoft.Office.Interop.PowerPoint.Presentation D, DateTime E, Behavior F)
	{
		A.Tags.Add(Base.TAG_ID, C.Id);
		Placeholders.Populate(A, D, C.PlaceholderText, E, "");
		Microsoft.Office.Interop.PowerPoint.Shape shape = A;
		switch (F)
		{
		case Behavior.AboveTopLeft:
			shape.Top = B.Top - shape.Height;
			shape.Left = B.Left;
			break;
		case Behavior.AboveTopRight:
			shape.Top = B.Top - A.Height;
			shape.Left = B.Left + B.Width - shape.Width;
			break;
		case Behavior.BelowBottomRight:
			shape.Top = B.Top + B.Height;
			shape.Left = B.Left + B.Width - shape.Width;
			break;
		case Behavior.BelowBottomLeft:
			shape.Top = B.Top + B.Height;
			shape.Left = B.Left;
			break;
		case Behavior.InsideTopLeft:
			shape.Top = B.Top;
			shape.Left = B.Left;
			break;
		case Behavior.InsideTopRight:
			shape.Top = B.Top;
			shape.Left = B.Left + B.Width - shape.Width;
			break;
		case Behavior.InsideBottomRight:
			shape.Top = B.Top + B.Height - shape.Height;
			shape.Left = B.Left + B.Width - shape.Width;
			break;
		case Behavior.InsideBottomLeft:
			shape.Top = B.Top + B.Height - shape.Height;
			shape.Left = B.Left;
			break;
		case Behavior.CenterInShape:
			shape.Top = B.Top + B.Height / 2f - shape.Height / 2f;
			shape.Left = B.Left + B.Width / 2f - shape.Width / 2f;
			break;
		}
		shape.Visible = MsoTriState.msoTrue;
		shape = null;
	}

	public static void Toggle(IRibbonControl control, bool blnAdd)
	{
		if (!Licensing.AllowMasterShapesOperation())
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
			Toggle(control.Tag, blnAdd);
			return;
		}
	}

	public static void Toggle(string strId, bool blnAdd)
	{
		UF a = default(UF);
		UF CS_0024_003C_003E8__locals18 = new UF(a);
		CS_0024_003C_003E8__locals18.A = null;
		int num = 0;
		if (!Base.MyMasterShapes.TryGetValue(strId, out CS_0024_003C_003E8__locals18.A))
		{
			return;
		}
		checked
		{
			Application application;
			Microsoft.Office.Interop.PowerPoint.Presentation activePresentation;
			try
			{
				if (blnAdd)
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
					if (CS_0024_003C_003E8__locals18.A.Type == MsoShapeType.msoPlaceholder)
					{
						Forms.ErrorMessage(AH.A(140062));
						CS_0024_003C_003E8__locals18.A = null;
						return;
					}
				}
				Master master = MasterShapeProperties(CS_0024_003C_003E8__locals18.A);
				application = NG.A.Application;
				activePresentation = application.ActivePresentation;
				if (activePresentation.Slides.Count == 0)
				{
					while (true)
					{
						switch (2)
						{
						case 0:
							continue;
						}
						Forms.WarningMessage(AH.A(140390));
						break;
					}
				}
				else
				{
					application.StartNewUndoEntry();
					string text = ((!blnAdd || !master.PlaceholderText.Contains(Placeholders.PLACEHOLDER_STAMP)) ? "" : Stamps.GetPresentationStamp(activePresentation));
					DateTime now = DateTime.Now;
					if (clsClipboard.CopyWithWait((Action)([SpecialName] () =>
					{
						CS_0024_003C_003E8__locals18.A.Copy();
					}), 4000))
					{
						IEnumerator enumerator13 = default(IEnumerator);
						IEnumerator enumerator14 = default(IEnumerator);
						IEnumerator enumerator12 = default(IEnumerator);
						IEnumerator enumerator11 = default(IEnumerator);
						IEnumerator enumerator9 = default(IEnumerator);
						IEnumerator enumerator7 = default(IEnumerator);
						IEnumerator enumerator8 = default(IEnumerator);
						IEnumerator enumerator6 = default(IEnumerator);
						IEnumerator enumerator4 = default(IEnumerator);
						IEnumerator enumerator5 = default(IEnumerator);
						IEnumerator enumerator3 = default(IEnumerator);
						IEnumerator enumerator = default(IEnumerator);
						IEnumerator enumerator2 = default(IEnumerator);
						while (true)
						{
							switch (6)
							{
							case 0:
								continue;
							}
							switch (Base.A(master.Name))
							{
							case Behavior.SelectedSlides:
							{
								Base.A(application);
								List<Slide> list = PowerPointAddIn1.Slides.Helpers.A(application);
								if (list.Count > 0)
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
									if (Base.A(application, B: true))
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
										break;
									}
									using (List<Slide>.Enumerator enumerator16 = list.GetEnumerator())
									{
										while (enumerator16.MoveNext())
										{
											AddRemove(enumerator16.Current, CS_0024_003C_003E8__locals18.A, blnAdd, master, activePresentation, now, text);
										}
										while (true)
										{
											switch (1)
											{
											case 0:
												break;
											default:
												goto end_IL_01d5;
											}
											continue;
											end_IL_01d5:
											break;
										}
									}
									if (!blnAdd)
									{
										break;
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
									if (list.Count != 1)
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
										break;
									}
									Microsoft.Office.Interop.PowerPoint.Shape shape = list[0].Shapes[list[0].Shapes.Count];
									shape.Select();
									if (shape.HasTextFrame == MsoTriState.msoTrue)
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
										shape.TextFrame2.TextRange.get_Characters(-1, -1).Select();
									}
									shape = null;
								}
								else
								{
									Forms.WarningMessage(AH.A(140455));
								}
								break;
							}
							case Behavior.AllSlides:
								foreach (Slide item in activePresentation.Slides.Range(RuntimeHelpers.GetObjectValue(Missing.Value)))
								{
									AddRemove(item, CS_0024_003C_003E8__locals18.A, blnAdd, master, activePresentation, now, text);
									num++;
								}
								break;
							case Behavior.AllLayouts:
								try
								{
									enumerator13 = activePresentation.Designs.GetEnumerator();
									while (enumerator13.MoveNext())
									{
										Design design5 = (Design)enumerator13.Current;
										try
										{
											enumerator14 = design5.SlideMaster.CustomLayouts.GetEnumerator();
											while (enumerator14.MoveNext())
											{
												A((CustomLayout)enumerator14.Current, CS_0024_003C_003E8__locals18.A, blnAdd, master, activePresentation, now, text);
												num++;
											}
										}
										finally
										{
											if (enumerator14 is IDisposable)
											{
												while (true)
												{
													switch (6)
													{
													case 0:
														continue;
													}
													(enumerator14 as IDisposable).Dispose();
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
											goto end_IL_03b1;
										}
										continue;
										end_IL_03b1:
										break;
									}
								}
								finally
								{
									if (enumerator13 is IDisposable)
									{
										while (true)
										{
											switch (6)
											{
											case 0:
												continue;
											}
											(enumerator13 as IDisposable).Dispose();
											break;
										}
									}
								}
								break;
							case Behavior.ContentSlides:
								try
								{
									enumerator12 = activePresentation.Slides.Range(RuntimeHelpers.GetObjectValue(Missing.Value)).GetEnumerator();
									while (enumerator12.MoveNext())
									{
										Slide sld3 = (Slide)enumerator12.Current;
										if (PowerPointAddIn1.Slides.Helpers.IsSpecialSlide(sld3))
										{
											continue;
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
										AddRemove(sld3, CS_0024_003C_003E8__locals18.A, blnAdd, master, activePresentation, now, text);
										num++;
									}
								}
								finally
								{
									if (enumerator12 is IDisposable)
									{
										while (true)
										{
											switch (3)
											{
											case 0:
												continue;
											}
											(enumerator12 as IDisposable).Dispose();
											break;
										}
									}
								}
								break;
							case Behavior.ContentLayouts:
								foreach (Design design6 in activePresentation.Designs)
								{
									{
										enumerator11 = design6.SlideMaster.CustomLayouts.GetEnumerator();
										try
										{
											while (enumerator11.MoveNext())
											{
												CustomLayout customLayout4 = (CustomLayout)enumerator11.Current;
												if (!PowerPointAddIn1.Slides.Helpers.IsSpecialLayout(customLayout4))
												{
													A(customLayout4, CS_0024_003C_003E8__locals18.A, blnAdd, master, activePresentation, now, text);
													num++;
												}
											}
											while (true)
											{
												switch (4)
												{
												case 0:
													break;
												default:
													goto end_IL_04ec;
												}
												continue;
												end_IL_04ec:
												break;
											}
										}
										finally
										{
											IDisposable disposable = enumerator11 as IDisposable;
											if (disposable != null)
											{
												disposable.Dispose();
											}
										}
									}
								}
								break;
							case Behavior.SlidesShowingBackgroundGraphics:
								try
								{
									enumerator9 = activePresentation.Slides.Range(RuntimeHelpers.GetObjectValue(Missing.Value)).GetEnumerator();
									while (enumerator9.MoveNext())
									{
										Slide slide = (Slide)enumerator9.Current;
										if (slide.CustomLayout.DisplayMasterShapes != MsoTriState.msoTrue)
										{
											continue;
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
										AddRemove(slide, CS_0024_003C_003E8__locals18.A, blnAdd, master, activePresentation, now, text);
										num++;
									}
									while (true)
									{
										switch (3)
										{
										case 0:
											break;
										default:
											goto end_IL_05ac;
										}
										continue;
										end_IL_05ac:
										break;
									}
								}
								finally
								{
									if (enumerator9 is IDisposable)
									{
										while (true)
										{
											switch (3)
											{
											case 0:
												continue;
											}
											(enumerator9 as IDisposable).Dispose();
											break;
										}
									}
								}
								break;
							case Behavior.LayoutsShowingBackgroundGraphics:
								try
								{
									enumerator7 = activePresentation.Designs.GetEnumerator();
									while (enumerator7.MoveNext())
									{
										Design design3 = (Design)enumerator7.Current;
										try
										{
											enumerator8 = design3.SlideMaster.CustomLayouts.GetEnumerator();
											while (enumerator8.MoveNext())
											{
												CustomLayout customLayout3 = (CustomLayout)enumerator8.Current;
												if (customLayout3.DisplayMasterShapes != MsoTriState.msoTrue)
												{
													continue;
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
												A(customLayout3, CS_0024_003C_003E8__locals18.A, blnAdd, master, activePresentation, now, text);
												num++;
											}
											while (true)
											{
												switch (4)
												{
												case 0:
													break;
												default:
													goto end_IL_065e;
												}
												continue;
												end_IL_065e:
												break;
											}
										}
										finally
										{
											if (enumerator8 is IDisposable)
											{
												while (true)
												{
													switch (7)
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
									while (true)
									{
										switch (4)
										{
										case 0:
											break;
										default:
											goto end_IL_0696;
										}
										continue;
										end_IL_0696:
										break;
									}
								}
								finally
								{
									if (enumerator7 is IDisposable)
									{
										while (true)
										{
											switch (3)
											{
											case 0:
												continue;
											}
											(enumerator7 as IDisposable).Dispose();
											break;
										}
									}
								}
								break;
							case Behavior.DynamicSlides:
								try
								{
									enumerator6 = activePresentation.Slides.Range(RuntimeHelpers.GetObjectValue(Missing.Value)).GetEnumerator();
									while (enumerator6.MoveNext())
									{
										Slide sld2 = (Slide)enumerator6.Current;
										SlideType slideType = PowerPointAddIn1.Slides.Helpers.GetSlideType(sld2);
										unchecked
										{
											if ((uint)(slideType - 4) <= 1u)
											{
												continue;
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
											if ((uint)(slideType - 9) <= 1u)
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
												break;
											}
											AddRemove(sld2, CS_0024_003C_003E8__locals18.A, blnAdd, master, activePresentation, now, text);
										}
										num++;
									}
									while (true)
									{
										switch (6)
										{
										case 0:
											break;
										default:
											goto end_IL_074d;
										}
										continue;
										end_IL_074d:
										break;
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
								break;
							case Behavior.DynamicLayouts:
								try
								{
									enumerator4 = activePresentation.Designs.GetEnumerator();
									while (enumerator4.MoveNext())
									{
										Design design2 = (Design)enumerator4.Current;
										{
											enumerator5 = design2.SlideMaster.CustomLayouts.GetEnumerator();
											try
											{
												while (enumerator5.MoveNext())
												{
													CustomLayout customLayout2 = (CustomLayout)enumerator5.Current;
													SlideType layoutType = PowerPointAddIn1.Slides.Helpers.GetLayoutType(customLayout2);
													unchecked
													{
														if ((uint)(layoutType - 4) <= 1u)
														{
															continue;
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
														if ((uint)(layoutType - 9) <= 1u)
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
														A(customLayout2, CS_0024_003C_003E8__locals18.A, blnAdd, master, activePresentation, now, text);
													}
													num++;
												}
												while (true)
												{
													switch (3)
													{
													case 0:
														break;
													default:
														goto end_IL_081d;
													}
													continue;
													end_IL_081d:
													break;
												}
											}
											finally
											{
												IDisposable disposable2 = enumerator5 as IDisposable;
												if (disposable2 != null)
												{
													disposable2.Dispose();
												}
											}
										}
									}
									while (true)
									{
										switch (3)
										{
										case 0:
											break;
										default:
											goto end_IL_084d;
										}
										continue;
										end_IL_084d:
										break;
									}
								}
								finally
								{
									if (enumerator4 is IDisposable)
									{
										while (true)
										{
											switch (3)
											{
											case 0:
												continue;
											}
											(enumerator4 as IDisposable).Dispose();
											break;
										}
									}
								}
								break;
							case Behavior.SpecialSlides:
								try
								{
									enumerator3 = activePresentation.Slides.Range(RuntimeHelpers.GetObjectValue(Missing.Value)).GetEnumerator();
									while (enumerator3.MoveNext())
									{
										Slide sld = (Slide)enumerator3.Current;
										if (PowerPointAddIn1.Slides.Helpers.IsSpecialSlide(sld))
										{
											AddRemove(sld, CS_0024_003C_003E8__locals18.A, blnAdd, master, activePresentation, now, text);
											num++;
										}
									}
									while (true)
									{
										switch (7)
										{
										case 0:
											break;
										default:
											goto end_IL_08e1;
										}
										continue;
										end_IL_08e1:
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
								break;
							case Behavior.SpecialLayouts:
								try
								{
									enumerator = activePresentation.Designs.GetEnumerator();
									while (enumerator.MoveNext())
									{
										Design design = (Design)enumerator.Current;
										try
										{
											enumerator2 = design.SlideMaster.CustomLayouts.GetEnumerator();
											while (enumerator2.MoveNext())
											{
												CustomLayout customLayout = (CustomLayout)enumerator2.Current;
												if (!PowerPointAddIn1.Slides.Helpers.IsSpecialLayout(customLayout))
												{
													continue;
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
												A(customLayout, CS_0024_003C_003E8__locals18.A, blnAdd, master, activePresentation, now, text);
												num++;
											}
											while (true)
											{
												switch (1)
												{
												case 0:
													break;
												default:
													goto end_IL_0996;
												}
												continue;
												end_IL_0996:
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
									while (true)
									{
										switch (2)
										{
										case 0:
											break;
										default:
											goto end_IL_09d0;
										}
										continue;
										end_IL_09d0:
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
								break;
							}
							clsClipboard.ClearClipboard();
							string text2;
							if (!blnAdd)
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
								text2 = AH.A(140498);
							}
							else
							{
								text2 = AH.A(65425);
							}
							Base.A(text2 + AH.A(140511));
							break;
						}
					}
					else
					{
						Forms.ErrorMessage(AH.A(139977));
					}
				}
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception a2 = ex;
				Base.A(a2);
				ProjectData.ClearProjectError();
			}
			application = null;
			activePresentation = null;
			CS_0024_003C_003E8__locals18.A = null;
		}
	}

	public static Master MasterShapeProperties(Microsoft.Office.Interop.PowerPoint.Shape shp)
	{
		Master result = new Master
		{
			PlaceholderText = A(shp)
		};
		Microsoft.Office.Interop.PowerPoint.Shape shape = shp;
		result.Name = shape.Name;
		result.Id = shape.Id.ToString();
		result.Top = shape.Top;
		result.Left = shape.Left;
		shape = null;
		return result;
	}

	public static void AddRemove(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shpMaster, bool blnAdd, Master master, Microsoft.Office.Interop.PowerPoint.Presentation pres, DateTime dt, string strStamp)
	{
		A(sld.Shapes, master.Id);
		if (blnAdd)
		{
			Microsoft.Office.Interop.PowerPoint.Shape shape = TryPaste(shpMaster, sld.Shapes);
			PrepShape(shape, master);
			Placeholders.Populate(shape, pres, master.PlaceholderText, dt, strStamp, sld);
			shape.Visible = MsoTriState.msoTrue;
		}
	}

	private static void A(CustomLayout A, Microsoft.Office.Interop.PowerPoint.Shape B, bool C, Master D, Microsoft.Office.Interop.PowerPoint.Presentation E, DateTime F, string G)
	{
		PowerPointAddIn1.MasterShapes.AddRemove.A(A.Shapes, D.Id);
		if (C)
		{
			Microsoft.Office.Interop.PowerPoint.Shape shape = TryPaste(B, A.Shapes);
			PrepShape(shape, D);
			Placeholders.Populate(shape, E, D.PlaceholderText, F, G);
			shape.Visible = MsoTriState.msoTrue;
		}
	}

	private static void A(Microsoft.Office.Interop.PowerPoint.Shapes A, string B)
	{
		for (int i = A.Count; i >= 1; i = checked(i + -1))
		{
			if (Base.A(A[i], B))
			{
				A[i].Delete();
			}
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
			return;
		}
	}

	public static void PrepShape(Microsoft.Office.Interop.PowerPoint.Shape shp, Master master)
	{
		shp.Top = master.Top;
		shp.Left = master.Left;
		shp.Tags.Add(Base.TAG_ID, master.Id);
		_ = null;
	}

	private static string A(Microsoft.Office.Interop.PowerPoint.Shape A)
	{
		if (A.HasTextFrame == MsoTriState.msoTrue)
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
			if (A.TextFrame2.HasText == MsoTriState.msoTrue)
			{
				return A.TextFrame2.TextRange.Text;
			}
		}
		if (A.Type == MsoShapeType.msoGroup)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					return PowerPointAddIn1.MasterShapes.AddRemove.A(A.GroupItems[1]);
				}
			}
		}
		return "";
	}
}
