using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using A;
using MacabacusMacros;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1.Slides;

public sealed class Numbers
{
	public static bool IsSlideNumberPlaceholder(Microsoft.Office.Interop.PowerPoint.Shape shp)
	{
		bool result;
		try
		{
			int num;
			if (shp.Type == MsoShapeType.msoPlaceholder)
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
				num = ((shp.PlaceholderFormat.Type == PpPlaceholderType.ppPlaceholderSlideNumber) ? 1 : 0);
			}
			else
			{
				num = 0;
			}
			result = (byte)num != 0;
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

	public static void ShowDialog()
	{
		if (!clsRibbon.CallbackSlideView(ShowWarning: true))
		{
			return;
		}
		bool flag = false;
		try
		{
			IEnumerable<wpfSlideNums> source = System.Windows.Application.Current.Windows.OfType<wpfSlideNums>();
			if (source.Any())
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
				source.ElementAt(0).Activate();
				flag = true;
			}
			source = null;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		if (!flag)
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
			new wpfSlideNums().Show();
			_ = null;
		}
		clsReporting.LogActivity((ActivityApp)2, (ActivityCategory)11, AH.A(119607));
	}

	public static void Renumber()
	{
		int num = 0;
		Microsoft.Office.Interop.PowerPoint.Application application = NG.A.Application;
		if (application.Windows.Count == 0)
		{
			return;
		}
		Microsoft.Office.Interop.PowerPoint.Slides slides = application.ActivePresentation.Slides;
		application = null;
		checked
		{
			if (KG.A.SequentialSlideNumbers)
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
				IEnumerator enumerator = default(IEnumerator);
				try
				{
					enumerator = slides.GetEnumerator();
					while (enumerator.MoveNext())
					{
						Slide slide = (Slide)enumerator.Current;
						if (Helpers.GetSlideType(slide) == SlideType.Blank)
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
						foreach (Microsoft.Office.Interop.PowerPoint.Shape shape3 in slide.Shapes)
						{
							try
							{
								if (!IsSlideNumberPlaceholder(shape3))
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
									if (num == 0)
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
										num = (KG.A.SlideNumbersStartAtOne ? 1 : slide.SlideIndex);
									}
									else
									{
										num++;
									}
									shape3.TextFrame.TextRange.Text = num.ToString();
									break;
								}
								break;
							}
							catch (Exception ex)
							{
								ProjectData.SetProjectError(ex);
								Exception ex2 = ex;
								ProjectData.ClearProjectError();
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
							goto end_IL_0157;
						}
						continue;
						end_IL_0157:
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
			}
			else if (KG.A.SlideNumbersStartAtOne)
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
				{
					IEnumerator enumerator3 = slides.GetEnumerator();
					try
					{
						IEnumerator enumerator4 = default(IEnumerator);
						while (enumerator3.MoveNext())
						{
							Slide slide2 = (Slide)enumerator3.Current;
							if (Helpers.GetSlideType(slide2) == SlideType.Blank)
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
							if (num > 0)
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
								num++;
							}
							try
							{
								enumerator4 = slide2.Shapes.GetEnumerator();
								while (true)
								{
									if (enumerator4.MoveNext())
									{
										Microsoft.Office.Interop.PowerPoint.Shape shape2 = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator4.Current;
										try
										{
											if (!IsSlideNumberPlaceholder(shape2))
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
												if (num == 0)
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
													num++;
												}
												shape2.TextFrame.TextRange.Text = num.ToString();
												break;
											}
											break;
										}
										catch (Exception ex3)
										{
											ProjectData.SetProjectError(ex3);
											Exception ex4 = ex3;
											ProjectData.ClearProjectError();
										}
										continue;
									}
									while (true)
									{
										switch (6)
										{
										case 0:
											break;
										default:
											goto end_IL_026c;
										}
										continue;
										end_IL_026c:
										break;
									}
									break;
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
						while (true)
						{
							switch (7)
							{
							case 0:
								break;
							default:
								goto end_IL_02a6;
							}
							continue;
							end_IL_02a6:
							break;
						}
					}
					finally
					{
						IDisposable disposable = enumerator3 as IDisposable;
						if (disposable != null)
						{
							disposable.Dispose();
						}
					}
				}
			}
			slides = null;
		}
	}
}
