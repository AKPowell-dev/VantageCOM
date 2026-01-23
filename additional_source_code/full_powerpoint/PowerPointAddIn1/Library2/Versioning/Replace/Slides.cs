using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading;
using A;
using MacabacusMacros;
using MacabacusMacros.Libraries;
using MacabacusMacros.Libraries.Versioning;
using MacabacusMacros.UI;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.Library2.Insert;
using PowerPointAddIn1.Links;
using PowerPointAddIn1.Shapes;
using PowerPointAddIn1.Slides;

namespace PowerPointAddIn1.Library2.Versioning.Replace;

public sealed class Slides
{
	internal static void A(SlideItem A, Application B, ref List<string> C)
	{
		List<Slide> list = new List<Slide>();
		Dictionary<SlideItem, string> dictionary = new Dictionary<SlideItem, string>();
		wpfLinkRefresh wpfLinkRefresh = new wpfLinkRefresh();
		wpfLinkRefresh.Show();
		PowerPointAddIn1.Links.Common.UpdateProgressStart(wpfLinkRefresh, 0, 1);
		try
		{
			list = Slides.A(A, ref C);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			dictionary.Add(A, ex2.Message);
			clsReporting.LogException(ex2);
			ProjectData.ClearProjectError();
		}
		PowerPointAddIn1.Links.Common.UpdateProgressFinish(wpfLinkRefresh, 1, 1);
		Thread.Sleep(500);
		wpfLinkRefresh.Close();
		wpfLinkRefresh = null;
		Slides.A(list);
		A.Slides = list;
		if (dictionary.Any())
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
			Forms.ErrorMessage(dictionary.Values.ElementAtOrDefault(0));
		}
		B = null;
		list = null;
	}

	private static List<Slide> A(SlideItem A, ref List<string> B)
	{
		//IL_00d9: Unknown result type (might be due to invalid IL or missing references)
		//IL_00de: Unknown result type (might be due to invalid IL or missing references)
		//IL_0213: Unknown result type (might be due to invalid IL or missing references)
		//IL_0218: Unknown result type (might be due to invalid IL or missing references)
		//IL_021a: Unknown result type (might be due to invalid IL or missing references)
		//IL_021c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0223: Unknown result type (might be due to invalid IL or missing references)
		//IL_0234: Unknown result type (might be due to invalid IL or missing references)
		//IL_023b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0249: Unknown result type (might be due to invalid IL or missing references)
		Application application = NG.A.Application;
		List<Slide> list = new List<Slide>();
		PowerPointAddIn1.Library2.Insert.Slides.A(application);
		checked
		{
			try
			{
				Slide slide = A.Slides[0];
				Microsoft.Office.Interop.PowerPoint.Presentation presentation = (Microsoft.Office.Interop.PowerPoint.Presentation)slide.Parent;
				int slideIndex = slide.SlideIndex;
				_ = null;
				string text = A.Slides[0].Tags[Tagging.A];
				string text2 = Common.A((ContentItem)(object)A);
				Microsoft.Office.Interop.PowerPoint.Presentation presentation2 = PowerPointAddIn1.Library2.Insert.Common.A(text2, application, ref B);
				bool flag;
				bool flag2;
				int count;
				int num;
				try
				{
					Microsoft.Office.Interop.PowerPoint.Presentation presentation3 = presentation;
					PageSetup pageSetup = presentation3.PageSetup;
					flag = presentation2.PageSetup.SlideWidth != pageSetup.SlideWidth;
					flag2 = presentation2.PageSetup.SlideHeight != pageSetup.SlideHeight;
					pageSetup = null;
					count = presentation3.Designs.Count;
					_ = null;
					if (((ContentItem)A).ContentInfo.KeepSourceFormatting)
					{
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
							num = PowerPointAddIn1.Library2.Slides.A(presentation2, application);
							break;
						}
					}
					else
					{
						num = presentation.Slides.InsertFromFile(text2, slideIndex);
					}
				}
				finally
				{
					presentation2.Close();
					JG.A(presentation2);
				}
				for (int i = A.Slides.Count - 1; i >= 0; i += -1)
				{
					try
					{
						A.Slides[i].Delete();
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
					switch (1)
					{
					case 0:
						continue;
					}
					presentation.Application.ActiveWindow.View.GotoSlide(slideIndex);
					_ = new int[num - 1 + 1];
					int num2 = num - 1;
					Slide slide2;
					for (int j = 0; j <= num2; list.Add(slide2), slide2 = null, j++)
					{
						slide2 = presentation.Slides[slideIndex + j];
						if (Operators.CompareString(text, string.Empty, TextCompare: false) != 0)
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
							slide2.Tags.Add(Tagging.A, text);
						}
						if (((ContentItem)A).IsLegacySlideLink)
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
							ContentInfo contentInfo = ((ContentItem)A).ContentInfo;
							string b = Base.ConvertCachedToRemotePath(contentInfo.ContentPath, contentInfo.LibraryId);
							PowerPointAddIn1.Links.Slides.A(slide2, b, contentInfo.ContentId, contentInfo.KeepSourceFormatting);
							contentInfo = default(ContentInfo);
						}
						if (!flag)
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
							if (!flag2)
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
						}
						try
						{
							PowerPointAddIn1.Shapes.Images.FixDistortion(slide2, flag, flag2);
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
						switch (6)
						{
						case 0:
							continue;
						}
						presentation = null;
						if (!flag)
						{
							if (!flag2)
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
								break;
							}
						}
						try
						{
							Designs designs = presentation.Designs;
							if (designs.Count > count)
							{
								PowerPointAddIn1.Shapes.Images.FixDistortion(designs[designs.Count], flag, flag2);
							}
							designs = null;
						}
						catch (Exception ex5)
						{
							ProjectData.SetProjectError(ex5);
							Exception ex6 = ex5;
							ProjectData.ClearProjectError();
						}
						break;
					}
					break;
				}
			}
			finally
			{
				PowerPointAddIn1.Library2.Insert.Slides.B(application);
				application = null;
				Microsoft.Office.Interop.PowerPoint.Presentation presentation = null;
			}
			return list;
		}
	}

	private static void A(List<Slide> A)
	{
		try
		{
			PowerPointAddIn1.Slides.Helpers.SelectMultipleSlides((Microsoft.Office.Interop.PowerPoint.Presentation)A[0].Parent, A.Select([SpecialName] (Slide slide) => slide.SlideIndex));
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}
}
