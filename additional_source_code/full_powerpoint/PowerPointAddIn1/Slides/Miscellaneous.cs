using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using A;
using MacabacusMacros;
using MacabacusMacros.Auth;
using MacabacusMacros.UI;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1.Slides;

public sealed class Miscellaneous
{
	[Serializable]
	[CompilerGenerated]
	internal sealed class _Closure_0024__
	{
		public static readonly _Closure_0024__ A;

		public static Func<Slide, int> A;

		static _Closure_0024__()
		{
			_Closure_0024__.A = new _Closure_0024__();
		}

		[SpecialName]
		internal int A(Slide A)
		{
			return A.SlideIndex;
		}
	}

	public static void Rename()
	{
		Application application = NG.A.Application;
		Selection selection;
		try
		{
			selection = application.ActiveWindow.Selection;
			if (selection.SlideRange.Count != 1)
			{
				throw new Exception();
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
				Slide slide = selection.SlideRange[1];
				string text = Forms.InputBox(AH.A(118862), AH.A(118887), slide.Name);
				try
				{
					if (Operators.CompareString(text, string.Empty, TextCompare: false) != 0 && text.Length > 0)
					{
						while (true)
						{
							switch (5)
							{
							case 0:
								continue;
							}
							application.StartNewUndoEntry();
							slide.Name = text;
							clsReporting.LogActivity((ActivityApp)2, (ActivityCategory)14, AH.A(118862));
							break;
						}
					}
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					Forms.ErrorMessage(ex2.Message);
					ProjectData.ClearProjectError();
				}
				slide = null;
				break;
			}
		}
		catch (Exception ex3)
		{
			ProjectData.SetProjectError(ex3);
			Exception ex4 = ex3;
			Forms.WarningMessage(AH.A(118952));
			ProjectData.ClearProjectError();
		}
		selection = null;
		application = null;
	}

	public static void SendToEnd()
	{
		Application application = NG.A.Application;
		if (application.Presentations.Count > 0)
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
			Selection selection;
			DocumentWindow activeWindow;
			try
			{
				activeWindow = application.ActiveWindow;
				if (activeWindow.ViewType == PpViewType.ppViewSlideSorter)
				{
					goto IL_008b;
				}
				if (activeWindow.Panes.Count > 1)
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
					if (activeWindow.Panes[2].ViewType == PpViewType.ppViewSlide)
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
						goto IL_008b;
					}
				}
				Forms.WarningMessage(AH.A(119177));
				goto end_IL_0033;
				IL_008b:
				selection = activeWindow.Selection;
				if (selection.Type == PpSelectionType.ppSelectionSlides)
				{
					while (true)
					{
						switch (5)
						{
						case 0:
							continue;
						}
						new List<Slide>();
						List<Slide> source = selection.SlideRange.Cast<Slide>().ToList();
						Func<Slide, int> keySelector;
						if (_Closure_0024__.A == null)
						{
							keySelector = (_Closure_0024__.A = [SpecialName] (Slide A) => A.SlideIndex);
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
							keySelector = _Closure_0024__.A;
						}
						int slideIndex = source.OrderBy(keySelector).ToList()[0].SlideIndex;
						application.StartNewUndoEntry();
						selection.SlideRange.MoveTo(application.ActivePresentation.Slides.Count);
						activeWindow.View.GotoSlide(slideIndex);
						clsReporting.LogActivity((ActivityApp)2, (ActivityCategory)14, AH.A(119049));
						break;
					}
				}
				else
				{
					Forms.WarningMessage(AH.A(119072));
				}
				end_IL_0033:;
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				Forms.ErrorMessage(ex2.Message);
				clsReporting.LogException(ex2);
				ProjectData.ClearProjectError();
			}
			selection = null;
			activeWindow = null;
		}
		application = null;
	}

	public static void RemoveUnusedLayouts()
	{
		if (!Access.AllowPowerPointOperation((PlanType)4, (Restriction)1, false))
		{
			return;
		}
		checked
		{
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
				int num = 0;
				Microsoft.Office.Interop.PowerPoint.Presentation activePresentation;
				try
				{
					Application application = NG.A.Application;
					activePresentation = application.ActivePresentation;
					application.StartNewUndoEntry();
					_ = null;
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					ProjectData.ClearProjectError();
					return;
				}
				int num2 = activePresentation.Designs.Count;
				while (num2 >= 1)
				{
					Master slideMaster = activePresentation.Designs[num2].SlideMaster;
					for (int i = slideMaster.CustomLayouts.Count; i >= 1; i += -1)
					{
						try
						{
							slideMaster.CustomLayouts[i].Delete();
							num++;
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
						switch (3)
						{
						case 0:
							continue;
						}
						if (slideMaster.CustomLayouts.Count == 0)
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
							try
							{
								activePresentation.Designs[num2].Delete();
							}
							catch (Exception ex5)
							{
								ProjectData.SetProjectError(ex5);
								Exception ex6 = ex5;
								ProjectData.ClearProjectError();
							}
						}
						slideMaster = null;
						num2 += -1;
						break;
					}
				}
				if (num != 0)
				{
					if (num != 1)
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
						Forms.SuccessMessage(AH.A(94493) + num + AH.A(119485));
					}
					else
					{
						Forms.SuccessMessage(AH.A(119390));
					}
				}
				else
				{
					Forms.InfoMessage(AH.A(119309));
				}
				activePresentation = null;
				clsReporting.LogActivity((ActivityApp)2, (ActivityCategory)14, AH.A(119564));
				return;
			}
		}
	}
}
