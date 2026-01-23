using System;
using System.Windows.Forms;
using A;
using MacabacusMacros;
using MacabacusMacros.UI;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.MasterShapes;
using PowerPointAddIn1.Slides;

namespace PowerPointAddIn1.Agenda;

public sealed class Sections
{
	internal static bool A()
	{
		Microsoft.Office.Interop.PowerPoint.Presentation activePresentation = NG.A.Application.ActivePresentation;
		bool result = false;
		checked
		{
			if (G())
			{
				result = true;
				wpfSectionAdd wpfSectionAdd2 = new wpfSectionAdd();
				wpfSectionAdd2.ShowDialog();
				if (wpfSectionAdd2.DialogResult.HasValue)
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
					if (wpfSectionAdd2.DialogResult.Value)
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
						int num = Helpers.GetSlideIndex() + 1;
						NG.A.Application.StartNewUndoEntry();
						if (num <= activePresentation.Slides.Count)
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
							activePresentation.SectionProperties.AddBeforeSlide(num, wpfSectionAdd2.txtTopic.Text);
						}
						else
						{
							Slide slide = activePresentation.Slides.Add(activePresentation.Slides.Count + 1, PpSlideLayout.ppLayoutBlank);
							activePresentation.SectionProperties.AddBeforeSlide(activePresentation.Slides.Count, wpfSectionAdd2.txtTopic.Text);
							slide.Delete();
							JG.A(slide);
						}
						D();
						activePresentation.Slides[num].Select();
						Flysheets.UseExistingAgendaSlideTitle(activePresentation, activePresentation.Slides[num]);
						A(AH.A(4924));
					}
				}
				wpfSectionAdd2 = null;
			}
			activePresentation = null;
			return result;
		}
	}

	internal static bool B()
	{
		Microsoft.Office.Interop.PowerPoint.Presentation activePresentation = NG.A.Application.ActivePresentation;
		bool result = false;
		SlideRange slideRange = null;
		if (G())
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
				slideRange = activePresentation.Windows[1].Selection.SlideRange;
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			if (slideRange != null)
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
				if (slideRange.Count > 0)
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
					result = true;
					int sectionIndex = slideRange[1].sectionIndex;
					string text = activePresentation.SectionProperties.Name(sectionIndex);
					frmSectionRename frmSectionRename = new frmSectionRename();
					System.Windows.Forms.TextBox txtName = frmSectionRename.txtName;
					txtName.Text = text;
					txtName.Select();
					if (Update.A(text))
					{
						txtName.SelectionStart = 1;
						txtName.SelectionLength = checked(text.Length - 1);
					}
					else
					{
						txtName.SelectAll();
					}
					txtName = null;
					if (frmSectionRename.ShowDialog() == DialogResult.OK)
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
						NG.A.Application.StartNewUndoEntry();
						activePresentation.SectionProperties.Rename(sectionIndex, frmSectionRename.txtName.Text);
						D();
						A(AH.A(4947));
					}
					frmSectionRename.Dispose();
					frmSectionRename = null;
					goto IL_016f;
				}
			}
			result = false;
		}
		goto IL_016f;
		IL_016f:
		activePresentation = null;
		slideRange = null;
		return result;
	}

	internal static bool C()
	{
		Microsoft.Office.Interop.PowerPoint.Application application = NG.A.Application;
		SlideRange slideRange = null;
		bool result = false;
		if (G())
		{
			try
			{
				slideRange = application.ActiveWindow.Selection.SlideRange;
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			if (slideRange != null && slideRange.Count > 0)
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
				SlideType slideType = Helpers.GetSlideType(slideRange[1]);
				if (slideType != SlideType.Flysheet)
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
					if (slideType != SlideType.Agenda)
					{
						goto IL_00f5;
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
				application.StartNewUndoEntry();
				application.ActivePresentation.SectionProperties.Delete(slideRange[1].sectionIndex, deleteSlides: false);
				result = true;
				slideRange[1].Delete();
				D();
				A(AH.A(4976));
			}
		}
		goto IL_00f5;
		IL_00f5:
		application = null;
		slideRange = null;
		return result;
	}

	internal static bool D()
	{
		Microsoft.Office.Interop.PowerPoint.Application application = NG.A.Application;
		SlideRange slideRange = null;
		bool result = false;
		if (G())
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
				slideRange = application.ActiveWindow.Selection.SlideRange;
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			if (slideRange != null)
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
				if (slideRange.Count > 0)
				{
					SlideType slideType = Helpers.GetSlideType(slideRange[1]);
					if (slideType != SlideType.Flysheet)
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
						if (slideType != SlideType.Agenda)
						{
							goto IL_00e3;
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
					application.StartNewUndoEntry();
					application.ActivePresentation.SectionProperties.Delete(slideRange[1].sectionIndex, deleteSlides: true);
					result = true;
					D();
					A(AH.A(5005));
				}
			}
		}
		goto IL_00e3;
		IL_00e3:
		application = null;
		slideRange = null;
		return result;
	}

	internal static void A()
	{
		Microsoft.Office.Interop.PowerPoint.Application application = NG.A.Application;
		Microsoft.Office.Interop.PowerPoint.Presentation activePresentation = application.ActivePresentation;
		if (G())
		{
			application.StartNewUndoEntry();
			Microsoft.Office.Interop.PowerPoint.Slides slides = activePresentation.Slides;
			for (int i = slides.Count; i >= 1; i = checked(i + -1))
			{
				SlideType slideType = Helpers.GetSlideType(slides[i]);
				if ((uint)(slideType - 2) > 1u)
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
					if (slideType != SlideType.Agenda)
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
				slides[i].Delete();
			}
			slides = null;
			SectionTitles.Remove(activePresentation);
			A(AH.A(5052));
		}
		application = null;
		activePresentation = null;
	}

	internal static bool E()
	{
		Microsoft.Office.Interop.PowerPoint.Application application = NG.A.Application;
		SlideRange slideRange = null;
		bool result = false;
		if (G())
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			try
			{
				slideRange = application.ActiveWindow.Selection.SlideRange;
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			if (slideRange != null)
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
				if (slideRange.Count > 0)
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
					int sectionIndex = slideRange[1].sectionIndex;
					application.StartNewUndoEntry();
					application.ActivePresentation.SectionProperties.Move(sectionIndex, checked(sectionIndex - 1));
					result = true;
					D();
					A(AH.A(5089));
				}
			}
		}
		application = null;
		slideRange = null;
		return result;
	}

	internal static bool F()
	{
		Microsoft.Office.Interop.PowerPoint.Application application = NG.A.Application;
		SlideRange slideRange = null;
		bool result = false;
		if (G())
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
				slideRange = application.ActiveWindow.Selection.SlideRange;
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			if (slideRange != null)
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
				if (slideRange.Count > 0)
				{
					int sectionIndex = slideRange[1].sectionIndex;
					application.StartNewUndoEntry();
					application.ActivePresentation.SectionProperties.Move(sectionIndex, checked(sectionIndex + 1));
					result = true;
					D();
					A(AH.A(5120));
				}
			}
		}
		application = null;
		slideRange = null;
		return result;
	}

	private static bool G()
	{
		return KG.A.OverrideSectionActions;
	}

	internal static void B()
	{
		if (!Licensing.AllowAgendaOperation())
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
			Microsoft.Office.Interop.PowerPoint.Application application = NG.A.Application;
			SlideRange slideRange = null;
			try
			{
				slideRange = application.ActiveWindow.Selection.SlideRange;
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			if (slideRange != null)
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
				if (slideRange.Count > 0)
				{
					int sectionIndex = slideRange[1].sectionIndex;
					SectionProperties sectionProperties = application.ActivePresentation.SectionProperties;
					if (Update.A(sectionProperties.Name(sectionIndex)))
					{
						application.StartNewUndoEntry();
						sectionProperties.Rename(sectionIndex, Strings.Mid(sectionProperties.Name(sectionIndex), 2));
						D();
					}
					sectionProperties = null;
					A(AH.A(5155));
					goto IL_00f2;
				}
			}
			Forms.ErrorMessage(AH.A(5186));
			goto IL_00f2;
			IL_00f2:
			application = null;
			slideRange = null;
			return;
		}
	}

	internal static void C()
	{
		if (!Licensing.AllowAgendaOperation())
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
			Microsoft.Office.Interop.PowerPoint.Application application = NG.A.Application;
			SlideRange slideRange = null;
			try
			{
				slideRange = application.ActiveWindow.Selection.SlideRange;
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			if (slideRange != null && slideRange.Count > 0)
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
				int sectionIndex = slideRange[1].sectionIndex;
				SectionProperties sectionProperties = application.ActivePresentation.SectionProperties;
				if (!Update.A(sectionProperties.Name(sectionIndex)))
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
					application.StartNewUndoEntry();
					sectionProperties.Rename(sectionIndex, Constants.SUBSECTION_PREFIX + sectionProperties.Name(sectionIndex));
					D();
				}
				sectionProperties = null;
				A(AH.A(5255));
			}
			else
			{
				Forms.ErrorMessage(AH.A(5186));
			}
			application = null;
			slideRange = null;
			return;
		}
	}

	private static void D()
	{
		Update.A(A: true);
	}

	private static void A(string A)
	{
		clsReporting.LogActivity((ActivityApp)2, (ActivityCategory)11, A);
	}
}
