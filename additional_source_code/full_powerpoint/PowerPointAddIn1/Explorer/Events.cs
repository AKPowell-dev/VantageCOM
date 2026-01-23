using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using A;
using MacabacusMacros.Explorer;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Tools;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.Shapes;

namespace PowerPointAddIn1.Explorer;

public sealed class Events
{
	private bool m_A;

	public static bool explorerDoNotClose = false;

	private static bool m_B = false;

	private static string m_A = "";

	public static bool blnBeforeClose = false;

	private static bool C = false;

	public Events()
	{
		this.m_A = false;
	}

	private static void A(Microsoft.Office.Interop.PowerPoint.Application A)
	{
		Microsoft.Office.Interop.PowerPoint.Application target = A;
		new ComAwareEventInfo(typeof(EApplication_Event), AH.A(105591)).AddEventHandler(target, new EApplication_WindowActivateEventHandler(Events.A));
		new ComAwareEventInfo(typeof(EApplication_Event), AH.A(12762)).AddEventHandler(target, new EApplication_WindowSelectionChangeEventHandler(WindowSelectionChange));
		new ComAwareEventInfo(typeof(EApplication_Event), AH.A(58943)).AddEventHandler(target, new EApplication_PresentationNewSlideEventHandler(PresentationNewSlide));
		new ComAwareEventInfo(typeof(EApplication_Event), AH.A(113050)).AddEventHandler(target, new EApplication_PresentationSaveEventHandler(Events.A));
		new ComAwareEventInfo(typeof(EApplication_Event), AH.A(113083)).AddEventHandler(target, new EApplication_AfterShapeSizeChangeEventHandler(Events.A));
		new ComAwareEventInfo(typeof(EApplication_Event), AH.A(56688)).AddEventHandler(target, new EApplication_PresentationBeforeCloseEventHandler(Events.A));
		new ComAwareEventInfo(typeof(EApplication_Event), AH.A(10098)).AddEventHandler(target, new EApplication_PresentationCloseFinalEventHandler(B));
		target = null;
	}

	public static void Disable(Microsoft.Office.Interop.PowerPoint.Application ppApp)
	{
		Microsoft.Office.Interop.PowerPoint.Application target = ppApp;
		new ComAwareEventInfo(typeof(EApplication_Event), AH.A(105591)).RemoveEventHandler(target, new EApplication_WindowActivateEventHandler(A));
		new ComAwareEventInfo(typeof(EApplication_Event), AH.A(12762)).RemoveEventHandler(target, new EApplication_WindowSelectionChangeEventHandler(WindowSelectionChange));
		new ComAwareEventInfo(typeof(EApplication_Event), AH.A(58943)).RemoveEventHandler(target, new EApplication_PresentationNewSlideEventHandler(PresentationNewSlide));
		new ComAwareEventInfo(typeof(EApplication_Event), AH.A(113050)).RemoveEventHandler(target, new EApplication_PresentationSaveEventHandler(A));
		new ComAwareEventInfo(typeof(EApplication_Event), AH.A(113083)).RemoveEventHandler(target, new EApplication_AfterShapeSizeChangeEventHandler(A));
		new ComAwareEventInfo(typeof(EApplication_Event), AH.A(56688)).RemoveEventHandler(target, new EApplication_PresentationBeforeCloseEventHandler(A));
		new ComAwareEventInfo(typeof(EApplication_Event), AH.A(10098)).RemoveEventHandler(target, new EApplication_PresentationCloseFinalEventHandler(B));
		target = null;
	}

	public static void Reset(Microsoft.Office.Interop.PowerPoint.Application ppApp)
	{
		Disable(ppApp);
		A(ppApp);
	}

	private static void A(Microsoft.Office.Interop.PowerPoint.Presentation A, DocumentWindow B)
	{
		if (!PB.Settings.ExplorerShowAllPresentations)
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
			Pane.A(A);
			return;
		}
	}

	public static void WindowSelectionChange(Selection Sel)
	{
		if (Sel.Type == PpSelectionType.ppSelectionText)
		{
			return;
		}
		explorerDoNotClose = false;
		blnBeforeClose = false;
		new ComAwareEventInfo(typeof(EApplication_Event), AH.A(58943)).RemoveEventHandler(Sel.Application, new EApplication_PresentationNewSlideEventHandler(PresentationNewSlide));
		new ComAwareEventInfo(typeof(EApplication_Event), AH.A(58943)).AddEventHandler(Sel.Application, new EApplication_PresentationNewSlideEventHandler(PresentationNewSlide));
		Slide slide = null;
		Microsoft.Office.Interop.PowerPoint.Shape shape = null;
		Microsoft.Office.Interop.PowerPoint.Presentation a;
		if (Sel.Type != PpSelectionType.ppSelectionNone)
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
			slide = Sel.SlideRange[1];
			a = (Microsoft.Office.Interop.PowerPoint.Presentation)slide.Parent;
			if (Sel.Type != PpSelectionType.ppSelectionShapes)
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
				if (Sel.Type != PpSelectionType.ppSelectionText)
				{
					goto IL_012b;
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
			try
			{
				shape = Base.SelectedShapes(Sel)[1];
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
		}
		else
		{
			a = Sel.Application.ActivePresentation;
		}
		goto IL_012b;
		IL_012b:
		PresentationItem presentationItem;
		try
		{
			if (Sel.Type != PpSelectionType.ppSelectionNone)
			{
				bool flag = false;
				presentationItem = A(a);
				if (presentationItem != null)
				{
					IEnumerator<SlideItem> enumerator = default(IEnumerator<SlideItem>);
					IEnumerator<ContentItem> enumerator2 = default(IEnumerator<ContentItem>);
					while (true)
					{
						switch (2)
						{
						case 0:
							continue;
						}
						if (!((BaseItem)presentationItem).IsExpanded)
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
							try
							{
								enumerator = presentationItem.Slides.GetEnumerator();
								while (enumerator.MoveNext())
								{
									SlideItem current = enumerator.Current;
									if (current.Slide != slide)
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
										if (shape != null && ((BaseItem)current).IsExpanded)
										{
											while (true)
											{
												switch (3)
												{
												case 0:
													continue;
												}
												try
												{
													enumerator2 = current.Children.GetEnumerator();
													while (true)
													{
														if (enumerator2.MoveNext())
														{
															ContentItem current2 = enumerator2.Current;
															if (current2.Shape == shape)
															{
																((BaseItem)current2).IsSelected = true;
																current2.Refresh();
																flag = true;
																break;
															}
															continue;
														}
														while (true)
														{
															switch (7)
															{
															case 0:
																break;
															default:
																goto end_IL_0207;
															}
															continue;
															end_IL_0207:
															break;
														}
														break;
													}
												}
												finally
												{
													if (enumerator2 != null)
													{
														while (true)
														{
															switch (6)
															{
															case 0:
																continue;
															}
															enumerator2.Dispose();
															break;
														}
													}
												}
												if (flag)
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
													if (current.ShapeCount != current.Slide.Shapes.Count)
													{
														current.Refresh();
													}
													else
													{
														current.IsSelected = true;
													}
													break;
												}
												break;
											}
										}
										else
										{
											current.IsSelected = true;
										}
										break;
									}
									break;
								}
							}
							finally
							{
								if (enumerator != null)
								{
									while (true)
									{
										switch (1)
										{
										case 0:
											continue;
										}
										enumerator.Dispose();
										break;
									}
								}
							}
							break;
						}
						break;
					}
				}
			}
			else
			{
				A((Microsoft.Office.Interop.PowerPoint.Presentation)slide.Parent).IsSelected = true;
			}
		}
		catch (Exception ex3)
		{
			ProjectData.SetProjectError(ex3);
			Exception ex4 = ex3;
			ProjectData.ClearProjectError();
		}
		a = null;
		slide = null;
		shape = null;
		presentationItem = null;
	}

	public static void PresentationNewSlide(Slide Sld)
	{
		Microsoft.Office.Interop.PowerPoint.Presentation presentation = (Microsoft.Office.Interop.PowerPoint.Presentation)Sld.Parent;
		if (presentation.Windows.Count == 0)
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
					presentation = null;
					return;
				}
			}
		}
		Events.m_B = true;
		C = true;
		checked
		{
			try
			{
				PresentationItem presentationItem = A(presentation);
				if (presentationItem != null)
				{
					if (presentationItem.Slides.Count < presentation.Slides.Count)
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
						SlideItem slideItem = new SlideItem(presentationItem, Sld);
						ObservableCollection<SlideItem> slides = presentationItem.Slides;
						slides.Add(slideItem);
						slides.Move(slides.Count - 1, Sld.SlideIndex - 1);
						_ = null;
						slideItem.IsSelected = true;
						slideItem = null;
					}
					presentationItem = null;
				}
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			presentation = null;
		}
	}

	private static void A(Microsoft.Office.Interop.PowerPoint.Presentation A)
	{
		try
		{
			Events.A(A).RefreshLabel();
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	private static void A(Microsoft.Office.Interop.PowerPoint.Presentation A, ref bool B)
	{
		blnBeforeClose = true;
		explorerDoNotClose = false;
		if (A.Saved == MsoTriState.msoTrue)
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
					explorerDoNotClose = true;
					return;
				}
			}
		}
		if (A.Path.Length <= 0)
		{
			return;
		}
		DialogResult dialogResult = MessageBox.Show(AH.A(113124) + A.Name + AH.A(113185), AH.A(113190), MessageBoxButtons.YesNoCancel, MessageBoxIcon.Exclamation);
		if (dialogResult != DialogResult.Cancel)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					if (dialogResult != DialogResult.Yes)
					{
						while (true)
						{
							switch (5)
							{
							case 0:
								break;
							default:
								if (dialogResult == DialogResult.No)
								{
									explorerDoNotClose = true;
									A.Saved = MsoTriState.msoTrue;
								}
								return;
							}
						}
					}
					explorerDoNotClose = true;
					A.Save();
					return;
				}
			}
		}
		B = true;
		blnBeforeClose = false;
	}

	private static void B(Microsoft.Office.Interop.PowerPoint.Presentation A)
	{
		blnBeforeClose = false;
		try
		{
			PresentationItem presentationItem = Events.A(A);
			if (presentationItem != null)
			{
				Pane.AllPresentations.Remove(presentationItem);
				presentationItem = null;
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		Dictionary<int, CustomTaskPane> TaskPanes = Pane.PanesCollection;
		clsPanes.RemoveOrphanedPanes(ref TaskPanes, Pane.PANE_TITLE);
		Pane.PanesCollection = TaskPanes;
		if (IG.A(NG.A.Application.Presentations) != 1)
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
			clsPanes.RemoveTaskPanesByTitle(Pane.PANE_TITLE);
			return;
		}
	}

	private static void A(Microsoft.Office.Interop.PowerPoint.Shape A)
	{
	}

	public static void RefreshPresentation(Microsoft.Office.Interop.PowerPoint.Presentation pres)
	{
		if (!Pane.IsOpen)
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
				A(pres).RefreshSlides();
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			Reset(pres.Application);
			return;
		}
	}

	private static PresentationItem A(Microsoft.Office.Interop.PowerPoint.Presentation A)
	{
		try
		{
			if (!PB.Settings.ExplorerShowAllPresentations)
			{
				foreach (KeyValuePair<int, CustomTaskPane> item in Pane.PanesCollection)
				{
					wpfExplorer wpfExplorer2 = Pane.A((ctpExplorer2)item.Value.Control);
					if (wpfExplorer2.ThisPresentationItem.Presentation == A)
					{
						while (true)
						{
							switch (5)
							{
							case 0:
								break;
							default:
								return wpfExplorer2.ThisPresentationItem;
							}
						}
					}
					wpfExplorer2 = null;
				}
			}
			else
			{
				IEnumerator<PresentationItem> enumerator2 = default(IEnumerator<PresentationItem>);
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
						enumerator2 = Pane.AllPresentations.GetEnumerator();
						while (enumerator2.MoveNext())
						{
							PresentationItem current = enumerator2.Current;
							if (current.Presentation != A)
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
								return current;
							}
						}
						while (true)
						{
							switch (7)
							{
							case 0:
								break;
							default:
								goto end_IL_0062;
							}
							continue;
							end_IL_0062:
							break;
						}
					}
					finally
					{
						if (enumerator2 != null)
						{
							while (true)
							{
								switch (5)
								{
								case 0:
									continue;
								}
								enumerator2.Dispose();
								break;
							}
						}
					}
					break;
				}
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		return null;
	}
}
