using System;
using System.Collections;
using System.Collections.ObjectModel;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Media;
using A;
using MacabacusMacros.Explorer;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.Agenda;
using PowerPointAddIn1.Library2.Versioning;
using PowerPointAddIn1.Links;
using PowerPointAddIn1.Shapes;
using PowerPointAddIn1.Slides;
using PowerPointAddIn1.Template;

namespace PowerPointAddIn1.Explorer;

public sealed class SlideItem : BaseItem
{
	[CompilerGenerated]
	private PresentationItem m_A;

	[CompilerGenerated]
	private Slide m_A;

	[CompilerGenerated]
	private SlideType m_A;

	[CompilerGenerated]
	private int m_A;

	private Brush m_A;

	private Brush B;

	private bool m_A;

	private ObservableCollection<ContentItem> m_A;

	private string m_A;

	private bool B;

	private Visibility m_A;

	private bool C;

	private Visibility B;

	public PresentationItem Parent
	{
		[CompilerGenerated]
		get
		{
			return this.m_A;
		}
		[CompilerGenerated]
		set
		{
			this.m_A = value;
		}
	}

	public Slide Slide
	{
		[CompilerGenerated]
		get
		{
			return this.m_A;
		}
		[CompilerGenerated]
		set
		{
			this.m_A = value;
		}
	}

	public SlideType SlideType
	{
		[CompilerGenerated]
		get
		{
			return this.m_A;
		}
		[CompilerGenerated]
		set
		{
			this.m_A = value;
		}
	}

	public int ShapeCount
	{
		[CompilerGenerated]
		get
		{
			return this.m_A;
		}
		[CompilerGenerated]
		set
		{
			this.m_A = value;
		}
	}

	public Brush FontColor
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			((BaseItem)this).NotifyPropertyChanged(AH.A(113606));
		}
	}

	public Brush IconColor
	{
		get
		{
			return this.B;
		}
		set
		{
			this.B = value;
			((BaseItem)this).NotifyPropertyChanged(AH.A(113625));
		}
	}

	public override bool IsSelected
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			((BaseItem)this).NotifyPropertyChanged(AH.A(62846));
			RefreshLabel();
			A();
		}
	}

	public ObservableCollection<ContentItem> Children
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			((BaseItem)this).NotifyPropertyChanged(AH.A(115335));
		}
	}

	public string Tooltip
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			((BaseItem)this).NotifyPropertyChanged(AH.A(113591));
		}
	}

	public bool IsLinked
	{
		get
		{
			return this.B;
		}
		set
		{
			this.B = value;
			LinkAdornerVisibility = ((!value) ? Visibility.Collapsed : Visibility.Visible);
		}
	}

	public Visibility LinkAdornerVisibility
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			((BaseItem)this).NotifyPropertyChanged(AH.A(113644));
		}
	}

	public bool IsLibraryContent
	{
		get
		{
			return C;
		}
		set
		{
			C = value;
			int libraryAdornerVisibility;
			if (!value)
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
				libraryAdornerVisibility = 2;
			}
			else
			{
				libraryAdornerVisibility = 0;
			}
			LibraryAdornerVisibility = (Visibility)libraryAdornerVisibility;
		}
	}

	public Visibility LibraryAdornerVisibility
	{
		get
		{
			return B;
		}
		set
		{
			B = value;
			((BaseItem)this).NotifyPropertyChanged(AH.A(113687));
		}
	}

	public SlideItem(PresentationItem pi, Slide sld)
		: base("", Pane.CachedObjects.GeoSlide)
	{
		this.m_A = Visibility.Collapsed;
		B = Visibility.Collapsed;
		Parent = pi;
		Slide = sld;
		IsLinked = false;
		IsLibraryContent = Tagging.A(sld);
		Tooltip = AH.A(36272) + Conversions.ToString(sld.SlideIndex);
		UpdateColors(sld.SlideShowTransition.Hidden);
		((BaseItem)this).IsExpanded = false;
		SlideType = PowerPointAddIn1.Slides.Helpers.GetSlideType(sld);
		switch (SlideType)
		{
		case SlideType.Title:
			((BaseItem)this).Label = AH.A(115653);
			((BaseItem)this).Icon = Pane.CachedObjects.GeoTitle;
			break;
		case SlideType.TableOfContents:
			((BaseItem)this).Label = AH.A(115674);
			((BaseItem)this).Icon = Pane.CachedObjects.GeoToC;
			break;
		case SlideType.Flysheet:
		case SlideType.Agenda:
			if (SlideType == SlideType.Agenda)
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
				Slide slide = TableOfContents.Slide(pi.Presentation);
				if (sld == slide)
				{
					((BaseItem)this).Label = AH.A(115674);
					((BaseItem)this).Icon = Pane.CachedObjects.GeoToC;
				}
				else
				{
					((BaseItem)this).Label = A(pi, sld);
					((BaseItem)this).Icon = Pane.CachedObjects.GeoFlysheet;
				}
				slide = null;
			}
			else
			{
				((BaseItem)this).Label = A(pi, sld);
				((BaseItem)this).Icon = Pane.CachedObjects.GeoFlysheet;
			}
			break;
		case SlideType.Legal:
			((BaseItem)this).Label = AH.A(115709);
			((BaseItem)this).Icon = Pane.CachedObjects.GeoLegal;
			break;
		case SlideType.Contact:
			((BaseItem)this).Label = AH.A(115736);
			((BaseItem)this).Icon = Pane.CachedObjects.GeoContact;
			break;
		case SlideType.Blank:
			((BaseItem)this).Label = AH.A(115775);
			((BaseItem)this).Icon = Pane.CachedObjects.GeoBlank;
			break;
		case SlideType.CoverFront:
			((BaseItem)this).Label = AH.A(115798);
			((BaseItem)this).Icon = Pane.CachedObjects.GeoFrontCover;
			break;
		case SlideType.CoverBack:
			((BaseItem)this).Label = AH.A(115821);
			((BaseItem)this).Icon = Pane.CachedObjects.GeoBackCover;
			break;
		default:
			((BaseItem)this).Label = A();
			break;
		}
		((BaseItem)this).Icon.Freeze();
		Children = new ObservableCollection<ContentItem>();
		Children.Add(new DummyItem());
	}

	private string A(PresentationItem A, Slide B)
	{
		string result;
		try
		{
			result = AH.A(15089) + A.Presentation.SectionProperties.Name(B.sectionIndex) + AH.A(115352);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			result = AH.A(2597);
			ProjectData.ClearProjectError();
		}
		return result;
	}

	public void Search(string strQuery)
	{
		((BaseItem)this).IsHighlighted = ((BaseItem)this).Label.ToLower().Contains(strQuery);
		if (((BaseItem)this).IsHighlighted)
		{
			return;
		}
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
			if (SlideType == SlideType.Content || !strQuery.StartsWith(AH.A(4150)))
			{
				return;
			}
			if (Operators.CompareString(strQuery, AH.A(115373), TextCompare: false) == 0)
			{
				goto IL_0234;
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
			if (Operators.CompareString(strQuery, AH.A(115390), TextCompare: false) == 0)
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
				if (SlideType == SlideType.Title)
				{
					goto IL_0234;
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
			if (Operators.CompareString(strQuery, AH.A(115403), TextCompare: false) == 0)
			{
				if (SlideType == SlideType.CoverFront)
				{
					goto IL_0234;
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
			if (Operators.CompareString(strQuery, AH.A(115416), TextCompare: false) == 0)
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
				if (SlideType == SlideType.CoverBack)
				{
					goto IL_0234;
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
			if (Operators.CompareString(strQuery, AH.A(115427), TextCompare: false) == 0)
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
				if (SlideType == SlideType.Legal)
				{
					goto IL_0234;
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
			}
			if (Operators.CompareString(strQuery, AH.A(115440), TextCompare: false) == 0)
			{
				if (SlideType == SlideType.Contact)
				{
					goto IL_0234;
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
			}
			if (Operators.CompareString(strQuery, AH.A(115457), TextCompare: false) == 0)
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
				if (SlideType == SlideType.TableOfContents)
				{
					goto IL_0234;
				}
			}
			if (Operators.CompareString(strQuery, AH.A(115466), TextCompare: false) == 0)
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
				if (SlideType == SlideType.Agenda)
				{
					goto IL_0234;
				}
			}
			if (Operators.CompareString(strQuery, AH.A(115481), TextCompare: false) == 0)
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
				if (SlideType == SlideType.Flysheet)
				{
					goto IL_0234;
				}
			}
			int isHighlighted = ((Operators.CompareString(strQuery, AH.A(115500), TextCompare: false) == 0 && SlideType == SlideType.Blank) ? 1 : 0);
			goto IL_0235;
			IL_0235:
			((BaseItem)this).IsHighlighted = (byte)isHighlighted != 0;
			return;
			IL_0234:
			isHighlighted = 1;
			goto IL_0235;
		}
	}

	public void Refresh()
	{
		A();
		RefreshLabel();
		Children.Clear();
		if (((BaseItem)this).IsExpanded)
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
			Populate();
		}
		else
		{
			Children.Add(new DummyItem());
		}
		ShapeCount = Slide.Shapes.Count;
	}

	public void RefreshLabel()
	{
		if (SlideType == SlideType.Content)
		{
			((BaseItem)this).Label = A();
			UpdateColors(Slide.SlideShowTransition.Hidden);
		}
		Tooltip = AH.A(36272) + Conversions.ToString(Slide.SlideIndex);
	}

	private void A()
	{
		int num = Parent.Slides.IndexOf(this);
		if (num <= -1)
		{
			return;
		}
		int num2 = checked(Slide.SlideIndex - 1);
		if (num == num2)
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
			Parent.Slides.Move(num, num2);
			return;
		}
	}

	public void Populate()
	{
		MySettings settings = PB.Settings;
		bool explorerShowCharts = settings.ExplorerShowCharts;
		bool explorerShowTables = settings.ExplorerShowTables;
		bool explorerShowEmbeddedExcel = settings.ExplorerShowEmbeddedExcel;
		bool explorerShowEmbeddedWord = settings.ExplorerShowEmbeddedWord;
		bool explorerShowSmartArt = settings.ExplorerShowSmartArt;
		bool explorerShowImages = settings.ExplorerShowImages;
		bool explorerShowMedia = settings.ExplorerShowMedia;
		bool explorerShowInk = settings.ExplorerShowInk;
		_ = null;
		Children.Clear();
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = Slide.Shapes.GetEnumerator();
			while (enumerator.MoveNext())
			{
				Microsoft.Office.Interop.PowerPoint.Shape a = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator.Current;
				A(a, explorerShowCharts, explorerShowTables, explorerShowEmbeddedExcel, explorerShowEmbeddedWord, explorerShowSmartArt, explorerShowImages, explorerShowMedia, explorerShowInk);
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
		MySettings settings2 = PB.Settings;
		if (settings2.ExplorerShowComments)
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
			foreach (Comment comment in Slide.Comments)
			{
				Children.Add(new CommentItem(this, comment, Constants.ColorPalette.Amber.Clone()));
			}
		}
		if (settings2.ExplorerShowNotes)
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
			if (Slide.HasNotesPage == MsoTriState.msoTrue)
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
				{
					IEnumerator enumerator3 = Slide.NotesPage.Shapes.GetEnumerator();
					try
					{
						while (enumerator3.MoveNext())
						{
							Microsoft.Office.Interop.PowerPoint.Shape shape = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator3.Current;
							try
							{
								Microsoft.Office.Interop.PowerPoint.Shape shape2 = shape;
								if (shape2.Type == MsoShapeType.msoPlaceholder)
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
									if (shape2.PlaceholderFormat.Type == PpPlaceholderType.ppPlaceholderBody)
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
										if (shape2.TextFrame2.TextRange.Text.Length > 0)
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
											Children.Add(new NotesItem(this, shape, Constants.ColorPalette.Brown.Clone()));
										}
									}
								}
								shape2 = null;
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
							switch (4)
							{
							case 0:
								break;
							default:
								goto end_IL_0271;
							}
							continue;
							end_IL_0271:
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
		}
		if (settings2.ExplorerShowHyperlinks)
		{
			IEnumerator enumerator4 = default(IEnumerator);
			try
			{
				enumerator4 = Slide.Hyperlinks.GetEnumerator();
				while (enumerator4.MoveNext())
				{
					Hyperlink hyp = (Hyperlink)enumerator4.Current;
					try
					{
						Children.Add(new HyperlinkItem(this, hyp, Constants.ColorPalette.Purple.Clone()));
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
					switch (4)
					{
					case 0:
						break;
					default:
						goto end_IL_030b;
					}
					continue;
					end_IL_030b:
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
		}
		settings2 = null;
		ShapeCount = Slide.Shapes.Count;
	}

	private void A(Microsoft.Office.Interop.PowerPoint.Shape A, bool B, bool C, bool D, bool E, bool F, bool G, bool H, bool I)
	{
		Microsoft.Office.Interop.PowerPoint.Shape shape = A;
		if (shape.Type != MsoShapeType.msoGroup)
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
			if (B && shape.HasChart == MsoTriState.msoTrue)
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
				string strLabel = shape.Name;
				if (shape.Chart.HasTitle)
				{
					strLabel = shape.Chart.ChartTitle.Text;
				}
				Children.Add(new ChartItem(this, strLabel, A, Constants.ColorPalette.Green.Clone()));
			}
			else if (G && Images.HasPictureOrGraphic(A))
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
				Children.Add(NewImageItem(A));
			}
			else
			{
				if (C)
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
					if (shape.HasTable == MsoTriState.msoTrue)
					{
						Children.Add(new TableItem(this, A, Constants.ColorPalette.DeepPurple.Clone()));
						goto IL_0344;
					}
				}
				if (D && shape.Type == MsoShapeType.msoEmbeddedOLEObject)
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
					if (Operators.CompareString(Strings.Mid(shape.OLEFormat.ProgID, 1, 11), AH.A(115513), TextCompare: false) == 0)
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
						Children.Add(new EmbeddedExcelItem(this, A, Constants.ColorPalette.Teal.Clone()));
						goto IL_0344;
					}
				}
				if (E)
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
					if (shape.Type == MsoShapeType.msoEmbeddedOLEObject && Operators.CompareString(Strings.Mid(A.OLEFormat.ProgID, 1, 13), AH.A(115536), TextCompare: false) == 0)
					{
						Children.Add(new EmbeddedWordItem(this, A, Constants.ColorPalette.Indigo.Clone()));
						goto IL_0344;
					}
				}
				if (F)
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
					if (shape.HasSmartArt == MsoTriState.msoTrue)
					{
						Children.Add(new SmartArtItem(this, A, Constants.ColorPalette.Pink.Clone()));
						goto IL_0344;
					}
				}
				if (!H)
				{
					goto IL_027f;
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
				if (shape.Type != MsoShapeType.msoMedia)
				{
					if (shape.Type != MsoShapeType.msoWebVideo)
					{
						goto IL_027f;
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
				Children.Add(new MediaItem(this, A, Constants.ColorPalette.LightGreen.Clone()));
			}
		}
		else
		{
			IEnumerator enumerator = default(IEnumerator);
			try
			{
				enumerator = shape.GroupItems.GetEnumerator();
				while (enumerator.MoveNext())
				{
					Microsoft.Office.Interop.PowerPoint.Shape a = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator.Current;
					this.A(a, B, C, D, E, F, G, H, I);
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
						(enumerator as IDisposable).Dispose();
						break;
					}
				}
			}
		}
		goto IL_0344;
		IL_027f:
		if (I)
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
			if (shape.Type != MsoShapeType.msoInk)
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
				if (shape.Type != MsoShapeType.msoInkComment)
				{
					goto IL_0344;
				}
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
			Children.Add(new InkItem(this, A, Constants.ColorPalette.Red.Clone()));
		}
		goto IL_0344;
		IL_0344:
		shape = null;
	}

	public ImageItem NewImageItem(Microsoft.Office.Interop.PowerPoint.Shape shp)
	{
		return new ImageItem(this, shp, Constants.ColorPalette.Blue.Clone());
	}

	public void MoveUp()
	{
		int num = Parent.Slides.IndexOf(this);
		checked
		{
			if (num > 0)
			{
				Slide.MoveTo(Slide.SlideIndex - 1);
				Parent.Slides.Move(num, num - 1);
			}
		}
	}

	public void MoveDown()
	{
		int num = Parent.Slides.IndexOf(this);
		checked
		{
			if (num >= Parent.Slides.Count - 1)
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
				Slide.MoveTo(Slide.SlideIndex + 1);
				Parent.Slides.Move(num, num + 1);
				return;
			}
		}
	}

	public void Hide()
	{
		Slide.SlideShowTransition.Hidden = MsoTriState.msoTrue;
		UpdateColors(MsoTriState.msoTrue);
	}

	public void Unhide()
	{
		Slide.SlideShowTransition.Hidden = MsoTriState.msoFalse;
		UpdateColors(MsoTriState.msoFalse);
	}

	public void Delete()
	{
		if (System.Windows.Forms.MessageBox.Show(AH.A(115563), AH.A(5874), MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation) == DialogResult.OK)
		{
			Slide.Application.StartNewUndoEntry();
			Slide.Delete();
		}
	}

	public void Rename()
	{
		Miscellaneous.Rename();
		RefreshLabel();
	}

	public void Insert()
	{
		InsertSlide.ShowDialog();
	}

	public void SendToEnd()
	{
		checked
		{
			int oldIndex = Slide.SlideIndex - 1;
			Slide.Select();
			Miscellaneous.SendToEnd();
			Parent.Slides.Move(oldIndex, Slide.SlideIndex - 1);
		}
	}

	public void RemoveChild(ContentItem itm)
	{
		Children.Remove(itm);
	}

	public bool IsPopulated()
	{
		if (Children.Count == 1)
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
					return !(Children[0] is DummyItem);
				}
			}
		}
		return true;
	}

	private string A()
	{
		string text = Slide.Name;
		try
		{
			text = PowerPointAddIn1.Slides.Helpers.GetSlideTitle(Slide).Replace(AH.A(7894), AH.A(14625)).Replace(AH.A(47331), AH.A(14625))
				.Replace(AH.A(47334), AH.A(14625))
				.Replace(AH.A(7894), AH.A(14625))
				.Replace(AH.A(115650), AH.A(14625))
				.Trim();
			if (text.Length == 0)
			{
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
					text = Slide.Name;
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
		return text;
	}

	private bool A(Microsoft.Office.Interop.PowerPoint.Shape A)
	{
		if (PB.Settings.ExplorerShowLinkedShapes)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					return !PowerPointAddIn1.Links.Shapes.IsLinked(A);
				}
			}
		}
		return true;
	}

	public void UpdateColors(MsoTriState hidden)
	{
		double opacity = ((hidden == MsoTriState.msoFalse) ? 1.0 : base.HIDDEN_OPACITY);
		FontColor = new SolidColorBrush(base.DEFAULT_FONT_COLOR);
		IconColor = FontColor;
		FontColor.Opacity = opacity;
		IconColor.Opacity = opacity;
	}
}
