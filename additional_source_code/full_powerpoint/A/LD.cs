using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Windows;
using System.Windows.Media;
using System.Xml;
using MacabacusMacros;
using MacabacusMacros.Libraries;
using MacabacusMacros.Libraries.Manage.Publish;
using MacabacusMacros.Libraries.Versioning;
using MacabacusMacros.Links;
using MacabacusMacros.UI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.Library2.Versioning;
using PowerPointAddIn1.Presentation;

namespace A;

internal sealed class LD : INotifyPropertyChanged
{
	[CompilerGenerated]
	private PropertyChangedEventHandler m_A;

	private bool m_A;

	private Visibility m_A;

	private Visibility m_B;

	[CompilerGenerated]
	private int m_A;

	[CompilerGenerated]
	private List<ImageSource> m_A;

	private ImageSource m_A;

	[CompilerGenerated]
	private SlideItem m_A;

	[CompilerGenerated]
	private string m_A;

	[CompilerGenerated]
	private List<Slide> m_A;

	[CompilerGenerated]
	private int m_B;

	[CompilerGenerated]
	private string m_B;

	[CompilerGenerated]
	private Geometry m_A;

	[CompilerGenerated]
	private string m_C;

	[CompilerGenerated]
	private string m_D;

	public bool IsChecked
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			A(AH.A(12198));
		}
	}

	public Visibility LeftArrowVisibility
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			A(AH.A(63229));
		}
	}

	public Visibility RightArrowVisibility
	{
		get
		{
			return this.m_B;
		}
		set
		{
			this.m_B = value;
			A(AH.A(63268));
		}
	}

	public int PreviewIndex
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

	private List<ImageSource> Thumbnails
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

	public ImageSource Thumbnail
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			A(AH.A(63309));
		}
	}

	public SlideItem SlideItem
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

	private string PresentationPath
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

	public List<Slide> OldSlides
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

	public int SlideCount
	{
		[CompilerGenerated]
		get
		{
			return this.m_B;
		}
		[CompilerGenerated]
		set
		{
			this.m_B = value;
		}
	}

	public string Modified
	{
		[CompilerGenerated]
		get
		{
			return this.m_B;
		}
		[CompilerGenerated]
		set
		{
			this.m_B = value;
		}
	}

	public Geometry LibraryIcon
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

	public string LibraryName
	{
		[CompilerGenerated]
		get
		{
			return this.m_C;
		}
		[CompilerGenerated]
		set
		{
			this.m_C = value;
		}
	}

	public string GroupName
	{
		[CompilerGenerated]
		get
		{
			return this.m_D;
		}
		[CompilerGenerated]
		set
		{
			this.m_D = value;
		}
	}

	public event PropertyChangedEventHandler PropertyChanged
	{
		[CompilerGenerated]
		add
		{
			PropertyChangedEventHandler propertyChangedEventHandler = this.m_A;
			PropertyChangedEventHandler propertyChangedEventHandler2;
			do
			{
				propertyChangedEventHandler2 = propertyChangedEventHandler;
				PropertyChangedEventHandler value2 = (PropertyChangedEventHandler)Delegate.Combine(propertyChangedEventHandler2, value);
				propertyChangedEventHandler = Interlocked.CompareExchange(ref this.m_A, value2, propertyChangedEventHandler2);
			}
			while ((object)propertyChangedEventHandler != propertyChangedEventHandler2);
		}
		[CompilerGenerated]
		remove
		{
			PropertyChangedEventHandler propertyChangedEventHandler = this.m_A;
			PropertyChangedEventHandler propertyChangedEventHandler2;
			do
			{
				propertyChangedEventHandler2 = propertyChangedEventHandler;
				PropertyChangedEventHandler value2 = (PropertyChangedEventHandler)Delegate.Remove(propertyChangedEventHandler2, value);
				propertyChangedEventHandler = Interlocked.CompareExchange(ref this.m_A, value2, propertyChangedEventHandler2);
			}
			while ((object)propertyChangedEventHandler != propertyChangedEventHandler2);
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
	}

	public LD(SlideItem A)
	{
		//IL_0023: Unknown result type (might be due to invalid IL or missing references)
		//IL_0028: Unknown result type (might be due to invalid IL or missing references)
		//IL_00d7: Unknown result type (might be due to invalid IL or missing references)
		//IL_00dc: Unknown result type (might be due to invalid IL or missing references)
		SlideItem = A;
		IsChecked = true;
		OldSlides = A.Slides;
		string directoryName = Path.GetDirectoryName(((ContentItem)A).ContentInfo.ContentPath);
		string a;
		try
		{
			XmlDocument xmlDocument = new XmlDocument();
			xmlDocument.Load(Manifests.GetManifestPath(directoryName));
			if (xmlDocument.DocumentElement.Attributes[AH.A(63335)] != null)
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
				GroupName = xmlDocument.DocumentElement.Attributes[AH.A(63335)].Value;
			}
			else
			{
				GroupName = directoryName.Split(Path.DirectorySeparatorChar).Last();
			}
			XmlNode xmlNode = Content.ContentIdNode(xmlDocument.DocumentElement, ((ContentItem)A).ContentInfo.ContentId);
			a = Path.Combine(directoryName, xmlNode.Attributes[AH.A(63344)].Value);
			string text = Path.Combine(directoryName, xmlNode.Attributes[AH.A(63355)].Value);
			try
			{
				SlideCount = Conversions.ToInteger(xmlNode.Attributes[AH.A(63364)].Value);
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				SlideCount = 1;
				ProjectData.ClearProjectError();
			}
			xmlNode = null;
			xmlDocument = null;
			Modified = Updates.GetLastModifiedTime(text);
			PresentationPath = text;
		}
		catch (Exception ex3)
		{
			ProjectData.SetProjectError(ex3);
			Exception ex4 = ex3;
			a = "";
			string text = "";
			SlideCount = 0;
			clsReporting.LogException(ex4);
			ProjectData.ClearProjectError();
		}
		using (List<LibraryItem>.Enumerator enumerator = Base.LibraryCollection.GetEnumerator())
		{
			while (true)
			{
				if (enumerator.MoveNext())
				{
					LibraryItem current = enumerator.Current;
					if (!directoryName.StartsWith(current.Location))
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
						LibraryName = current.Name;
						LibraryIcon = current.Icon;
						break;
					}
					break;
				}
				while (true)
				{
					switch (6)
					{
					case 0:
						break;
					default:
						goto end_IL_0227;
					}
					continue;
					end_IL_0227:
					break;
				}
				break;
			}
		}
		LeftArrowVisibility = Visibility.Hidden;
		RightArrowVisibility = Visibility.Hidden;
		PreviewIndex = 0;
		Thumbnails = new List<ImageSource>();
		Thumbnails.Add(this.A(a));
		C();
	}

	private void A(string A)
	{
		PropertyChangedEventHandler propertyChangedEventHandler = this.m_A;
		if (propertyChangedEventHandler == null)
		{
			return;
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
			propertyChangedEventHandler(this, new PropertyChangedEventArgs(A));
			return;
		}
	}

	public void A()
	{
		checked
		{
			if (PreviewIndex >= SlideCount - 1)
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
				PreviewIndex++;
				if (Thumbnails.Count - 1 < PreviewIndex)
				{
					int sLIDE_THUMB_WIDTH = Core.SLIDE_THUMB_WIDTH;
					Presentation presentation = Helpers.OpenQuietly(NG.A.Application, PresentationPath);
					try
					{
						int scaleHeight = clsPowerPoint.ScaleHeight(sLIDE_THUMB_WIDTH, presentation);
						int count = presentation.Slides.Count;
						for (int i = 2; i <= count; i++)
						{
							Slide slide = presentation.Slides[i];
							string text = modFunctionsIO.PathGetTempFileName();
							slide.Export(text, AH.A(63328), sLIDE_THUMB_WIDTH, scaleHeight);
							Thumbnails.Add(A(text));
							File.Delete(text);
						}
						while (true)
						{
							switch (4)
							{
							case 0:
								break;
							default:
								goto end_IL_00e8;
							}
							continue;
							end_IL_00e8:
							break;
						}
					}
					catch (Exception ex)
					{
						ProjectData.SetProjectError(ex);
						Exception ex2 = ex;
						ProjectData.ClearProjectError();
					}
					presentation.Saved = MsoTriState.msoTrue;
					presentation.Close();
					presentation = null;
				}
				C();
				D();
				return;
			}
		}
	}

	public void B()
	{
		checked
		{
			if (PreviewIndex > 0)
			{
				PreviewIndex--;
				C();
				D();
			}
		}
	}

	private void C()
	{
		Thumbnail = Thumbnails[PreviewIndex];
	}

	public void D()
	{
		if (PreviewIndex == 0)
		{
			LeftArrowVisibility = Visibility.Hidden;
			RightArrowVisibility = Visibility.Visible;
			return;
		}
		if (PreviewIndex == checked(SlideCount - 1))
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					LeftArrowVisibility = Visibility.Visible;
					RightArrowVisibility = Visibility.Hidden;
					return;
				}
			}
		}
		LeftArrowVisibility = Visibility.Visible;
		RightArrowVisibility = Visibility.Visible;
	}

	private ImageSource A(string A)
	{
		ImageSource imageSource;
		try
		{
			Bitmap bitmap = (Bitmap)Image.FromFile(A);
			try
			{
				imageSource = Forms.GetImageSource(bitmap);
			}
			finally
			{
				if (bitmap != null)
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
						((IDisposable)bitmap).Dispose();
						break;
					}
				}
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			imageSource = Forms.GetImageSource(OB.InsertDrawingCanvas);
			ProjectData.ClearProjectError();
		}
		return imageSource;
	}

	public void A(bool A)
	{
	}
}
