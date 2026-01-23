using System;
using System.ComponentModel;
using System.IO;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using A;
using MacabacusMacros;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1.Shapes.Arrange;

public sealed class ShapeItem : INotifyPropertyChanged
{
	[CompilerGenerated]
	private PropertyChangedEventHandler m_A;

	[CompilerGenerated]
	private Shape m_A;

	[CompilerGenerated]
	private Size m_A;

	private string m_A;

	private ImageSource m_A;

	[CompilerGenerated]
	private string m_B;

	[CompilerGenerated]
	private int m_A;

	[CompilerGenerated]
	private bool m_A;

	[CompilerGenerated]
	private bool m_B;

	private bool m_C;

	public Shape Shape
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

	private Size OriginalSize
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

	public string SvgPath
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			NotifyPropertyChanged(AH.A(68947));
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
			NotifyPropertyChanged(AH.A(63309));
		}
	}

	private string PngPath
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

	public int ThumbHeight
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

	public bool UseSvg
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

	public bool UsePng
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

	public bool IsChecked
	{
		get
		{
			return this.m_C;
		}
		set
		{
			this.m_C = value;
			NotifyPropertyChanged(AH.A(12198));
		}
	}

	private event PropertyChangedEventHandler PropertyChanged
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
				return;
			}
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
	}

	internal ShapeItem(Shape A, bool B = true)
	{
		ThumbHeight = 66;
		Shape = A;
		OriginalSize = new Size(A.Width, A.Height);
		UseSvg = B;
		UsePng = !B;
		Thumbnail = null;
		IsChecked = false;
	}

	public void NotifyPropertyChanged(string propertyName)
	{
		PropertyChangedEventHandler propertyChangedEventHandler = this.m_A;
		if (propertyChangedEventHandler == null)
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
			propertyChangedEventHandler(this, new PropertyChangedEventArgs(propertyName));
			return;
		}
	}

	internal void A()
	{
		Shape.Width = (float)OriginalSize.Width;
		Shape.Height = (float)OriginalSize.Height;
	}

	internal void B()
	{
		string a = modFunctionsIO.PathGetTempFileName();
		if (UseSvg)
		{
			A(a);
		}
		else
		{
			B(a);
		}
	}

	private void A(string A)
	{
		if (SvgPath != null)
		{
			if (SvgPath.Length != 0)
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
				break;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
		}
		Shape.Export(A, (PpShapeFormat)6);
		SvgPath = A;
	}

	private void B(string A)
	{
		if (Thumbnail != null)
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
			Shape.Export(A, PpShapeFormat.ppShapeFormatPNG);
			Size imageSize = Images.GetImageSize(A);
			try
			{
				Thumbnail = this.A(imageSize, A);
				PngPath = A;
				return;
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

	internal void C()
	{
		try
		{
			if (UseSvg)
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
						File.Delete(SvgPath);
						return;
					}
				}
			}
			File.Delete(PngPath);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	private BitmapImage A(Size A, string B)
	{
		BitmapImage bitmapImage = new BitmapImage();
		bitmapImage.BeginInit();
		bitmapImage.CreateOptions = BitmapCreateOptions.IgnoreImageCache;
		bitmapImage.CacheOption = BitmapCacheOption.OnLoad;
		double a;
		if (A.Height < (double)ThumbHeight)
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
			if (A.Height > 0.0)
			{
				a = A.Height;
				goto IL_0074;
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
		a = ThumbHeight;
		goto IL_0074;
		IL_0074:
		bitmapImage.DecodePixelHeight = checked((int)Math.Round(a));
		bitmapImage.UriSource = new Uri(B);
		bitmapImage.EndInit();
		bitmapImage.Freeze();
		_ = null;
		return bitmapImage;
	}
}
