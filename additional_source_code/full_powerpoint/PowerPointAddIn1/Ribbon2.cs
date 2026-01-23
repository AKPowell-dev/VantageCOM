using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.CompilerServices;
using A;
using Microsoft.Office.Core;
using Microsoft.Office.Tools.Ribbon;

namespace PowerPointAddIn1;

public sealed class Ribbon2 : RibbonBase
{
	private IContainer m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("TabMacabacus")]
	private RibbonTab m_A;

	[AccessedThroughProperty("GroupSlides")]
	[CompilerGenerated]
	private RibbonGroup m_A;

	[AccessedThroughProperty("Button1")]
	[CompilerGenerated]
	private RibbonButton m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("MenuInsertSlide")]
	private RibbonMenu m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("ImportSlidePaste")]
	private RibbonButton B;

	[AccessedThroughProperty("InsertCover")]
	[CompilerGenerated]
	private RibbonButton C;

	[AccessedThroughProperty("InsertTableOfContents")]
	[CompilerGenerated]
	private RibbonButton D;

	[AccessedThroughProperty("InsertFlysheet")]
	[CompilerGenerated]
	private RibbonButton E;

	[AccessedThroughProperty("InsertLegal")]
	[CompilerGenerated]
	private RibbonButton F;

	[CompilerGenerated]
	[AccessedThroughProperty("InsertContact")]
	private RibbonButton G;

	[AccessedThroughProperty("MenuSlideTools")]
	[CompilerGenerated]
	private RibbonMenu B;

	[CompilerGenerated]
	[AccessedThroughProperty("Separator1")]
	private RibbonSeparator m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("InsertContent")]
	private RibbonButton H;

	[AccessedThroughProperty("GroupShapes")]
	[CompilerGenerated]
	private RibbonGroup B;

	[AccessedThroughProperty("Button2")]
	[CompilerGenerated]
	private RibbonButton I;

	internal virtual RibbonTab TabMacabacus
	{
		[CompilerGenerated]
		get
		{
			return this.m_A;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_A = value;
		}
	}

	internal virtual RibbonGroup GroupSlides
	{
		[CompilerGenerated]
		get
		{
			return this.m_A;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_A = value;
		}
	}

	internal virtual RibbonButton Button1
	{
		[CompilerGenerated]
		get
		{
			return this.m_A;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_A = value;
		}
	}

	internal virtual RibbonMenu MenuInsertSlide
	{
		[CompilerGenerated]
		get
		{
			return this.m_A;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_A = value;
		}
	}

	internal virtual RibbonButton ImportSlidePaste
	{
		[CompilerGenerated]
		get
		{
			return this.B;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			RibbonControlEventHandler value2 = A;
			RibbonButton ribbonButton = this.B;
			if (ribbonButton != null)
			{
				ribbonButton.Click -= value2;
			}
			this.B = value;
			ribbonButton = this.B;
			if (ribbonButton == null)
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
				ribbonButton.Click += value2;
				return;
			}
		}
	}

	internal virtual RibbonButton InsertCover
	{
		[CompilerGenerated]
		get
		{
			return C;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			C = value;
		}
	}

	internal virtual RibbonButton InsertTableOfContents
	{
		[CompilerGenerated]
		get
		{
			return D;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			D = value;
		}
	}

	internal virtual RibbonButton InsertFlysheet
	{
		[CompilerGenerated]
		get
		{
			return E;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			E = value;
		}
	}

	internal virtual RibbonButton InsertLegal
	{
		[CompilerGenerated]
		get
		{
			return F;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			F = value;
		}
	}

	internal virtual RibbonButton InsertContact
	{
		[CompilerGenerated]
		get
		{
			return G;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			G = value;
		}
	}

	internal virtual RibbonMenu MenuSlideTools
	{
		[CompilerGenerated]
		get
		{
			return this.B;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.B = value;
		}
	}

	internal virtual RibbonSeparator Separator1
	{
		[CompilerGenerated]
		get
		{
			return this.m_A;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_A = value;
		}
	}

	internal virtual RibbonButton InsertContent
	{
		[CompilerGenerated]
		get
		{
			return H;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			H = value;
		}
	}

	internal virtual RibbonGroup GroupShapes
	{
		[CompilerGenerated]
		get
		{
			return B;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			B = value;
		}
	}

	internal virtual RibbonButton Button2
	{
		[CompilerGenerated]
		get
		{
			return I;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			I = value;
		}
	}

	[DebuggerNonUserCode]
	public Ribbon2(IContainer container)
		: this()
	{
		container?.Add(this);
	}

	[DebuggerNonUserCode]
	public Ribbon2()
		: base(NG.A.GetRibbonFactory())
	{
		base.Load += A;
		A();
	}

	[DebuggerNonUserCode]
	protected override void Dispose(bool disposing)
	{
		try
		{
			if (!disposing || this.m_A == null)
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
				this.m_A.Dispose();
				return;
			}
		}
		finally
		{
			base.Dispose(disposing);
		}
	}

	[DebuggerStepThrough]
	private void A()
	{
		TabMacabacus = base.Factory.CreateRibbonTab();
		GroupSlides = base.Factory.CreateRibbonGroup();
		Button1 = base.Factory.CreateRibbonButton();
		MenuInsertSlide = base.Factory.CreateRibbonMenu();
		InsertCover = base.Factory.CreateRibbonButton();
		InsertTableOfContents = base.Factory.CreateRibbonButton();
		InsertFlysheet = base.Factory.CreateRibbonButton();
		InsertLegal = base.Factory.CreateRibbonButton();
		InsertContact = base.Factory.CreateRibbonButton();
		Separator1 = base.Factory.CreateRibbonSeparator();
		InsertContent = base.Factory.CreateRibbonButton();
		MenuSlideTools = base.Factory.CreateRibbonMenu();
		ImportSlidePaste = base.Factory.CreateRibbonButton();
		GroupShapes = base.Factory.CreateRibbonGroup();
		Button2 = base.Factory.CreateRibbonButton();
		TabMacabacus.SuspendLayout();
		GroupSlides.SuspendLayout();
		GroupShapes.SuspendLayout();
		TabMacabacus.Groups.Add(GroupSlides);
		TabMacabacus.Groups.Add(GroupShapes);
		TabMacabacus.KeyTip = AH.A(7905);
		TabMacabacus.Label = AH.A(166360);
		TabMacabacus.Name = AH.A(166393);
		GroupSlides.Items.Add(Button1);
		GroupSlides.Items.Add(MenuInsertSlide);
		GroupSlides.Items.Add(MenuSlideTools);
		GroupSlides.Items.Add(ImportSlidePaste);
		GroupSlides.Label = AH.A(114590);
		GroupSlides.Name = AH.A(166418);
		Button1.ControlSize = RibbonControlSize.RibbonControlSizeLarge;
		Button1.ImageName = AH.A(166441);
		Button1.Label = AH.A(166456);
		Button1.Name = AH.A(166473);
		Button1.ShowImage = true;
		MenuInsertSlide.ControlSize = RibbonControlSize.RibbonControlSizeLarge;
		MenuInsertSlide.Items.Add(InsertCover);
		MenuInsertSlide.Items.Add(InsertTableOfContents);
		MenuInsertSlide.Items.Add(InsertFlysheet);
		MenuInsertSlide.Items.Add(InsertLegal);
		MenuInsertSlide.Items.Add(InsertContact);
		MenuInsertSlide.Items.Add(Separator1);
		MenuInsertSlide.Items.Add(InsertContent);
		MenuInsertSlide.Label = AH.A(166488);
		MenuInsertSlide.Name = AH.A(166507);
		MenuInsertSlide.OfficeImageId = AH.A(166538);
		MenuInsertSlide.ShowImage = true;
		InsertCover.Image = OB.TitlePage;
		InsertCover.KeyTip = AH.A(7908);
		InsertCover.Label = AH.A(166555);
		InsertCover.Name = AH.A(166576);
		InsertCover.ShowImage = true;
		InsertTableOfContents.KeyTip = AH.A(7959);
		InsertTableOfContents.Label = AH.A(115674);
		InsertTableOfContents.Name = AH.A(166599);
		InsertTableOfContents.OfficeImageId = AH.A(3199);
		InsertTableOfContents.ShowImage = true;
		InsertFlysheet.Image = OB.Flysheet;
		InsertFlysheet.KeyTip = AH.A(7917);
		InsertFlysheet.Label = AH.A(166642);
		InsertFlysheet.Name = AH.A(166659);
		InsertFlysheet.ShowImage = true;
		InsertLegal.Image = OB.SetPertWeights;
		InsertLegal.KeyTip = AH.A(7935);
		InsertLegal.Label = AH.A(115709);
		InsertLegal.Name = AH.A(166688);
		InsertLegal.ShowImage = true;
		InsertContact.KeyTip = AH.A(7944);
		InsertContact.Label = AH.A(115736);
		InsertContact.Name = AH.A(166711);
		InsertContact.OfficeImageId = AH.A(166738);
		InsertContact.ShowImage = true;
		Separator1.Name = AH.A(166747);
		InsertContent.KeyTip = AH.A(7956);
		InsertContent.Label = AH.A(166768);
		InsertContent.Name = AH.A(166801);
		InsertContent.ShowImage = true;
		MenuSlideTools.ControlSize = RibbonControlSize.RibbonControlSizeLarge;
		MenuSlideTools.KeyTip = AH.A(7944);
		MenuSlideTools.Label = AH.A(166828);
		MenuSlideTools.Name = AH.A(166839);
		MenuSlideTools.OfficeImageId = AH.A(166868);
		MenuSlideTools.ShowImage = true;
		ImportSlidePaste.ControlSize = RibbonControlSize.RibbonControlSizeLarge;
		ImportSlidePaste.Enabled = false;
		ImportSlidePaste.Label = AH.A(166897);
		ImportSlidePaste.Name = AH.A(166918);
		ImportSlidePaste.OfficeImageId = AH.A(166951);
		ImportSlidePaste.ShowImage = true;
		GroupShapes.Items.Add(Button2);
		GroupShapes.Label = AH.A(166970);
		GroupShapes.Name = AH.A(166983);
		Button2.Label = AH.A(167006);
		Button2.Name = AH.A(167006);
		base.Name = AH.A(167021);
		base.RibbonType = AH.A(167036);
		base.Tabs.Add(TabMacabacus);
		TabMacabacus.ResumeLayout(performLayout: false);
		TabMacabacus.PerformLayout();
		GroupSlides.ResumeLayout(performLayout: false);
		GroupSlides.PerformLayout();
		GroupShapes.ResumeLayout(performLayout: false);
		GroupShapes.PerformLayout();
	}

	private void A(object A, RibbonUIEventArgs B)
	{
	}

	private void A(object A, RibbonControlEventArgs B)
	{
	}
}
