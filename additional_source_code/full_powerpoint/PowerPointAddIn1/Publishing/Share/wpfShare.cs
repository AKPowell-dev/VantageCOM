using System;
using System.CodeDom.Compiler;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Markup;
using System.Xml;
using A;
using MacabacusMacros;
using MacabacusMacros.Publishing;
using MacabacusMacros.UI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.Colors;
using PowerPointAddIn1.Links;
using PowerPointAddIn1.MasterShapes;
using PowerPointAddIn1.Shapes;
using PowerPointAddIn1.Slides;
using PowerPointAddIn1.Template;

namespace PowerPointAddIn1.Publishing.Share;

[DesignerGenerated]
public sealed class wpfShare : System.Windows.Controls.UserControl, IComponentConnector
{
	private struct LF
	{
		public bool A;

		public bool B;

		public bool C;

		public bool D;

		public bool E;

		public bool F;

		public bool G;

		public bool H;

		public bool I;

		public bool J;

		public bool K;

		public bool L;

		public bool M;

		public bool N;

		public bool O;

		public bool P;

		public bool Q;

		public bool R;

		public bool S;

		public bool T;

		public bool U;

		public bool V;

		public bool W;

		public bool X;

		public bool Y;
	}

	private static int m_A;

	private static int m_B;

	private static int m_C;

	private static int m_D;

	private static int E;

	private static int F;

	private static int G;

	private static int H;

	private static int I;

	private static int J;

	private static int K;

	private static int L;

	private static int M;

	private static int N;

	private static int O;

	private static int P;

	private static int Q;

	private static int R;

	private static int S;

	private static int T;

	private static int U;

	private static int V;

	private static int W;

	[AccessedThroughProperty("scroller")]
	[CompilerGenerated]
	private ScrollViewer m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("chkHiddenSlides")]
	private System.Windows.Controls.CheckBox m_A;

	[AccessedThroughProperty("chkHiddenShapes")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox m_B;

	[CompilerGenerated]
	[AccessedThroughProperty("chkOffSlideShapes")]
	private System.Windows.Controls.CheckBox m_C;

	[AccessedThroughProperty("chkUnusedLayouts")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox m_D;

	[CompilerGenerated]
	[AccessedThroughProperty("chkMasterShapes")]
	private System.Windows.Controls.CheckBox E;

	[AccessedThroughProperty("chkStyles")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox F;

	[AccessedThroughProperty("chkSections")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox G;

	[AccessedThroughProperty("chkAnimations")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox H;

	[AccessedThroughProperty("chkTransitions")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox I;

	[CompilerGenerated]
	[AccessedThroughProperty("chkInk")]
	private System.Windows.Controls.CheckBox J;

	[CompilerGenerated]
	[AccessedThroughProperty("chkSpeakerNotes")]
	private System.Windows.Controls.CheckBox K;

	[AccessedThroughProperty("chkSlideComments")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox L;

	[AccessedThroughProperty("chkTags")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox M;

	[CompilerGenerated]
	[AccessedThroughProperty("chkAltText")]
	private System.Windows.Controls.CheckBox N;

	[AccessedThroughProperty("chkDocProperties")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox O;

	[CompilerGenerated]
	[AccessedThroughProperty("chkCustomXml")]
	private System.Windows.Controls.CheckBox P;

	[CompilerGenerated]
	[AccessedThroughProperty("chkBreakHyperlinks")]
	private System.Windows.Controls.CheckBox Q;

	[CompilerGenerated]
	[AccessedThroughProperty("chkConvertOLEs")]
	private System.Windows.Controls.CheckBox R;

	[CompilerGenerated]
	[AccessedThroughProperty("chkConvertCharts")]
	private System.Windows.Controls.CheckBox S;

	[AccessedThroughProperty("chkFreezeColors")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox T;

	[CompilerGenerated]
	[AccessedThroughProperty("chkFixGrayscale")]
	private System.Windows.Controls.CheckBox U;

	[CompilerGenerated]
	[AccessedThroughProperty("chkLockSlides")]
	private System.Windows.Controls.CheckBox V;

	[AccessedThroughProperty("radLockShapes")]
	[CompilerGenerated]
	private System.Windows.Controls.RadioButton m_A;

	[AccessedThroughProperty("radLockImages")]
	[CompilerGenerated]
	private System.Windows.Controls.RadioButton m_B;

	[AccessedThroughProperty("radShapeLocked")]
	[CompilerGenerated]
	private System.Windows.Controls.RadioButton m_C;

	[AccessedThroughProperty("chkMarkFinal")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox W;

	[CompilerGenerated]
	[AccessedThroughProperty("optThis")]
	private System.Windows.Controls.RadioButton m_D;

	[CompilerGenerated]
	[AccessedThroughProperty("optCopy")]
	private System.Windows.Controls.RadioButton E;

	[AccessedThroughProperty("chkEmail")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox X;

	[AccessedThroughProperty("btnClear")]
	[CompilerGenerated]
	private System.Windows.Controls.Button m_A;

	[AccessedThroughProperty("btnPrepare")]
	[CompilerGenerated]
	private System.Windows.Controls.Button m_B;

	private bool m_A;

	internal virtual ScrollViewer scroller
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

	internal virtual System.Windows.Controls.CheckBox chkHiddenSlides
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

	internal virtual System.Windows.Controls.CheckBox chkHiddenShapes
	{
		[CompilerGenerated]
		get
		{
			return this.m_B;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_B = value;
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkOffSlideShapes
	{
		[CompilerGenerated]
		get
		{
			return this.m_C;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_C = value;
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkUnusedLayouts
	{
		[CompilerGenerated]
		get
		{
			return this.m_D;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_D = value;
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkMasterShapes
	{
		[CompilerGenerated]
		get
		{
			return this.E;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.E = value;
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkStyles
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

	internal virtual System.Windows.Controls.CheckBox chkSections
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

	internal virtual System.Windows.Controls.CheckBox chkAnimations
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

	internal virtual System.Windows.Controls.CheckBox chkTransitions
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

	internal virtual System.Windows.Controls.CheckBox chkInk
	{
		[CompilerGenerated]
		get
		{
			return J;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			J = value;
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkSpeakerNotes
	{
		[CompilerGenerated]
		get
		{
			return K;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			K = value;
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkSlideComments
	{
		[CompilerGenerated]
		get
		{
			return L;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			L = value;
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkTags
	{
		[CompilerGenerated]
		get
		{
			return M;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			M = value;
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkAltText
	{
		[CompilerGenerated]
		get
		{
			return N;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			N = value;
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkDocProperties
	{
		[CompilerGenerated]
		get
		{
			return O;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			O = value;
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkCustomXml
	{
		[CompilerGenerated]
		get
		{
			return P;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			P = value;
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkBreakHyperlinks
	{
		[CompilerGenerated]
		get
		{
			return Q;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			Q = value;
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkConvertOLEs
	{
		[CompilerGenerated]
		get
		{
			return R;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			R = value;
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkConvertCharts
	{
		[CompilerGenerated]
		get
		{
			return S;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			S = value;
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkFreezeColors
	{
		[CompilerGenerated]
		get
		{
			return T;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			T = value;
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkFixGrayscale
	{
		[CompilerGenerated]
		get
		{
			return U;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			U = value;
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkLockSlides
	{
		[CompilerGenerated]
		get
		{
			return V;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			V = value;
		}
	}

	internal virtual System.Windows.Controls.RadioButton radLockShapes
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

	internal virtual System.Windows.Controls.RadioButton radLockImages
	{
		[CompilerGenerated]
		get
		{
			return this.m_B;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_B = value;
		}
	}

	internal virtual System.Windows.Controls.RadioButton radShapeLocked
	{
		[CompilerGenerated]
		get
		{
			return this.m_C;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_C = value;
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkMarkFinal
	{
		[CompilerGenerated]
		get
		{
			return W;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			W = value;
		}
	}

	internal virtual System.Windows.Controls.RadioButton optThis
	{
		[CompilerGenerated]
		get
		{
			return this.m_D;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_D = value;
		}
	}

	internal virtual System.Windows.Controls.RadioButton optCopy
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

	internal virtual System.Windows.Controls.CheckBox chkEmail
	{
		[CompilerGenerated]
		get
		{
			return X;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			X = value;
		}
	}

	internal virtual System.Windows.Controls.Button btnClear
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
			RoutedEventHandler value2 = ClearAllCheckBoxes;
			System.Windows.Controls.Button button = this.m_A;
			if (button != null)
			{
				button.Click -= value2;
			}
			this.m_A = value;
			button = this.m_A;
			if (button == null)
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
				button.Click += value2;
				return;
			}
		}
	}

	internal virtual System.Windows.Controls.Button btnPrepare
	{
		[CompilerGenerated]
		get
		{
			return this.m_B;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			RoutedEventHandler value2 = PrepareToShare_Click;
			System.Windows.Controls.Button button = this.m_B;
			if (button != null)
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
				button.Click -= value2;
			}
			this.m_B = value;
			button = this.m_B;
			if (button == null)
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
				button.Click += value2;
				return;
			}
		}
	}

	public wpfShare()
	{
		base.Unloaded += wpfShare_Unloaded;
		base.KeyDown += ExecuteOnEnter;
		InitializeComponent();
		chkSpeakerNotes.Tag = AH.A(103378);
		chkSlideComments.Tag = AH.A(103403);
		chkHiddenSlides.Tag = AH.A(103430);
		chkHiddenShapes.Tag = AH.A(103455);
		chkMasterShapes.Tag = AH.A(103480);
		chkStyles.Tag = AH.A(103517);
		chkOffSlideShapes.Tag = AH.A(103542);
		chkTags.Tag = AH.A(103571);
		chkAltText.Tag = AH.A(103592);
		chkUnusedLayouts.Tag = AH.A(103619);
		chkSections.Tag = AH.A(103658);
		chkAnimations.Tag = AH.A(103687);
		chkTransitions.Tag = AH.A(103720);
		chkInk.Tag = AH.A(103755);
		chkBreakHyperlinks.Tag = AH.A(103774);
		chkFreezeColors.Tag = AH.A(103805);
		chkFixGrayscale.Tag = AH.A(103830);
		chkConvertOLEs.Tag = AH.A(103855);
		chkConvertCharts.Tag = AH.A(103878);
		chkDocProperties.Tag = AH.A(103905);
		chkCustomXml.Tag = AH.A(103932);
		chkLockSlides.Tag = AH.A(103951);
		chkMarkFinal.Tag = AH.A(103972);
		chkEmail.Tag = AH.A(103991);
	}

	private void wpfShare_Unloaded(object sender, RoutedEventArgs e)
	{
		C();
	}

	private void PaneSizeChanged(object sender, SizeChangedEventArgs e)
	{
		Panes.PaneSizeChanged(scroller, e);
	}

	public void ShowPane()
	{
		A();
		B();
		base.SizeChanged += PaneSizeChanged;
	}

	public void HidePane()
	{
		optThis.Checked -= TargetChanged;
		optCopy.Checked -= TargetChanged;
		C();
		base.SizeChanged -= PaneSizeChanged;
	}

	private XmlNode A(XmlDocument A)
	{
		return A.DocumentElement.SelectSingleNode(AH.A(101733));
	}

	private void A()
	{
		XmlDocument settingsXml;
		XmlNode xmlNode;
		try
		{
			settingsXml = KG.A.SettingsXml;
			xmlNode = A(settingsXml);
			System.Windows.Controls.CheckBox[] array = A();
			foreach (System.Windows.Controls.CheckBox checkBox in array)
			{
				checkBox.IsChecked = Conversions.ToBoolean(xmlNode.Attributes[checkBox.Tag.ToString()].Value);
			}
			optThis.Checked += TargetChanged;
			optCopy.Checked += TargetChanged;
			if (Conversions.ToBoolean(xmlNode.Attributes[AH.A(101778)].Value))
			{
				optCopy.IsChecked = true;
			}
			else
			{
				optThis.IsChecked = true;
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			A(AH.A(101795));
			ProjectData.ClearProjectError();
		}
		settingsXml = null;
		xmlNode = null;
	}

	private void SaveCheckSetting(object sender, RoutedEventArgs e)
	{
		System.Windows.Controls.CheckBox checkBox = (System.Windows.Controls.CheckBox)sender;
		A(checkBox.Tag.ToString(), checkBox.IsChecked.Value);
		checkBox = null;
	}

	private void SaveRadioSetting(object sender, RoutedEventArgs e)
	{
		A(AH.A(101778), ((System.Windows.Controls.RadioButton)sender).IsChecked.Value);
	}

	private void A(string A, bool B)
	{
		XmlDocument settingsXml = KG.A.SettingsXml;
		this.A(settingsXml).Attributes[A].Value = B.ToString();
		KG.A.SaveSettings(settingsXml);
		settingsXml = null;
	}

	private void B()
	{
		System.Windows.Controls.CheckBox[] array = A();
		foreach (System.Windows.Controls.CheckBox obj in array)
		{
			obj.Checked += SaveCheckSetting;
			obj.Unchecked += SaveCheckSetting;
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
			optCopy.Checked += SaveRadioSetting;
			optCopy.Unchecked += SaveRadioSetting;
			return;
		}
	}

	private void C()
	{
		System.Windows.Controls.CheckBox[] array = A();
		foreach (System.Windows.Controls.CheckBox obj in array)
		{
			obj.Checked -= SaveCheckSetting;
			obj.Unchecked -= SaveCheckSetting;
		}
		optCopy.Checked -= SaveRadioSetting;
		optCopy.Unchecked -= SaveRadioSetting;
	}

	private System.Windows.Controls.CheckBox[] A()
	{
		return new System.Windows.Controls.CheckBox[24]
		{
			chkSpeakerNotes, chkSlideComments, chkHiddenSlides, chkHiddenShapes, chkMasterShapes, chkStyles, chkOffSlideShapes, chkTags, chkAltText, chkUnusedLayouts,
			chkSections, chkAnimations, chkTransitions, chkInk, chkBreakHyperlinks, chkFreezeColors, chkFixGrayscale, chkConvertOLEs, chkConvertCharts, chkDocProperties,
			chkCustomXml, chkLockSlides, chkMarkFinal, chkEmail
		};
	}

	private void TargetChanged(object sender, RoutedEventArgs e)
	{
		if (optCopy.IsChecked == true)
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
					chkMarkFinal.IsChecked = false;
					chkMarkFinal.IsEnabled = false;
					return;
				}
			}
		}
		chkMarkFinal.IsEnabled = true;
	}

	private void ClearAllCheckBoxes(object sender, RoutedEventArgs e)
	{
		System.Windows.Controls.CheckBox[] array = A();
		for (int i = 0; i < array.Length; i = checked(i + 1))
		{
			array[i].IsChecked = false;
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
			return;
		}
	}

	private void ExecuteOnEnter(object sender, System.Windows.Input.KeyEventArgs e)
	{
		if (e.Key != Key.Return)
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
			D();
			return;
		}
	}

	private void PrepareToShare_Click(object sender, RoutedEventArgs e)
	{
		D();
	}

	[MethodImpl(MethodImplOptions.NoInlining | MethodImplOptions.NoOptimization)]
	private void D()
	{
		LF lF = default(LF);
		lF = new LF
		{
			A = chkSpeakerNotes.IsChecked.Value,
			B = chkSlideComments.IsChecked.Value,
			C = chkHiddenSlides.IsChecked.Value,
			D = chkHiddenShapes.IsChecked.Value,
			E = chkMasterShapes.IsChecked.Value,
			F = chkStyles.IsChecked.Value,
			G = chkOffSlideShapes.IsChecked.Value,
			H = chkTags.IsChecked.Value,
			I = chkAltText.IsChecked.Value,
			J = chkUnusedLayouts.IsChecked.Value,
			K = chkSections.IsChecked.Value,
			L = chkAnimations.IsChecked.Value,
			M = chkTransitions.IsChecked.Value,
			N = chkInk.IsChecked.Value,
			O = chkBreakHyperlinks.IsChecked.Value,
			P = chkFreezeColors.IsChecked.Value,
			Q = chkFixGrayscale.IsChecked.Value,
			R = chkConvertOLEs.IsChecked.Value,
			S = chkConvertCharts.IsChecked.Value,
			T = chkDocProperties.IsChecked.Value,
			U = chkCustomXml.IsChecked.Value,
			V = chkLockSlides.IsChecked.Value,
			W = radLockImages.IsChecked.Value,
			X = radShapeLocked.IsChecked.Value,
			Y = chkMarkFinal.IsChecked.Value
		};
		bool value = optCopy.IsChecked.Value;
		bool value2 = chkEmail.IsChecked.Value;
		wpfShare.m_A = 0;
		wpfShare.m_B = 0;
		wpfShare.m_C = 0;
		wpfShare.m_D = 0;
		wpfShare.E = 0;
		wpfShare.F = 0;
		wpfShare.G = 0;
		wpfShare.H = 0;
		wpfShare.I = 0;
		wpfShare.J = 0;
		wpfShare.K = 0;
		wpfShare.L = 0;
		wpfShare.M = 0;
		wpfShare.N = 0;
		wpfShare.O = 0;
		wpfShare.P = 0;
		wpfShare.Q = 0;
		wpfShare.R = 0;
		wpfShare.S = 0;
		wpfShare.T = 0;
		wpfShare.U = 0;
		wpfShare.V = 0;
		wpfShare.W = 0;
		Microsoft.Office.Interop.PowerPoint.Application application = NG.A.Application;
		Microsoft.Office.Interop.PowerPoint.Presentation presentation = application.ActivePresentation;
		string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(presentation.Name);
		if (value)
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
			presentation = A(presentation);
		}
		application.StartNewUndoEntry();
		_ = null;
		int num = presentation.Designs.Count;
		checked
		{
			while (num >= 1)
			{
				Design design = presentation.Designs[num];
				Master slideMaster = design.SlideMaster;
				A(slideMaster.Shapes, null, presentation, lF);
				for (int i = slideMaster.CustomLayouts.Count; i >= 1; i += -1)
				{
					CustomLayout customLayout = slideMaster.CustomLayouts[i];
					if (lF.J)
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
						try
						{
							customLayout.Delete();
							wpfShare.R++;
						}
						catch (Exception ex)
						{
							ProjectData.SetProjectError(ex);
							Exception ex2 = ex;
							ProjectData.ClearProjectError();
							goto IL_04ce;
						}
						continue;
					}
					goto IL_04ce;
					IL_04ce:
					A(customLayout.Shapes, null, presentation, lF);
					if (lF.O)
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
						for (int j = customLayout.Hyperlinks.Count; j >= 1; j += -1)
						{
							PowerPointAddIn1.Links.Hyperlinks.C(customLayout.Hyperlinks[j]);
							wpfShare.Q++;
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
					customLayout = null;
				}
				while (true)
				{
					switch (4)
					{
					case 0:
						continue;
					}
					if (lF.J)
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
						if (slideMaster.CustomLayouts.Count == 0)
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
							try
							{
								design.Delete();
								wpfShare.S++;
							}
							catch (Exception ex3)
							{
								ProjectData.SetProjectError(ex3);
								Exception ex4 = ex3;
								ProjectData.ClearProjectError();
							}
						}
					}
					slideMaster = null;
					num += -1;
					break;
				}
			}
			IEnumerator enumerator = default(IEnumerator);
			IEnumerator enumerator2 = default(IEnumerator);
			while (true)
			{
				switch (4)
				{
				case 0:
					continue;
				}
				for (int k = presentation.Slides.Count; k >= 1; k += -1)
				{
					Slide slide = presentation.Slides[k];
					Slide slide2 = slide;
					if (lF.C)
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
						if (slide2.SlideShowTransition.Hidden == MsoTriState.msoTrue)
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
							slide2.Delete();
							wpfShare.m_C++;
							continue;
						}
					}
					if (slide2.HasNotesPage == MsoTriState.msoTrue)
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
						SlideRange notesPage = slide2.NotesPage;
						if (lF.A)
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
							for (int l = notesPage.Shapes.Count; l >= 1; l += -1)
							{
								Microsoft.Office.Interop.PowerPoint.Shape shape = notesPage.Shapes[l];
								Microsoft.Office.Interop.PowerPoint.Shape shape2 = shape;
								if (shape2.Type == MsoShapeType.msoPlaceholder)
								{
									if (shape2.PlaceholderFormat.Type == PpPlaceholderType.ppPlaceholderBody)
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
										shape2.TextFrame2.DeleteText();
										wpfShare.m_A++;
									}
									try
									{
										A(shape, null, presentation, lF);
									}
									catch (Exception ex5)
									{
										ProjectData.SetProjectError(ex5);
										Exception ex6 = ex5;
										clsReporting.LogException(ex6);
										ProjectData.ClearProjectError();
									}
								}
								else
								{
									shape2.Delete();
									wpfShare.m_A++;
								}
								shape2 = null;
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
						else
						{
							try
							{
								A(notesPage.Shapes, null, presentation, lF);
							}
							catch (Exception ex7)
							{
								ProjectData.SetProjectError(ex7);
								Exception ex8 = ex7;
								clsReporting.LogException(ex8);
								ProjectData.ClearProjectError();
							}
						}
						notesPage = null;
					}
					if (lF.B)
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
						for (int m = slide.Comments.Count; m >= 1; m += -1)
						{
							slide.Comments[m].Delete();
							wpfShare.m_B++;
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
					if (lF.M)
					{
						try
						{
							if (slide.SlideShowTransition.EntryEffect != PpEntryEffect.ppEffectNone)
							{
								while (true)
								{
									switch (2)
									{
									case 0:
										continue;
									}
									slide.SlideShowTransition.EntryEffect = PpEntryEffect.ppEffectNone;
									wpfShare.J++;
									break;
								}
							}
						}
						catch (Exception ex9)
						{
							ProjectData.SetProjectError(ex9);
							Exception ex10 = ex9;
							ProjectData.ClearProjectError();
						}
					}
					if (lF.H)
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
						A(slide2.Tags);
					}
					A(slide2.Shapes, slide, presentation, lF);
					if (lF.O)
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
						for (int n = slide2.Hyperlinks.Count; n >= 1; n += -1)
						{
							PowerPointAddIn1.Links.Hyperlinks.C(slide2.Hyperlinks[n]);
							wpfShare.Q++;
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
					slide2 = null;
				}
				while (true)
				{
					switch (7)
					{
					case 0:
						continue;
					}
					if (lF.V)
					{
						if (!lF.X)
						{
							Protection.LockSlidesAll(presentation, lF.W);
						}
						wpfShare.W = presentation.Slides.Count;
					}
					if (lF.H)
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
						A(presentation.Tags);
					}
					if (lF.K)
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
						for (int num2 = presentation.SectionProperties.Count; num2 >= 1; num2 += -1)
						{
							presentation.SectionProperties.Delete(num2, deleteSlides: false);
							wpfShare.P++;
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
					if (lF.P)
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
						Freeze.A(A: true);
					}
					if (lF.T)
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
						if (!value)
						{
							{
								enumerator = ((IEnumerable)presentation.CustomDocumentProperties).GetEnumerator();
								try
								{
									while (enumerator.MoveNext())
									{
										DocumentProperty documentProperty = (DocumentProperty)enumerator.Current;
										try
										{
											documentProperty.Delete();
											wpfShare.T++;
										}
										catch (Exception ex11)
										{
											ProjectData.SetProjectError(ex11);
											Exception ex12 = ex11;
											ProjectData.ClearProjectError();
										}
									}
									while (true)
									{
										switch (3)
										{
										case 0:
											break;
										default:
											goto end_IL_09f0;
										}
										continue;
										end_IL_09f0:
										break;
									}
								}
								finally
								{
									IDisposable disposable = enumerator as IDisposable;
									if (disposable != null)
									{
										disposable.Dispose();
									}
								}
							}
						}
						try
						{
							enumerator2 = ((IEnumerable)presentation.BuiltInDocumentProperties).GetEnumerator();
							while (enumerator2.MoveNext())
							{
								DocumentProperty documentProperty2 = (DocumentProperty)enumerator2.Current;
								try
								{
									switch (documentProperty2.Type)
									{
									case MsoDocProperties.msoPropertyTypeString:
										documentProperty2.Value = string.Empty;
										wpfShare.U++;
										break;
									case MsoDocProperties.msoPropertyTypeDate:
										try
										{
											if (!(documentProperty2.Value is DateTime))
											{
												break;
											}
											while (true)
											{
												switch (6)
												{
												case 0:
													continue;
												}
												documentProperty2.Value = DateTime.MinValue;
												wpfShare.U++;
												break;
											}
											break;
										}
										catch (Exception ex13)
										{
											ProjectData.SetProjectError(ex13);
											Exception ex14 = ex13;
											ProjectData.ClearProjectError();
										}
										break;
									case MsoDocProperties.msoPropertyTypeNumber:
									case MsoDocProperties.msoPropertyTypeBoolean:
									case MsoDocProperties.msoPropertyTypeFloat:
										break;
									}
								}
								catch (Exception ex15)
								{
									ProjectData.SetProjectError(ex15);
									Exception ex16 = ex15;
									ProjectData.ClearProjectError();
								}
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
					if (lF.U)
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
						for (int num3 = presentation.CustomXMLParts.Count; num3 >= 1; num3 += -1)
						{
							try
							{
								presentation.CustomXMLParts[num3].Delete();
								wpfShare.V++;
							}
							catch (Exception ex17)
							{
								ProjectData.SetProjectError(ex17);
								Exception ex18 = ex17;
								ProjectData.ClearProjectError();
							}
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
					if (lF.Y)
					{
						try
						{
							presentation.Final = true;
						}
						catch (Exception ex19)
						{
							ProjectData.SetProjectError(ex19);
							Exception ex20 = ex19;
							if (value)
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
								B(AH.A(101981));
							}
							else
							{
								A(ex20.Message);
							}
							ProjectData.ClearProjectError();
						}
					}
					if (value2)
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
						string text = Forms.InputBox(AH.A(5874), AH.A(102042), fileNameWithoutExtension);
						if (Operators.CompareString(text, string.Empty, TextCompare: false) != 0 && text.Length > 0)
						{
							if (presentation.Path.Length > 0)
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
								text += Path.GetExtension(presentation.Name);
							}
							else
							{
								text += AH.A(102167);
							}
							string text2 = Path.Combine(NB.A.FileSystem.SpecialDirectories.Temp, text);
							presentation.SaveCopyAs(text2);
							clsPublish.AttachToEmail(text2, false);
							try
							{
								FileSystem.Kill(text2);
							}
							catch (Exception ex21)
							{
								ProjectData.SetProjectError(ex21);
								Exception ex22 = ex21;
								ProjectData.ClearProjectError();
							}
						}
					}
					List<string> list = new List<string>();
					List<string> list2 = list;
					if (lF.A)
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
						list2.Add(AH.A(94493) + wpfShare.m_A + AH.A(102178));
					}
					if (lF.B)
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
						list2.Add(AH.A(94493) + wpfShare.m_B + AH.A(102241));
					}
					if (lF.C)
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
						list2.Add(AH.A(94493) + wpfShare.m_C + AH.A(102274));
					}
					if (lF.D)
					{
						list2.Add(AH.A(94493) + wpfShare.m_D + AH.A(102305));
					}
					if (lF.E)
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
						if (!lF.D)
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
							list2.Add(AH.A(94493) + wpfShare.E + AH.A(102336));
						}
					}
					if (lF.F)
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
						if (!lF.D)
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
							list2.Add(AH.A(94493) + wpfShare.F + AH.A(102365));
						}
					}
					if (lF.G)
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
						list2.Add(AH.A(94493) + wpfShare.G + AH.A(102382));
					}
					if (lF.I)
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
						list2.Add(AH.A(102419) + wpfShare.H + AH.A(72774));
					}
					if (lF.H)
					{
						list2.Add(AH.A(94493) + wpfShare.O + AH.A(102464));
					}
					if (lF.L)
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
						list2.Add(AH.A(94493) + wpfShare.I + AH.A(102537));
					}
					if (lF.M)
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
						list2.Add(AH.A(94493) + wpfShare.J + AH.A(102562));
					}
					if (lF.N)
					{
						list2.Add(AH.A(94493) + wpfShare.K + AH.A(102589));
					}
					if (lF.J)
					{
						list2.Add(AH.A(94493) + wpfShare.R + AH.A(102618));
						if (wpfShare.S > 0)
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
							list2.Add(AH.A(94493) + wpfShare.S + AH.A(102651));
						}
					}
					if (lF.K)
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
						list2.Add(AH.A(94493) + wpfShare.P + AH.A(102696));
					}
					if (lF.O)
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
						list2.Add(AH.A(102717) + wpfShare.Q + AH.A(102730));
					}
					if (lF.Q)
					{
						list2.Add(AH.A(102755) + wpfShare.N + AH.A(72774));
					}
					if (lF.P)
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
						list2.Add(AH.A(102794));
					}
					if (lF.R)
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
						list2.Add(AH.A(73404) + wpfShare.L + AH.A(102847));
					}
					if (lF.S)
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
						list2.Add(AH.A(73404) + wpfShare.M + AH.A(102914));
					}
					if (lF.T)
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
						list2.Add(AH.A(94493) + wpfShare.T + AH.A(102955));
						list2.Add(AH.A(103012) + wpfShare.U + AH.A(103029));
					}
					if (lF.U)
					{
						list2.Add(AH.A(94493) + wpfShare.V + AH.A(103090));
					}
					if (lF.V)
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
						list2.Add(AH.A(103127) + wpfShare.W + AH.A(103150));
					}
					if (lF.Y)
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
						if (presentation.Final)
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
							list2.Add(AH.A(103167));
						}
					}
					list2 = null;
					if (list.Count > 0)
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
						Share.ShowResults(list);
					}
					clsReporting.LogActivity((ActivityApp)2, (ActivityCategory)4, AH.A(103220));
					presentation = null;
					Design design = null;
					Slide slide = null;
					Microsoft.Office.Interop.PowerPoint.Shape shape = null;
					list = null;
					return;
				}
			}
		}
	}

	private void A(Microsoft.Office.Interop.PowerPoint.Shapes A, Slide B, Microsoft.Office.Interop.PowerPoint.Presentation C, LF D)
	{
		for (int i = A.Count; i >= 1; i = checked(i + -1))
		{
			wpfShare.A(A[i], B, C, D);
		}
	}

	private static void A(Microsoft.Office.Interop.PowerPoint.Shape A, Slide B, Microsoft.Office.Interop.PowerPoint.Presentation C, LF D)
	{
		int num = wpfShare.A(A);
		Microsoft.Office.Interop.PowerPoint.Shape shape = A;
		checked
		{
			if (D.D)
			{
				if (shape.Visible == MsoTriState.msoFalse)
				{
					shape.Delete();
					wpfShare.m_D += num;
					return;
				}
			}
			else
			{
				if (D.E && PowerPointAddIn1.MasterShapes.Base.A(A))
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
					if (shape.Visible == MsoTriState.msoFalse)
					{
						while (true)
						{
							switch (5)
							{
							case 0:
								break;
							default:
								shape.Delete();
								wpfShare.E += num;
								return;
							}
						}
					}
				}
				if (D.F)
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
					if (Styles.HasStyleShapeName(A))
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
						if (shape.Visible == MsoTriState.msoFalse)
						{
							while (true)
							{
								switch (7)
								{
								case 0:
									break;
								default:
									shape.Delete();
									wpfShare.F += num;
									return;
								}
							}
						}
					}
				}
			}
			if (D.G)
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
				if (!(shape.Left < shape.Width * -1f) && !(shape.Top < shape.Height * -1f))
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
					if (!(shape.Left > C.PageSetup.SlideWidth))
					{
						if (!(shape.Top > C.PageSetup.SlideHeight))
						{
							goto IL_0160;
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
				}
				shape.Delete();
				wpfShare.G += num;
				return;
			}
			goto IL_0160;
		}
		IL_0329:
		if (D.V)
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
			if (D.X)
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
				try
				{
					B.Application.ActiveWindow.View.GotoSlide(B.SlideIndex);
					NewLateBinding.LateSetComplex(A, null, AH.A(69417), new object[1] { true }, null, null, OptimisticSet: false, RValueBase: true);
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					ProjectData.ClearProjectError();
				}
			}
		}
		goto IL_03a1;
		IL_01ad:
		if (D.H)
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
			wpfShare.A(shape.Tags);
		}
		checked
		{
			if (D.I)
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
					if (Operators.CompareString(shape.AlternativeText, string.Empty, TextCompare: false) != 0)
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
						if (shape.AlternativeText.Length > 0)
						{
							goto IL_0257;
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
					if (Operators.CompareString(shape.Title, string.Empty, TextCompare: false) != 0)
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
						if (shape.Title.Length > 0)
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
							goto IL_0257;
						}
					}
					goto end_IL_01df;
					IL_0257:
					shape.AlternativeText = "";
					shape.Title = "";
					wpfShare.H++;
					end_IL_01df:;
				}
				catch (Exception ex3)
				{
					ProjectData.SetProjectError(ex3);
					Exception ex4 = ex3;
					ProjectData.ClearProjectError();
				}
			}
			if (D.L)
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
				try
				{
					A.PickupAnimation();
					Animation.Remove(A);
					wpfShare.I++;
				}
				catch (Exception ex5)
				{
					ProjectData.SetProjectError(ex5);
					Exception ex6 = ex5;
					ProjectData.ClearProjectError();
				}
			}
			if (B != null)
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
				if (D.V)
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
					if (D.W)
					{
						goto IL_0329;
					}
				}
				if (D.R)
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
					ConvertToPicture.ConvertEmbedded(B, A, ref wpfShare.L);
				}
				if (D.S)
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
					ConvertToPicture.ConvertChart(B, A, ref wpfShare.M);
				}
				goto IL_0329;
			}
			goto IL_03a1;
		}
		IL_03a1:
		checked
		{
			if (D.Q)
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
				Grayscale.A(A);
				wpfShare.N++;
			}
			shape = null;
			if (A.Type != MsoShapeType.msoGroup)
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
				List<Microsoft.Office.Interop.PowerPoint.Shape> list = new List<Microsoft.Office.Interop.PowerPoint.Shape>();
				foreach (Microsoft.Office.Interop.PowerPoint.Shape groupItem in A.GroupItems)
				{
					list.Add(groupItem);
				}
				for (int i = list.Count - 1; i >= 0; i += -1)
				{
					wpfShare.A(list[i], B, C, D);
				}
				while (true)
				{
					switch (3)
					{
					case 0:
						continue;
					}
					list = null;
					return;
				}
			}
		}
		IL_0160:
		checked
		{
			if (D.N)
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
				if (shape.Type != MsoShapeType.msoInk)
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
					if (shape.Type != MsoShapeType.msoInkComment)
					{
						goto IL_01ad;
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
				shape.Delete();
				wpfShare.K += num;
				return;
			}
			goto IL_01ad;
		}
	}

	private static int A(Microsoft.Office.Interop.PowerPoint.Shape A)
	{
		int num = 0;
		checked
		{
			if (A.Type == MsoShapeType.msoGroup)
			{
				{
					IEnumerator enumerator = A.GroupItems.GetEnumerator();
					try
					{
						while (enumerator.MoveNext())
						{
							Microsoft.Office.Interop.PowerPoint.Shape a = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator.Current;
							num += wpfShare.A(a);
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
							break;
						}
					}
					finally
					{
						IDisposable disposable = enumerator as IDisposable;
						if (disposable != null)
						{
							disposable.Dispose();
						}
					}
				}
			}
			else
			{
				num++;
			}
			return num;
		}
	}

	private static void A(Tags A)
	{
		Tags tags = A;
		checked
		{
			for (int i = tags.Count; i >= 1; i += -1)
			{
				tags.Delete(tags.Name(i));
				wpfShare.O++;
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
				tags = null;
				return;
			}
		}
	}

	private Microsoft.Office.Interop.PowerPoint.Presentation A(Microsoft.Office.Interop.PowerPoint.Presentation A)
	{
		Microsoft.Office.Interop.PowerPoint.Application application = NG.A.Application;
		Microsoft.Office.Interop.PowerPoint.Presentation presentation = Create.NewBlankPresentation(application);
		Design design = presentation.Designs[1];
		PageSetup pageSetup = presentation.PageSetup;
		pageSetup.SlideHeight = A.PageSetup.SlideHeight;
		pageSetup.SlideWidth = A.PageSetup.SlideWidth;
		_ = null;
		A.Slides.Range(RuntimeHelpers.GetObjectValue(Missing.Value)).Copy();
		application.CommandBars.ExecuteMso(AH.A(58900));
		System.Windows.Forms.Application.DoEvents();
		try
		{
			design.Delete();
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		design = null;
		return presentation;
	}

	private void A(string A)
	{
		Forms.ErrorMessage(Window.GetWindow(this), A);
	}

	private void B(string A)
	{
		Forms.WarningMessage(Window.GetWindow(this), A);
	}

	private void C(string A)
	{
		Forms.InfoMessage(Window.GetWindow(this), A);
	}

	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	[DebuggerNonUserCode]
	public void InitializeComponent()
	{
		if (this.m_A)
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
			this.m_A = true;
			Uri resourceLocator = new Uri(AH.A(103253), UriKind.Relative);
			System.Windows.Application.LoadComponent(this, resourceLocator);
			return;
		}
	}

	void IComponentConnector.InitializeComponent()
	{
		//ILSpy generated this explicit interface implementation from .override directive in InitializeComponent
		this.InitializeComponent();
	}

	[EditorBrowsable(EditorBrowsableState.Never)]
	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	[DebuggerNonUserCode]
	public void System_Windows_Markup_IComponentConnector_Connect(int connectionId, object target)
	{
		if (connectionId == 1)
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
					scroller = (ScrollViewer)target;
					return;
				}
			}
		}
		if (connectionId == 2)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					chkHiddenSlides = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 3)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					chkHiddenShapes = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 4)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					chkOffSlideShapes = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 5)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					chkUnusedLayouts = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 6)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					chkMasterShapes = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 7)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					chkStyles = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 8)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					chkSections = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 9)
		{
			chkAnimations = (System.Windows.Controls.CheckBox)target;
			return;
		}
		if (connectionId == 10)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					chkTransitions = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 11)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					chkInk = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 12)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					chkSpeakerNotes = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 13)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					chkSlideComments = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 14)
		{
			chkTags = (System.Windows.Controls.CheckBox)target;
			return;
		}
		if (connectionId == 15)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					chkAltText = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 16)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					chkDocProperties = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 17)
		{
			chkCustomXml = (System.Windows.Controls.CheckBox)target;
			return;
		}
		if (connectionId == 18)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					chkBreakHyperlinks = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 19)
		{
			chkConvertOLEs = (System.Windows.Controls.CheckBox)target;
			return;
		}
		if (connectionId == 20)
		{
			chkConvertCharts = (System.Windows.Controls.CheckBox)target;
			return;
		}
		if (connectionId == 21)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					chkFreezeColors = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 22)
		{
			chkFixGrayscale = (System.Windows.Controls.CheckBox)target;
			return;
		}
		if (connectionId == 23)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					chkLockSlides = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 24)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					radLockShapes = (System.Windows.Controls.RadioButton)target;
					return;
				}
			}
		}
		if (connectionId == 25)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					radLockImages = (System.Windows.Controls.RadioButton)target;
					return;
				}
			}
		}
		if (connectionId == 26)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					radShapeLocked = (System.Windows.Controls.RadioButton)target;
					return;
				}
			}
		}
		if (connectionId == 27)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					chkMarkFinal = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 28)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					optThis = (System.Windows.Controls.RadioButton)target;
					return;
				}
			}
		}
		if (connectionId == 29)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					optCopy = (System.Windows.Controls.RadioButton)target;
					return;
				}
			}
		}
		if (connectionId == 30)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					chkEmail = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 31)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					btnClear = (System.Windows.Controls.Button)target;
					return;
				}
			}
		}
		if (connectionId == 32)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					btnPrepare = (System.Windows.Controls.Button)target;
					return;
				}
			}
		}
		this.m_A = true;
	}

	void IComponentConnector.Connect(int connectionId, object target)
	{
		//ILSpy generated this explicit interface implementation from .override directive in System_Windows_Markup_IComponentConnector_Connect
		this.System_Windows_Markup_IComponentConnector_Connect(connectionId, target);
	}
}
