using System;
using System.CodeDom.Compiler;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Markup;
using System.Xml;
using A;
using ExcelAddIn1.Formulas;
using MacabacusMacros;
using MacabacusMacros.ImportExport;
using MacabacusMacros.Publishing;
using MacabacusMacros.UI;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Publishing.Share;

[DesignerGenerated]
public sealed class wpfShare : UserControl, IComponentConnector
{
	[AccessedThroughProperty("scroller")]
	[CompilerGenerated]
	private ScrollViewer m_A;

	[AccessedThroughProperty("chkReplaceFormulas")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("chkFormulaErrors")]
	private System.Windows.Controls.CheckBox m_B;

	[AccessedThroughProperty("chkRemoveNames")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox m_C;

	[AccessedThroughProperty("chkCleanCells")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox m_D;

	[AccessedThroughProperty("chkRecolorFonts")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox E;

	[AccessedThroughProperty("chkPrintAreas")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox F;

	[AccessedThroughProperty("chkHidePageBreaks")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox G;

	[CompilerGenerated]
	[AccessedThroughProperty("chkHideGridlines")]
	private System.Windows.Controls.CheckBox H;

	[AccessedThroughProperty("chkZoomTo100")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox I;

	[CompilerGenerated]
	[AccessedThroughProperty("chkCellA1")]
	private System.Windows.Controls.CheckBox J;

	[CompilerGenerated]
	[AccessedThroughProperty("chkActFirstSheet")]
	private System.Windows.Controls.CheckBox K;

	[CompilerGenerated]
	[AccessedThroughProperty("chkRemoveHiddenSheets")]
	private System.Windows.Controls.CheckBox L;

	[AccessedThroughProperty("chkBuryHiddenSheets")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox M;

	[AccessedThroughProperty("chkDeleteHiddenRowsCols")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox N;

	[AccessedThroughProperty("chkCollapse")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox O;

	[AccessedThroughProperty("chkExpand")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox P;

	[AccessedThroughProperty("chkRemoveCharts")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox Q;

	[AccessedThroughProperty("chkRemoveWatches")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox R;

	[CompilerGenerated]
	[AccessedThroughProperty("chkRemoveInk")]
	private System.Windows.Controls.CheckBox S;

	[CompilerGenerated]
	[AccessedThroughProperty("chkRemoveNotes")]
	private System.Windows.Controls.CheckBox T;

	[CompilerGenerated]
	[AccessedThroughProperty("chkDocProperties")]
	private System.Windows.Controls.CheckBox U;

	[AccessedThroughProperty("chkCustomXml")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox V;

	[CompilerGenerated]
	[AccessedThroughProperty("chkBreakHyperlinks")]
	private System.Windows.Controls.CheckBox W;

	[AccessedThroughProperty("optThis")]
	[CompilerGenerated]
	private RadioButton m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("optCopy")]
	private RadioButton m_B;

	[CompilerGenerated]
	[AccessedThroughProperty("chkEmail")]
	private System.Windows.Controls.CheckBox X;

	[CompilerGenerated]
	[AccessedThroughProperty("btnClear")]
	private Button m_A;

	[AccessedThroughProperty("btnPrepare")]
	[CompilerGenerated]
	private Button m_B;

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

	internal virtual System.Windows.Controls.CheckBox chkReplaceFormulas
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

	internal virtual System.Windows.Controls.CheckBox chkFormulaErrors
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

	internal virtual System.Windows.Controls.CheckBox chkRemoveNames
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

	internal virtual System.Windows.Controls.CheckBox chkCleanCells
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

	internal virtual System.Windows.Controls.CheckBox chkRecolorFonts
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

	internal virtual System.Windows.Controls.CheckBox chkPrintAreas
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

	internal virtual System.Windows.Controls.CheckBox chkHidePageBreaks
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

	internal virtual System.Windows.Controls.CheckBox chkHideGridlines
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

	internal virtual System.Windows.Controls.CheckBox chkZoomTo100
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

	internal virtual System.Windows.Controls.CheckBox chkCellA1
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

	internal virtual System.Windows.Controls.CheckBox chkActFirstSheet
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

	internal virtual System.Windows.Controls.CheckBox chkRemoveHiddenSheets
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

	internal virtual System.Windows.Controls.CheckBox chkBuryHiddenSheets
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

	internal virtual System.Windows.Controls.CheckBox chkDeleteHiddenRowsCols
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

	internal virtual System.Windows.Controls.CheckBox chkCollapse
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
			RoutedEventHandler value2 = CollapseRowsColsChecked;
			System.Windows.Controls.CheckBox checkBox = O;
			if (checkBox != null)
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
				checkBox.Checked -= value2;
			}
			O = value;
			checkBox = O;
			if (checkBox == null)
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
				checkBox.Checked += value2;
				return;
			}
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkExpand
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
			RoutedEventHandler value2 = ExpandRowsColsChecked;
			System.Windows.Controls.CheckBox checkBox = P;
			if (checkBox != null)
			{
				checkBox.Checked -= value2;
			}
			P = value;
			checkBox = P;
			if (checkBox != null)
			{
				checkBox.Checked += value2;
			}
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkRemoveCharts
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

	internal virtual System.Windows.Controls.CheckBox chkRemoveWatches
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

	internal virtual System.Windows.Controls.CheckBox chkRemoveInk
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

	internal virtual System.Windows.Controls.CheckBox chkRemoveNotes
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

	internal virtual System.Windows.Controls.CheckBox chkDocProperties
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

	internal virtual System.Windows.Controls.CheckBox chkCustomXml
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

	internal virtual System.Windows.Controls.CheckBox chkBreakHyperlinks
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

	internal virtual RadioButton optThis
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

	internal virtual RadioButton optCopy
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

	internal virtual Button btnClear
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
			Button button = this.m_A;
			if (button != null)
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
				button.Click += value2;
				return;
			}
		}
	}

	internal virtual Button btnPrepare
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
			Button button = this.m_B;
			if (button != null)
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
				switch (2)
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
		chkReplaceFormulas.Tag = VH.A(99540);
		chkRecolorFonts.Tag = VH.A(99571);
		chkRemoveNotes.Tag = VH.A(99592);
		chkRemoveNames.Tag = VH.A(99621);
		chkRemoveHiddenSheets.Tag = VH.A(99644);
		chkBuryHiddenSheets.Tag = VH.A(98888);
		chkDeleteHiddenRowsCols.Tag = VH.A(99681);
		chkCollapse.Tag = VH.A(99722);
		chkExpand.Tag = VH.A(98979);
		chkRemoveCharts.Tag = VH.A(99755);
		chkRemoveWatches.Tag = VH.A(99780);
		chkRemoveInk.Tag = VH.A(99106);
		chkFormulaErrors.Tag = VH.A(99807);
		chkPrintAreas.Tag = VH.A(99830);
		chkHidePageBreaks.Tag = VH.A(99125);
		chkHideGridlines.Tag = VH.A(99861);
		chkZoomTo100.Tag = VH.A(99888);
		chkCellA1.Tag = VH.A(98921);
		chkActFirstSheet.Tag = VH.A(98942);
		chkCleanCells.Tag = VH.A(99008);
		chkBreakHyperlinks.Tag = VH.A(99029);
		chkDocProperties.Tag = VH.A(99060);
		chkCustomXml.Tag = VH.A(99087);
		chkEmail.Tag = VH.A(98869);
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
		C();
		base.SizeChanged -= PaneSizeChanged;
	}

	private XmlNode A(XmlDocument A)
	{
		return A.DocumentElement.SelectSingleNode(VH.A(98832));
	}

	private void A()
	{
		XmlDocument settingsXml;
		XmlNode xmlNode;
		try
		{
			settingsXml = KH.A.SettingsXml;
			xmlNode = A(settingsXml);
			try
			{
				string[] array = new string[11]
				{
					VH.A(98869),
					VH.A(98888),
					VH.A(98921),
					VH.A(98942),
					VH.A(98979),
					VH.A(99008),
					VH.A(99029),
					VH.A(99060),
					VH.A(99087),
					VH.A(99106),
					VH.A(99125)
				};
				foreach (string name in array)
				{
					if (xmlNode.Attributes[name] != null)
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
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					XmlAttribute xmlAttribute = settingsXml.CreateAttribute(name);
					xmlAttribute.Value = VH.A(63226);
					xmlNode.Attributes.Append(xmlAttribute);
					xmlAttribute = null;
				}
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			System.Windows.Controls.CheckBox[] array2 = A();
			foreach (System.Windows.Controls.CheckBox checkBox in array2)
			{
				checkBox.IsChecked = Conversions.ToBoolean(xmlNode.Attributes[checkBox.Tag.ToString()].Value);
			}
			if (Conversions.ToBoolean(xmlNode.Attributes[VH.A(99154)].Value))
			{
				while (true)
				{
					switch (5)
					{
					case 0:
						continue;
					}
					optCopy.IsChecked = true;
					break;
				}
			}
			else
			{
				optThis.IsChecked = true;
			}
		}
		catch (Exception ex3)
		{
			ProjectData.SetProjectError(ex3);
			Exception ex4 = ex3;
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
		A(VH.A(99154), ((RadioButton)sender).IsChecked.Value);
	}

	private void A(string A, bool B)
	{
		XmlDocument settingsXml = KH.A.SettingsXml;
		this.A(settingsXml).Attributes[A].Value = B.ToString();
		KH.A.SaveSettings(settingsXml);
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
		optCopy.Checked += SaveRadioSetting;
		optCopy.Unchecked += SaveRadioSetting;
	}

	private void C()
	{
		System.Windows.Controls.CheckBox[] array = A();
		foreach (System.Windows.Controls.CheckBox obj in array)
		{
			obj.Checked -= SaveCheckSetting;
			obj.Unchecked -= SaveCheckSetting;
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
			optCopy.Checked -= SaveRadioSetting;
			optCopy.Unchecked -= SaveRadioSetting;
			return;
		}
	}

	private System.Windows.Controls.CheckBox[] A()
	{
		return new System.Windows.Controls.CheckBox[24]
		{
			chkReplaceFormulas, chkRecolorFonts, chkRemoveNotes, chkRemoveNames, chkRemoveHiddenSheets, chkBuryHiddenSheets, chkDeleteHiddenRowsCols, chkCollapse, chkExpand, chkRemoveCharts,
			chkRemoveWatches, chkRemoveInk, chkFormulaErrors, chkPrintAreas, chkHidePageBreaks, chkHideGridlines, chkZoomTo100, chkCellA1, chkActFirstSheet, chkCleanCells,
			chkBreakHyperlinks, chkDocProperties, chkCustomXml, chkEmail
		};
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
			switch (7)
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

	private void CollapseRowsColsChecked(object sender, RoutedEventArgs e)
	{
		chkExpand.IsChecked = false;
	}

	private void ExpandRowsColsChecked(object sender, RoutedEventArgs e)
	{
		chkCollapse.IsChecked = false;
	}

	private void ExecuteOnEnter(object sender, KeyEventArgs e)
	{
		if (e.Key == Key.Return)
		{
			D();
		}
	}

	private void PrepareToShare_Click(object sender, RoutedEventArgs e)
	{
		D();
	}

	[MethodImpl(MethodImplOptions.NoInlining | MethodImplOptions.NoOptimization)]
	private void D()
	{
		Collection collection = new Collection();
		bool value = chkReplaceFormulas.IsChecked.Value;
		bool value2 = chkRecolorFonts.IsChecked.Value;
		bool value3 = chkRemoveNotes.IsChecked.Value;
		bool value4 = chkRemoveNames.IsChecked.Value;
		bool value5 = chkRemoveHiddenSheets.IsChecked.Value;
		bool value6 = chkBuryHiddenSheets.IsChecked.Value;
		bool value7 = chkDeleteHiddenRowsCols.IsChecked.Value;
		bool value8 = chkCollapse.IsChecked.Value;
		bool value9 = chkExpand.IsChecked.Value;
		bool value10 = chkRemoveCharts.IsChecked.Value;
		bool value11 = chkRemoveWatches.IsChecked.Value;
		bool value12 = chkRemoveInk.IsChecked.Value;
		bool value13 = chkFormulaErrors.IsChecked.Value;
		bool value14 = chkPrintAreas.IsChecked.Value;
		bool value15 = chkHidePageBreaks.IsChecked.Value;
		bool value16 = chkHideGridlines.IsChecked.Value;
		bool value17 = chkZoomTo100.IsChecked.Value;
		bool value18 = chkCellA1.IsChecked.Value;
		bool value19 = chkActFirstSheet.IsChecked.Value;
		bool value20 = chkCleanCells.IsChecked.Value;
		bool value21 = chkBreakHyperlinks.IsChecked.Value;
		bool value22 = chkDocProperties.IsChecked.Value;
		bool value23 = chkCustomXml.IsChecked.Value;
		if (!value && !value2 && !value3)
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
			if (!value4 && !value5)
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
				if (!value6)
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
					if (!value7)
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
						if (!value8)
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
							if (!value9 && !value10 && !value11)
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
								if (!value13)
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
									if (!value14 && !value15 && !value16)
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
										if (!value17)
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
											if (!value18)
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
												if (!value19 && !value20)
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
													if (!value21)
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
														if (!value22)
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
															if (!value23 && !value12)
															{
																B(VH.A(99171));
																return;
															}
														}
													}
												}
											}
										}
									}
								}
							}
						}
					}
				}
			}
		}
		bool value24 = optCopy.IsChecked.Value;
		bool value25 = chkEmail.IsChecked.Value;
		Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
		application.ScreenUpdating = false;
		application.EnableEvents = false;
		application.DisplayAlerts = false;
		try
		{
			Microsoft.Office.Interop.Excel.Workbook activeWorkbook = application.ActiveWorkbook;
			int num = Conversions.ToInteger(NewLateBinding.LateGet(activeWorkbook.ActiveSheet, null, VH.A(48135), new object[0], null, null, null));
			Microsoft.Office.Interop.Excel.Workbook workbook;
			if (value24)
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
				IEnumerator enumerator = default(IEnumerator);
				try
				{
					enumerator = activeWorkbook.Sheets.GetEnumerator();
					while (enumerator.MoveNext())
					{
						object objectValue = RuntimeHelpers.GetObjectValue(enumerator.Current);
						if (!Conversions.ToBoolean(Operators.NotObject(NewLateBinding.LateGet(objectValue, null, VH.A(41367), new object[0], null, null, null))))
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
							break;
						}
						collection.Add(RuntimeHelpers.GetObjectValue(objectValue));
						NewLateBinding.LateSet(objectValue, null, VH.A(41367), new object[1] { XlSheetVisibility.xlSheetVisible }, null, null);
					}
					while (true)
					{
						switch (5)
						{
						case 0:
							break;
						default:
							goto end_IL_04c5;
						}
						continue;
						end_IL_04c5:
						break;
					}
				}
				finally
				{
					if (enumerator is IDisposable)
					{
						while (true)
						{
							switch (6)
							{
							case 0:
								continue;
							}
							(enumerator as IDisposable).Dispose();
							break;
						}
					}
				}
				workbook = Helpers.A(activeWorkbook, MH.A.Application);
				Microsoft.Office.Interop.Excel.Workbook workbook2 = workbook;
				for (int i = workbook2.Worksheets.Count; i >= 2; i = checked(i + -1))
				{
					NewLateBinding.LateCall(workbook2.Worksheets[i], null, VH.A(60691), new object[0], null, null, null, IgnoreReturn: true);
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
				activeWorkbook.Sheets.Copy(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(workbook2.Worksheets[1]));
				((Worksheet)workbook2.Worksheets[1]).Delete();
				workbook2 = null;
			}
			else
			{
				workbook = application.ActiveWorkbook;
			}
			object objectValue2 = RuntimeHelpers.GetObjectValue(workbook.Sheets[num]);
			if (value24)
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
				IEnumerator enumerator2 = default(IEnumerator);
				try
				{
					enumerator2 = collection.GetEnumerator();
					while (enumerator2.MoveNext())
					{
						object objectValue = RuntimeHelpers.GetObjectValue(enumerator2.Current);
						try
						{
							NewLateBinding.LateSet(objectValue, null, VH.A(41367), new object[1] { XlSheetVisibility.xlSheetHidden }, null, null);
							object instance = workbook.Sheets[RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(objectValue, null, VH.A(19019), new object[0], null, null, null))];
							if (value5)
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
								NewLateBinding.LateCall(instance, null, VH.A(60691), new object[0], null, null, null, IgnoreReturn: true);
							}
							else if (value6)
							{
								NewLateBinding.LateSetComplex(instance, null, VH.A(41367), new object[1] { XlSheetVisibility.xlSheetVeryHidden }, null, null, OptimisticSet: false, RValueBase: true);
							}
							else
							{
								NewLateBinding.LateSetComplex(instance, null, VH.A(41367), new object[1] { XlSheetVisibility.xlSheetHidden }, null, null, OptimisticSet: false, RValueBase: true);
							}
							instance = null;
						}
						catch (Exception ex)
						{
							ProjectData.SetProjectError(ex);
							Exception ex2 = ex;
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
							switch (2)
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
			else if (value5)
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
				Base.C(workbook);
			}
			else if (value6)
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
				Base.B(workbook);
			}
			IEnumerator enumerator3 = default(IEnumerator);
			bool flag = default(bool);
			try
			{
				enumerator3 = workbook.Worksheets.GetEnumerator();
				while (enumerator3.MoveNext())
				{
					object objectValue3 = RuntimeHelpers.GetObjectValue(enumerator3.Current);
					if (value7)
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
						Base.H((Worksheet)objectValue3);
					}
					if (value8)
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
						Base.I((Worksheet)objectValue3);
					}
					else if (value9)
					{
						Base.J((Worksheet)objectValue3);
					}
					if (value2)
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
						clsImportExport.RecolorFonts((Range)NewLateBinding.LateGet(objectValue3, null, VH.A(82416), new object[0], null, null, null), false);
					}
					if (value3)
					{
						Base.A((Worksheet)objectValue3);
					}
					if (value10)
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
						Base.B((Worksheet)objectValue3);
					}
					if (value12)
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
						Base.C((Worksheet)objectValue3);
					}
					if (value13 && !flag)
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
							Range range = null;
							try
							{
								range = Base.A((Worksheet)objectValue3);
							}
							catch (Exception ex3)
							{
								ProjectData.SetProjectError(ex3);
								Exception ex4 = ex3;
								ProjectData.ClearProjectError();
							}
							if (range != null)
							{
								while (true)
								{
									switch (3)
									{
									case 0:
										continue;
									}
									flag = true;
									break;
								}
							}
						}
						catch (Exception ex5)
						{
							ProjectData.SetProjectError(ex5);
							Exception ex6 = ex5;
							ProjectData.ClearProjectError();
						}
					}
					if (value14)
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
						Base.K((Worksheet)objectValue3);
					}
					if (value15)
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
						Base.D((Worksheet)objectValue3);
					}
					if (value16)
					{
						Base.E((Worksheet)objectValue3);
					}
					if (value17)
					{
						Base.F((Worksheet)objectValue3);
					}
					if (value21)
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
						Base.L((Worksheet)objectValue3);
					}
					if (value18)
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
						Base.G((Worksheet)objectValue3);
					}
					if (value20)
					{
						Base.A((Range)NewLateBinding.LateGet(objectValue3, null, VH.A(82416), new object[0], null, null, null));
					}
				}
				while (true)
				{
					switch (5)
					{
					case 0:
						break;
					default:
						goto end_IL_0954;
					}
					continue;
					end_IL_0954:
					break;
				}
			}
			finally
			{
				if (enumerator3 is IDisposable)
				{
					while (true)
					{
						switch (7)
						{
						case 0:
							continue;
						}
						(enumerator3 as IDisposable).Dispose();
						break;
					}
				}
			}
			if (value19)
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
				IEnumerator enumerator4 = default(IEnumerator);
				try
				{
					enumerator4 = workbook.Sheets.GetEnumerator();
					while (true)
					{
						if (enumerator4.MoveNext())
						{
							object objectValue4 = RuntimeHelpers.GetObjectValue(enumerator4.Current);
							if (!Operators.ConditionalCompareObjectEqual(NewLateBinding.LateGet(objectValue4, null, VH.A(41367), new object[0], null, null, null), XlSheetVisibility.xlSheetVisible, TextCompare: false))
							{
								continue;
							}
							while (true)
							{
								switch (3)
								{
								case 0:
									continue;
								}
								NewLateBinding.LateCall(objectValue4, null, VH.A(39985), new object[0], null, null, null, IgnoreReturn: true);
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
								goto end_IL_0a17;
							}
							continue;
							end_IL_0a17:
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
							switch (1)
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
			else
			{
				NewLateBinding.LateCall(objectValue2, null, VH.A(39985), new object[0], null, null, null, IgnoreReturn: true);
			}
			if (value)
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
				Base.A(workbook);
			}
			if (value10)
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
				Base.DeleteCharts(workbook);
			}
			if (value4)
			{
				if (value)
				{
					try
					{
						IEnumerator enumerator5 = workbook.Names.GetEnumerator();
						try
						{
							while (enumerator5.MoveNext())
							{
								((Name)enumerator5.Current).Delete();
							}
							while (true)
							{
								switch (1)
								{
								case 0:
									break;
								default:
									goto end_IL_0ac7;
								}
								continue;
								end_IL_0ac7:
								break;
							}
						}
						finally
						{
							IDisposable disposable = enumerator5 as IDisposable;
							if (disposable != null)
							{
								disposable.Dispose();
							}
						}
					}
					catch (Exception ex7)
					{
						ProjectData.SetProjectError(ex7);
						Exception ex8 = ex7;
						ProjectData.ClearProjectError();
					}
				}
				else
				{
					try
					{
						IEnumerator enumerator6 = default(IEnumerator);
						try
						{
							enumerator6 = workbook.Names.GetEnumerator();
							while (enumerator6.MoveNext())
							{
								Name obj = (Name)enumerator6.Current;
								ExcelAddIn1.Formulas.Names.Unapply(obj);
								obj.Delete();
							}
							while (true)
							{
								switch (6)
								{
								case 0:
									break;
								default:
									goto end_IL_0b37;
								}
								continue;
								end_IL_0b37:
								break;
							}
						}
						finally
						{
							if (enumerator6 is IDisposable)
							{
								while (true)
								{
									switch (7)
									{
									case 0:
										continue;
									}
									(enumerator6 as IDisposable).Dispose();
									break;
								}
							}
						}
					}
					catch (Exception ex9)
					{
						ProjectData.SetProjectError(ex9);
						Exception ex10 = ex9;
						ProjectData.ClearProjectError();
					}
					if (value13)
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
						if (!flag)
						{
							try
							{
								IEnumerator enumerator7 = default(IEnumerator);
								try
								{
									enumerator7 = workbook.Worksheets.GetEnumerator();
									while (enumerator7.MoveNext())
									{
										object objectValue5 = RuntimeHelpers.GetObjectValue(enumerator7.Current);
										Range range = null;
										try
										{
											range = Base.A((Worksheet)objectValue5);
										}
										catch (Exception ex11)
										{
											ProjectData.SetProjectError(ex11);
											Exception ex12 = ex11;
											ProjectData.ClearProjectError();
										}
										if (range == null)
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
											flag = true;
											break;
										}
										break;
									}
								}
								finally
								{
									if (enumerator7 is IDisposable)
									{
										while (true)
										{
											switch (1)
											{
											case 0:
												continue;
											}
											(enumerator7 as IDisposable).Dispose();
											break;
										}
									}
								}
							}
							catch (Exception ex13)
							{
								ProjectData.SetProjectError(ex13);
								Exception ex14 = ex13;
								ProjectData.ClearProjectError();
							}
						}
					}
				}
			}
			if (value11)
			{
				Base.A();
			}
			if (value22)
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
				Base.A(workbook, value24);
			}
			if (value23)
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
				Base.D(workbook);
			}
			if (flag)
			{
				B(VH.A(99200));
			}
			else if (value25)
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
				if (!value24)
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
					workbook.Sheets.Copy(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
				}
				Microsoft.Office.Interop.Excel.Workbook activeWorkbook2 = application.ActiveWorkbook;
				if (Helpers.A(activeWorkbook2))
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
						string text = Base.SaveTempWorkbook(activeWorkbook, activeWorkbook2);
						string text2 = ((Operators.CompareString(activeWorkbook.Path, "", TextCompare: false) == 0) ? activeWorkbook.Name : Path.GetFileNameWithoutExtension(activeWorkbook.Name));
						string text3 = A(text, text2 + Path.GetExtension(text), C: true);
						clsPublish.AttachToEmail(text3, false);
						try
						{
							FileSystem.Kill(text3);
						}
						catch (Exception ex15)
						{
							ProjectData.SetProjectError(ex15);
							Exception ex16 = ex15;
							ProjectData.ClearProjectError();
						}
					}
					catch (Exception ex17)
					{
						ProjectData.SetProjectError(ex17);
						Exception ex18 = ex17;
						A(ex18.Message);
						ProjectData.ClearProjectError();
					}
				}
				else if (value24)
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
					activeWorkbook2.Close(false, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
				}
				activeWorkbook2 = null;
			}
			clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)8, VH.A(99275));
		}
		catch (Exception ex19)
		{
			ProjectData.SetProjectError(ex19);
			Exception ex20 = ex19;
			A(ex20.Message);
			ProjectData.ClearProjectError();
		}
		finally
		{
			Microsoft.Office.Interop.Excel.Workbook activeWorkbook = null;
			Microsoft.Office.Interop.Excel.Workbook workbook = null;
			Range range = null;
			object objectValue2 = null;
			application.CutCopyMode = (XlCutCopyMode)0;
			application.DisplayAlerts = true;
			application.EnableEvents = true;
			application.ScreenUpdating = true;
			List<string> list = new List<string>();
			_ = null;
			if (list.Any())
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
				Share.ShowResults(list);
			}
			list = null;
		}
		application = null;
		collection = null;
	}

	[MethodImpl(MethodImplOptions.NoInlining | MethodImplOptions.NoOptimization)]
	private string A(string A, string B, bool C)
	{
		if (C)
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
			string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(B);
			B = Forms.InputBox(VH.A(40448), VH.A(99308), fileNameWithoutExtension);
			if (Operators.CompareString(B, string.Empty, TextCompare: false) != 0)
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
				if (B.Length != 0)
				{
					goto IL_0075;
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
			B = fileNameWithoutExtension;
			goto IL_0075;
		}
		goto IL_0087;
		IL_0075:
		B += Path.GetExtension(A);
		goto IL_0087;
		IL_0087:
		B = Path.Combine(global::A.I.A.FileSystem.SpecialDirectories.Temp, B);
		try
		{
			FileSystem.Kill(B);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		global::A.I.A.FileSystem.RenameFile(A, Path.GetFileName(B));
		return B;
	}

	private void A(string A)
	{
		Forms.ErrorMessage(System.Windows.Window.GetWindow(this), A);
	}

	private void B(string A)
	{
		Forms.WarningMessage(System.Windows.Window.GetWindow(this), A);
	}

	private void C(string A)
	{
		Forms.InfoMessage(System.Windows.Window.GetWindow(this), A);
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
			Uri resourceLocator = new Uri(VH.A(99425), UriKind.Relative);
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
				switch (7)
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
				switch (4)
				{
				case 0:
					break;
				default:
					chkReplaceFormulas = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 3)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					chkFormulaErrors = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 4)
		{
			chkRemoveNames = (System.Windows.Controls.CheckBox)target;
			return;
		}
		if (connectionId == 5)
		{
			chkCleanCells = (System.Windows.Controls.CheckBox)target;
			return;
		}
		if (connectionId == 6)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					chkRecolorFonts = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 7)
		{
			chkPrintAreas = (System.Windows.Controls.CheckBox)target;
			return;
		}
		if (connectionId == 8)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					chkHidePageBreaks = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 9)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					chkHideGridlines = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 10)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					chkZoomTo100 = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 11)
		{
			chkCellA1 = (System.Windows.Controls.CheckBox)target;
			return;
		}
		if (connectionId == 12)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					chkActFirstSheet = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 13)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					chkRemoveHiddenSheets = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 14)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					chkBuryHiddenSheets = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 15)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					chkDeleteHiddenRowsCols = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 16)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					chkCollapse = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 17)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					chkExpand = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 18)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					chkRemoveCharts = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 19)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					chkRemoveWatches = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 20)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					chkRemoveInk = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 21)
		{
			chkRemoveNotes = (System.Windows.Controls.CheckBox)target;
			return;
		}
		if (connectionId == 22)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					chkDocProperties = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 23)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					chkCustomXml = (System.Windows.Controls.CheckBox)target;
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
					chkBreakHyperlinks = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 25)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					optThis = (RadioButton)target;
					return;
				}
			}
		}
		if (connectionId == 26)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					optCopy = (RadioButton)target;
					return;
				}
			}
		}
		switch (connectionId)
		{
		case 27:
			chkEmail = (System.Windows.Controls.CheckBox)target;
			break;
		case 28:
			btnClear = (Button)target;
			break;
		case 29:
			while (true)
			{
				switch (4)
				{
				case 0:
					continue;
				}
				btnPrepare = (Button)target;
				return;
			}
		default:
			this.m_A = true;
			break;
		}
	}

	void IComponentConnector.Connect(int connectionId, object target)
	{
		//ILSpy generated this explicit interface implementation from .override directive in System_Windows_Markup_IComponentConnector_Connect
		this.System_Windows_Markup_IComponentConnector_Connect(connectionId, target);
	}
}
