using System;
using System.CodeDom.Compiler;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Markup;
using A;
using MacabacusMacros;
using MacabacusMacros.Aiwa;
using MacabacusMacros.Aiwa.UI;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1.Aiwa;

[DesignerGenerated]
public sealed class wpfWorkspace : UserControl, INotifyPropertyChanged, IComponentConnector
{
	[CompilerGenerated]
	private PropertyChangedEventHandler m_A;

	private BackgroundWorker m_A;

	private string m_A;

	private string m_B;

	private string C;

	private UserControl m_A;

	private UserControl m_B;

	[CompilerGenerated]
	private List<string> m_A;

	private int m_A;

	private int m_B;

	[CompilerGenerated]
	private bool m_A;

	private bool m_B;

	private bool C;

	private bool D;

	[CompilerGenerated]
	private wpfHome m_A;

	[CompilerGenerated]
	private JsonFeature m_A;

	private string D;

	private string E;

	private string F;

	private Dictionary<string, string> m_A;

	private bool E;

	[CompilerGenerated]
	private JsonLanguage m_A;

	private TextResult m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("cbxLanguages")]
	private ComboBox m_A;

	private bool F;

	public string InputText
	{
		get
		{
			return this.m_B;
		}
		set
		{
			this.m_B = value;
			A(AH.A(10324));
			ButtonLabel = Title;
			A();
		}
	}

	public string OutputText
	{
		get
		{
			return this.C;
		}
		set
		{
			this.C = value;
			A(AH.A(10343));
		}
	}

	public UserControl ProcessingView
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			A(AH.A(10364));
		}
	}

	public UserControl ErrorView
	{
		get
		{
			return this.m_B;
		}
		set
		{
			this.m_B = value;
			A(AH.A(10162));
		}
	}

	private List<string> Suggestions
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

	public int SuggestionIndex
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			A(AH.A(10393));
		}
	}

	public int TotalSuggestions
	{
		get
		{
			return this.m_B;
		}
		set
		{
			this.m_B = value;
			A(AH.A(10424));
		}
	}

	private bool AllowMultipleSuggestions
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

	public bool ShowSuggestions
	{
		get
		{
			return this.m_B;
		}
		set
		{
			this.m_B = value;
			A(AH.A(10457));
		}
	}

	public bool IsPreviousSuggestionEnabled
	{
		get
		{
			return C;
		}
		set
		{
			C = value;
			A(AH.A(10488));
		}
	}

	public bool IsNextSuggestionEnabled
	{
		get
		{
			return this.D;
		}
		set
		{
			this.D = value;
			A(AH.A(10543));
		}
	}

	private wpfHome Home
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

	private JsonFeature FeatureData
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

	public string Title
	{
		get
		{
			return D;
		}
		set
		{
			D = value;
			A(AH.A(10590));
		}
	}

	public string Subtitle
	{
		get
		{
			return this.E;
		}
		set
		{
			this.E = value;
			A(AH.A(10601));
		}
	}

	public string ButtonLabel
	{
		get
		{
			return this.F;
		}
		set
		{
			this.F = value;
			A(AH.A(10618));
		}
	}

	public Dictionary<string, string> Languages
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			A(AH.A(10641));
		}
	}

	public bool ShowLanguages
	{
		get
		{
			return E;
		}
		set
		{
			E = value;
			A(AH.A(10660));
		}
	}

	private JsonLanguage OutputLanguage
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

	internal virtual ComboBox cbxLanguages
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
	}

	public wpfWorkspace(wpfHome parent, JsonFeature featureData)
	{
		//IL_0049: Unknown result type (might be due to invalid IL or missing references)
		//IL_004e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0050: Unknown result type (might be due to invalid IL or missing references)
		//IL_005c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0074: Unknown result type (might be due to invalid IL or missing references)
		//IL_0087: Unknown result type (might be due to invalid IL or missing references)
		//IL_008c: Unknown result type (might be due to invalid IL or missing references)
		//IL_009c: Unknown result type (might be due to invalid IL or missing references)
		//IL_00ac: Unknown result type (might be due to invalid IL or missing references)
		//IL_00ae: Invalid comparison between Unknown and I4
		base.Unloaded += ViewUnloaded;
		this.m_A = null;
		this.m_B = null;
		E = false;
		OutputLanguage = null;
		InitializeComponent();
		Home = parent;
		FeatureData = featureData;
		UiCopy uiCopy = Pane.GetUiCopy(featureData);
		Title = uiCopy.Title;
		Subtitle = uiCopy.Subtitle;
		this.m_A = AH.A(10921) + uiCopy.Title;
		FeatureType featureType = featureData.FeatureType;
		AllowMultipleSuggestions = !((IEnumerable<FeatureType>)(object)new FeatureType[2]
		{
			(FeatureType)4,
			(FeatureType)5
		}).Contains(featureType);
		if ((int)featureType == 5)
		{
			Languages = Translate.Languages();
			cbxLanguages.SelectedIndex = 0;
			ShowLanguages = true;
		}
		A();
		this.m_A = new BackgroundWorker();
		this.m_A.WorkerSupportsCancellation = true;
		this.m_A.DoWork += DoAiWorkAsync;
		this.m_A.RunWorkerCompleted += AiWorkCompleted;
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
			switch (7)
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

	private void ViewUnloaded(object sender, RoutedEventArgs e)
	{
		Suggestions = null;
		Home = null;
	}

	private void ProcessText(object sender, RoutedEventArgs e)
	{
		//IL_0014: Unknown result type (might be due to invalid IL or missing references)
		//IL_0019: Unknown result type (might be due to invalid IL or missing references)
		//IL_0032: Unknown result type (might be due to invalid IL or missing references)
		//IL_0037: Unknown result type (might be due to invalid IL or missing references)
		//IL_003a: Invalid comparison between Unknown and I4
		//IL_0075: Unknown result type (might be due to invalid IL or missing references)
		//IL_007f: Expected O, but got Unknown
		string empty = string.Empty;
		if (!Text.ValidateTextInput(InputText, FeatureData.FeatureType, ref empty))
		{
			B(empty);
			return;
		}
		if ((int)FeatureData.FeatureType == 5)
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
			if (OutputLanguage == null)
			{
				B(AH.A(10687));
				return;
			}
		}
		ProcessingView = (UserControl)new wpfProcessing((Action<object, RoutedEventArgs>)CancelAiWork);
		this.m_A.RunWorkerAsync();
	}

	private void OutputLanguageChanged(object sender, SelectionChangedEventArgs e)
	{
		ComboBox comboBox = sender as ComboBox;
		OutputLanguage = ((comboBox.SelectedIndex > 0) ? Translate.GetOutputLanguage(comboBox) : null);
		comboBox = null;
	}

	private void DoAiWorkAsync(object sender, DoWorkEventArgs e)
	{
		//IL_000f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0015: Invalid comparison between Unknown and I4
		//IL_0051: Unknown result type (might be due to invalid IL or missing references)
		//IL_0056: Unknown result type (might be due to invalid IL or missing references)
		//IL_0065: Unknown result type (might be due to invalid IL or missing references)
		//IL_0075: Expected O, but got Unknown
		this.m_A = null;
		JsonLanguage val = null;
		if ((int)FeatureData.FeatureType == 5 && base.Dispatcher.Invoke([SpecialName] () => cbxLanguages.SelectedIndex) > -1)
		{
			KeyValuePair<string, string> keyValuePair = base.Dispatcher.Invoke([SpecialName] () =>
			{
				object selectedItem = cbxLanguages.SelectedItem;
				if (selectedItem == null)
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
							return default(KeyValuePair<string, string>);
						}
					}
				}
				return (KeyValuePair<string, string>)selectedItem;
			});
			val = new JsonLanguage
			{
				Code = keyValuePair.Key,
				Name = keyValuePair.Value
			};
		}
		this.m_A = Text.GenerateText(FeatureData, this.m_A, InputText, val);
	}

	private void AiWorkCompleted(object sender, RunWorkerCompletedEventArgs e)
	{
		ProcessingView = null;
		if (e.Error != null)
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
				B(e.Error.Message);
				return;
			}
		}
		if (!e.Cancelled)
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
			TextResult a = this.m_A;
			if (a == null)
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
			}
			else if (a.WasCancelled)
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
				goto IL_0078;
			}
			TextResult a2 = this.m_A;
			object obj;
			if (a2 == null)
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
				obj = null;
			}
			else
			{
				obj = a2.OutputText;
			}
			if (obj == null)
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
				obj = "";
			}
			OutputText = (string)obj;
			TextResult a3 = this.m_A;
			object obj2;
			if (a3 == null)
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
				obj2 = null;
			}
			else
			{
				obj2 = a3.ErrorMsg;
			}
			bool flag = !modFunctionsStr.IsBlank((string)obj2);
			if (!modFunctionsStr.IsBlank(OutputText) && !flag)
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
				Suggestions.Add(OutputText);
				SuggestionIndex = Suggestions.Count;
				TotalSuggestions = Suggestions.Count;
				IsPreviousSuggestionEnabled = Suggestions.Count > 1;
				IsNextSuggestionEnabled = false;
				ShowSuggestions = AllowMultipleSuggestions;
				ButtonLabel = AH.A(10783);
			}
			clsReporting.LogActivity((ActivityApp)2, (ActivityCategory)6, this.m_A);
			if (flag)
			{
				B(this.m_A.ErrorMsg);
				return;
			}
			TextResult a4 = this.m_A;
			if (!modFunctionsStr.IsBlank((a4 != null) ? a4.UserMsg : null))
			{
				this.m_A.ShowUserMsg();
			}
			return;
		}
		goto IL_0078;
		IL_0078:
		B(AH.A(10744));
	}

	private void CancelAiWork(object sender, RoutedEventArgs e)
	{
		if (!this.m_A.IsBusy)
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
			this.m_A.CancelAsync();
			return;
		}
	}

	private void B(string A)
	{
		//IL_000e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0018: Expected O, but got Unknown
		ErrorView = (UserControl)new wpfError(A, (Action<object, RoutedEventArgs>)DismissError);
	}

	private void DismissError(object sender, RoutedEventArgs e)
	{
		ErrorView = null;
	}

	private void PreviousSuggestion(object sender, RoutedEventArgs e)
	{
		A(-1);
		IsPreviousSuggestionEnabled = SuggestionIndex > 1;
		IsNextSuggestionEnabled = true;
	}

	private void NextSuggestion(object sender, RoutedEventArgs e)
	{
		A(1);
		IsPreviousSuggestionEnabled = true;
		IsNextSuggestionEnabled = SuggestionIndex < Suggestions.Count;
	}

	private void A(int A)
	{
		checked
		{
			SuggestionIndex += A;
			OutputText = Suggestions[SuggestionIndex - 1];
		}
	}

	private void A()
	{
		ButtonLabel = Title;
		OutputText = null;
		SuggestionIndex = 0;
		TotalSuggestions = 0;
		ShowSuggestions = false;
		IsPreviousSuggestionEnabled = false;
		IsNextSuggestionEnabled = false;
		Suggestions = new List<string>();
	}

	private void GoHome(object sender, RoutedEventArgs e)
	{
		this.m_A.CancelAsync();
		this.m_A = null;
		Home.A();
	}

	private void CopyText(object sender, RoutedEventArgs e)
	{
		Text.Copy(OutputText);
	}

	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	[DebuggerNonUserCode]
	public void InitializeComponent()
	{
		if (!F)
		{
			F = true;
			Uri resourceLocator = new Uri(AH.A(10812), UriKind.Relative);
			Application.LoadComponent(this, resourceLocator);
		}
	}

	void IComponentConnector.InitializeComponent()
	{
		//ILSpy generated this explicit interface implementation from .override directive in InitializeComponent
		this.InitializeComponent();
	}

	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	[DebuggerNonUserCode]
	[EditorBrowsable(EditorBrowsableState.Never)]
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
					((Button)target).Click += GoHome;
					return;
				}
			}
		}
		if (connectionId == 2)
		{
			cbxLanguages = (ComboBox)target;
			cbxLanguages.SelectionChanged += OutputLanguageChanged;
			return;
		}
		if (connectionId == 3)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					((Button)target).Click += ProcessText;
					return;
				}
			}
		}
		switch (connectionId)
		{
		case 4:
			while (true)
			{
				switch (2)
				{
				case 0:
					continue;
				}
				((Button)target).Click += CopyText;
				return;
			}
		case 5:
			((Button)target).Click += NextSuggestion;
			break;
		case 6:
			((Button)target).Click += PreviousSuggestion;
			break;
		default:
			F = true;
			break;
		}
	}

	void IComponentConnector.Connect(int connectionId, object target)
	{
		//ILSpy generated this explicit interface implementation from .override directive in System_Windows_Markup_IComponentConnector_Connect
		this.System_Windows_Markup_IComponentConnector_Connect(connectionId, target);
	}

	[SpecialName]
	[CompilerGenerated]
	private int A()
	{
		return cbxLanguages.SelectedIndex;
	}

	[SpecialName]
	[CompilerGenerated]
	private KeyValuePair<string, string> A()
	{
		object selectedItem = cbxLanguages.SelectedItem;
		if (selectedItem == null)
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
					return default(KeyValuePair<string, string>);
				}
			}
		}
		return (KeyValuePair<string, string>)selectedItem;
	}
}
