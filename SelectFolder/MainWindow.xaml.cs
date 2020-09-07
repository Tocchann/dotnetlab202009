using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace SelectFolder
{
	/// <summary>
	/// Interaction logic for MainWindow.xaml
	/// </summary>
	public partial class MainWindow : Window
	{
		public MainWindow()
		{
			InitializeComponent();
		}

		private void OnClick_SelectFolder( object sender, RoutedEventArgs e )
		{
			var dlg = new PickupFolder.PickupFolderDialog();
			dlg.SelectedPath = EditSelFolder.Text;
			dlg.InitialFolder = AppDomain.CurrentDomain.BaseDirectory;
			dlg.Title = "フォルダを選択してください";
			dlg.AddPlace( @"C:\Projects\Source\Tocchann\dotnetlab202009", PickupFolder.FDAP.TOP );
			if( dlg.ShowDialog() )
			{
				EditSelFolder.Text = dlg.SelectedPath;
			}
		}
		private void OnClick_Cancel( object sender, RoutedEventArgs e )
		{
			Close();
		}
	}
}
