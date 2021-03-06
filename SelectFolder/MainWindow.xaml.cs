﻿using System;
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
		private string documents;
		public MainWindow()
		{
			InitializeComponent();
			//	アプリケーションのドキュメントフォルダ
			documents = System.IO.Path.Combine( Environment.GetFolderPath( Environment.SpecialFolder.MyDocuments ), "Wankuma" );
			System.IO.Directory.CreateDirectory( documents );
		}

		private void OnClick_SelectFolder( object sender, RoutedEventArgs e )
		{
			var dlg = new PickupFolder.PickupFolderDialog();
			dlg.SelectedPath = EditSelFolder.Text;
			dlg.InitialFolder = documents;
			dlg.Title = "フォルダを選択してください";
			dlg.AddPlace( documents, PickupFolder.FDAP.TOP );
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
