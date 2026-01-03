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
using TocBuilder_dotnet_framework.ViewModels;

namespace TocBuilder_dotnet_framework
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void PreviewGroupBox_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            if (DataContext is MainViewModel vm)
            {
                // Вычитаем отступы GroupBox (Header + Padding)
                double viewportWidth = Math.Max(0, e.NewSize.Width);
                double viewportHeight = Math.Max(0, e.NewSize.Height - 40);

                vm.UpdatePreviewViewportSize(viewportWidth, viewportHeight);
            }
        }
    }
}
