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
using Corel.Interop.VGCore;
using SborkaPreflight.ViewModel;


namespace SborkaPreflight.View
{
    /// <summary>
    /// Interaction logic for UserControl1.xaml
    /// </summary>
    public partial class MainUserControl : UserControl
    {
        public MainUserControl(Corel.Interop.VGCore.Application app)
        {
            InitializeComponent();
            EventManager.RegisterClassHandler(typeof(TextBox), UIElement.GotFocusEvent, new RoutedEventHandler(SelectAllText), true);
            DataContext = new MainUserControlWM(app);
        }

        private static void SelectAllText(object sender, RoutedEventArgs e)
        {
            var textBox = e.OriginalSource as TextBox;
            if (textBox != null)
                textBox.SelectAll();
        }

        private async void pListNewOrders_IsVisibleChanged(object sender, DependencyPropertyChangedEventArgs e)
        {
            try
            {
                await Task.Run(() => { System.Threading.Thread.Sleep(50); });
                lbListNewOrders.SelectedIndex = 0;
                Keyboard.Focus(lbListNewOrders.ItemContainerGenerator.ContainerFromIndex(0) as ListBoxItem);
            }
            catch { };
        }

        private void btnAdditionalMenu_IsEnabledChanged(object sender, DependencyPropertyChangedEventArgs e)
        {
            if ((bool)e.NewValue == false)
                ((System.Windows.Controls.Primitives.ToggleButton)sender).IsChecked = false;
        }
    }
}
