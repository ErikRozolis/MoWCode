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
using System.Windows.Shapes;

namespace Meals_On_Wheels
{
    /// <summary>
    /// Interaction logic for DeleteRouteConfirmation.xaml
    /// </summary>
    public partial class DeleteRouteConfirmation : Window
    {
        public bool Confirmation = false;
        public DeleteRouteConfirmation()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            Confirmation = true;
            this.Close();
        }
    }
}
