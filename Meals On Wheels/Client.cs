using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Meals_On_Wheels
{
    [Serializable]
    class Client
    {
        public string Name { get; set; }
        public string Address { get; set; }
        public int ColdMeals { get; set; }
        public int HotMeals { get; set; }
        public string SpecialInstructions { get; set; }
        public string PhoneNumber { get; set; }
        public double Latitude { get; set; }
        public double Longitude { get; set; }

        public static Client NewClient()
        {
            NewClientWindow win = new NewClientWindow();
            win.ShowDialog();
            try
            {
                Client newClient = new Client { Name = win.ClientNameTextBox.Text, Address = win.ClientAddressTextBox.Text, ColdMeals = Convert.ToInt32(win.ColdMealsTextBox.Text), HotMeals = Convert.ToInt32(win.HotMealsTextBox.Text), SpecialInstructions = win.SpecialInstructionsTextBox.Text, PhoneNumber = win.PhoneNumberTextBox.Text };
                return newClient;
            }
            catch
            {
                System.Windows.MessageBox.Show("Error entering client. Please confirm you are entering valid data and try again");
                return null;
            }
        }
    }
}
