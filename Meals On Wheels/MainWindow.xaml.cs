using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Runtime.Serialization.Formatters.Binary;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Web;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using static Meals_On_Wheels.GoogleMapsJSON;

namespace Meals_On_Wheels
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        List<Route> routeList;
        Route currentRoute;
        Button createNewClientButton;

        public MainWindow()
        {
            InitializeComponent();
            LoadingText.Visibility = Visibility.Hidden;
            routeList = new List<Route>();
            if (File.Exists("Routes.GOAT"))
            {
                DeserializeRoutes();
            }
            RepopulateRouteList();
        }

        private void SerializeRoutes()
        {
            Stream stream = File.OpenWrite("Routes.GOAT");
            BinaryFormatter bf = new BinaryFormatter();
            bf.Serialize(stream, routeList);
            stream.Close();
        }

        private void DeserializeRoutes()
        {
            Stream stream = File.OpenRead("Routes.GOAT");
            BinaryFormatter bf = new BinaryFormatter();
            routeList = (List<Route>) bf.Deserialize(stream);
            Console.WriteLine("Deserializing");
            stream.Close();
        }

        private void NewRouteButton_Click(object sender, RoutedEventArgs e)
        {
            Route newRoute = Route.NewRoute();
            if(routeList.Find(x=> x.Name == newRoute.Name)==null)
            {
                currentRoute = newRoute;
                if (currentRoute != null)
                {
                    //clear formatting
                    ClientList.ColumnDefinitions.Clear();
                    ClientList.RowDefinitions.Clear();

                    //reformat for new route
                    ColumnDefinition ClientNameColumn = new ColumnDefinition { Name = "ClientName", Width = new GridLength(360) };
                    ColumnDefinition ClientAddressColumn = new ColumnDefinition { Name = "ClientAddress", Width = new GridLength(360) };
                    ColumnDefinition ClientEditColumn = new ColumnDefinition { Name = "ClientEdit", Width = new GridLength(90) };
                    ColumnDefinition ClientDeleteColumn = new ColumnDefinition { Name = "ClientDelete", Width = new GridLength(90) };
                    RowDefinition ClientRow = new RowDefinition { Name = "ClientRow1", Height = new GridLength(40) };
                    ClientList.ColumnDefinitions.Add(ClientNameColumn);
                    ClientList.ColumnDefinitions.Add(ClientAddressColumn);
                    ClientList.ColumnDefinitions.Add(ClientEditColumn);
                    ClientList.ColumnDefinitions.Add(ClientDeleteColumn);
                    ClientList.RowDefinitions.Add(ClientRow);

                    AddRoute(currentRoute);
                }
            }
            SerializeRoutes();
        }

        private void DeleteRouteButton_Click(object sender, RoutedEventArgs e)
        {
            DeleteRouteConfirmation deleteRoute = new DeleteRouteConfirmation();
            deleteRoute.ShowDialog();
            if(deleteRoute.Confirmation == true)
            {
                RemoveRoute(currentRoute);
            }
            SerializeRoutes();
        }

        private void AddRoute(Route routeToAdd)
        {
            ClientList.Children.Clear();
            AddCreateNewClientButton();
            routeList.Add(routeToAdd);
            RouteSelectionBox.SelectedItem = currentRoute.Name;
            RepopulateRouteList();
        }

        private void RemoveRoute(Route routeToRemove)
        {
            ClientList.Children.Clear();
            ClientList.ColumnDefinitions.Clear();
            ClientList.RowDefinitions.Clear();
            routeList.Remove(routeToRemove);
            RepopulateRouteList();
            routeList.Remove(currentRoute);
            currentRoute = null;
            RouteSelectionBox.SelectedItem = null;
            SerializeRoutes();
        }

        private void RepopulateRouteList()
        {
            RouteSelectionBox.SelectionChanged -= RouteSelectionBox_SelectionChanged;
            RouteSelectionBox.Items.Clear();
            foreach (Route route in routeList)
            {
                RouteSelectionBox.Items.Add(route.Name);
            }
            if(currentRoute != null)
            {
                RouteSelectionBox.SelectedItem = currentRoute.Name;
            }
            RouteSelectionBox.SelectionChanged += RouteSelectionBox_SelectionChanged;
        }

        private void NewClientButton_Click(object sender, RoutedEventArgs e)
        {
            Client newClient = Client.NewClient();
            if(newClient != null)
            {
                AppendListNewClient(newClient);
                currentRoute.Clients.Add(newClient);
                AddCreateNewClientButton();
            }
            SerializeRoutes();
        }

        private void AppendListExistingClient(Client createClient, int position)
        {
            ColumnDefinition ClientNameColumn = new ColumnDefinition { Name = "ClientName", Width = new GridLength(360) };
            ColumnDefinition ClientAddressColumn = new ColumnDefinition { Name = "ClientAddress", Width = new GridLength(360) };
            ColumnDefinition ClientEditColumn = new ColumnDefinition { Name = "ClientEdit", Width = new GridLength(90) };
            ColumnDefinition ClientDeleteColumn = new ColumnDefinition { Name = "ClientDelete", Width = new GridLength(90) };
            RowDefinition ClientRow = new RowDefinition { Name = "Row" + position, Height = new GridLength(40) };
            ClientList.ColumnDefinitions.Add(ClientNameColumn);
            ClientList.ColumnDefinitions.Add(ClientAddressColumn);
            ClientList.ColumnDefinitions.Add(ClientEditColumn);
            ClientList.ColumnDefinitions.Add(ClientDeleteColumn);
            ClientList.RowDefinitions.Add(ClientRow);

            TextBlock newClientNameText = new TextBlock { Text = createClient.Name };
            newClientNameText.VerticalAlignment = VerticalAlignment.Center;
            newClientNameText.HorizontalAlignment = HorizontalAlignment.Left;
            newClientNameText.FontSize = 20;
            Grid.SetColumn(newClientNameText, 0);
            Grid.SetRow(newClientNameText, position);
            ClientList.Children.Add(newClientNameText);

            TextBlock newClientAddressText = new TextBlock { Text = createClient.Address };
            newClientAddressText.VerticalAlignment = VerticalAlignment.Center;
            newClientAddressText.HorizontalAlignment = HorizontalAlignment.Left;
            newClientAddressText.FontSize = 20;
            Grid.SetColumn(newClientAddressText, 1);
            Grid.SetRow(newClientAddressText, position);
            ClientList.Children.Add(newClientAddressText);

            Button editClientButton = new Button();
            editClientButton.Tag = createClient.Name;
            editClientButton.Click += EditClient;
            editClientButton.Content = "EDIT";
            Grid.SetColumn(editClientButton, 2);
            Grid.SetRow(editClientButton, position);
            ClientList.Children.Add(editClientButton);

            Button deleteClientButton = new Button();
            deleteClientButton.Tag = createClient.Name;
            deleteClientButton.Click += DeleteClient;
            deleteClientButton.Content = "X";
            Grid.SetColumn(deleteClientButton, 3);
            Grid.SetRow(deleteClientButton, position);
            ClientList.Children.Add(deleteClientButton);
        }

        private void AppendListNewClient(Client createClient)
        {
            ColumnDefinition ClientNameColumn = new ColumnDefinition { Name = "ClientName", Width = new GridLength(360) };
            ColumnDefinition ClientAddressColumn = new ColumnDefinition { Name = "ClientAddress", Width = new GridLength(360) };
            ColumnDefinition ClientEditColumn = new ColumnDefinition { Name = "ClientEdit", Width = new GridLength(90) };
            ColumnDefinition ClientDeleteColumn = new ColumnDefinition { Name = "ClientDelete", Width = new GridLength(90) };
            RowDefinition ClientRow = new RowDefinition { Name = "Row" + currentRoute.Clients.Count, Height = new GridLength(40) };
            ClientList.ColumnDefinitions.Add(ClientNameColumn);
            ClientList.ColumnDefinitions.Add(ClientAddressColumn);
            ClientList.ColumnDefinitions.Add(ClientEditColumn);
            ClientList.ColumnDefinitions.Add(ClientDeleteColumn);
            ClientList.RowDefinitions.Add(ClientRow);

            TextBlock newClientNameText = new TextBlock { Text = createClient.Name };
            newClientNameText.VerticalAlignment = VerticalAlignment.Center;
            newClientNameText.HorizontalAlignment = HorizontalAlignment.Left;
            newClientNameText.FontSize = 20;
            Grid.SetColumn(newClientNameText, 0);
            Grid.SetRow(newClientNameText, currentRoute.Clients.Count);
            ClientList.Children.Add(newClientNameText);

            TextBlock newClientAddressText = new TextBlock { Text = createClient.Address };
            newClientAddressText.VerticalAlignment = VerticalAlignment.Center;
            newClientAddressText.HorizontalAlignment = HorizontalAlignment.Left;
            newClientAddressText.FontSize = 20;
            Grid.SetColumn(newClientAddressText, 1);
            Grid.SetRow(newClientAddressText, currentRoute.Clients.Count);
            ClientList.Children.Add(newClientAddressText);

            Button editClientButton = new Button();
            editClientButton.Tag = createClient.Name;
            editClientButton.Click += EditClient;
            editClientButton.Content = "EDIT";
            Grid.SetColumn(editClientButton, 2);
            Grid.SetRow(editClientButton, currentRoute.Clients.Count);
            ClientList.Children.Add(editClientButton);

            Button deleteClientButton = new Button();
            deleteClientButton.Tag = createClient.Name;
            deleteClientButton.Click += DeleteClient;
            deleteClientButton.Content = "X";
            Grid.SetColumn(deleteClientButton, 3);
            Grid.SetRow(deleteClientButton, currentRoute.Clients.Count);
            ClientList.Children.Add(deleteClientButton);
        }
        private void DeleteClient(object sender, RoutedEventArgs e)
        {
            //Remove Client data from Route List
            Client deleteClient = currentRoute.Clients.Find(x => x.Name == (string)((Button)sender).Tag);
            currentRoute.Clients.Remove(deleteClient);

            ClientList.ColumnDefinitions.Clear();
            ClientList.RowDefinitions.Clear();
            ClientList.Children.Clear();

            ColumnDefinition ClientNameColumn = new ColumnDefinition { Name = "ClientName", Width = new GridLength(360) };
            ColumnDefinition ClientAddressColumn = new ColumnDefinition { Name = "ClientAddress", Width = new GridLength(360) };
            ColumnDefinition ClientEditColumn = new ColumnDefinition { Name = "ClientEdit", Width = new GridLength(90) };
            ColumnDefinition ClientDeleteColumn = new ColumnDefinition { Name = "ClientDelete", Width = new GridLength(90) };
            RowDefinition ClientRow = new RowDefinition { Name = "ClientRow1", Height = new GridLength(40) };
            ClientList.ColumnDefinitions.Add(ClientNameColumn);
            ClientList.ColumnDefinitions.Add(ClientAddressColumn);
            ClientList.ColumnDefinitions.Add(ClientEditColumn);
            ClientList.ColumnDefinitions.Add(ClientDeleteColumn);
            ClientList.RowDefinitions.Add(ClientRow);

            int position = 0;
            foreach (Client client in currentRoute.Clients)
            {
                AppendListExistingClient(client, position++);
            }
            AddCreateNewClientButton();
            SerializeRoutes();
        }

        private void EditClient(object sender, RoutedEventArgs e)
        {
            Client editClient = currentRoute.Clients.Find(x => x.Name == (string)((Button)sender).Tag);
            NewClientWindow editClientWindow = new NewClientWindow();
            editClientWindow.ClientNameTextBox.Text = editClient.Name;
            editClientWindow.ClientAddressTextBox.Text = editClient.Address;
            editClientWindow.SpecialInstructionsTextBox.Text = editClient.SpecialInstructions;
            editClientWindow.HotMealsTextBox.Text = Convert.ToString(editClient.HotMeals);
            editClientWindow.ColdMealsTextBox.Text = Convert.ToString(editClient.ColdMeals);
            editClientWindow.PhoneNumberTextBox.Text = editClient.PhoneNumber;
            editClientWindow.ShowDialog();
            editClient.Name = editClientWindow.ClientNameTextBox.Text;
            editClient.Address = editClientWindow.ClientAddressTextBox.Text;
            editClient.SpecialInstructions = editClientWindow.SpecialInstructionsTextBox.Text;
            editClient.HotMeals = Convert.ToInt32(editClientWindow.HotMealsTextBox.Text);
            editClient.ColdMeals = Convert.ToInt32(editClient.ColdMeals);
            editClient.PhoneNumber = editClientWindow.PhoneNumberTextBox.Text;

            ClientList.ColumnDefinitions.Clear();
            ClientList.RowDefinitions.Clear();
            ClientList.Children.Clear();

            ColumnDefinition ClientNameColumn = new ColumnDefinition { Name = "ClientName", Width = new GridLength(360) };
            ColumnDefinition ClientAddressColumn = new ColumnDefinition { Name = "ClientAddress", Width = new GridLength(360) };
            ColumnDefinition ClientEditColumn = new ColumnDefinition { Name = "ClientEdit", Width = new GridLength(90) };
            ColumnDefinition ClientDeleteColumn = new ColumnDefinition { Name = "ClientDelete", Width = new GridLength(90) };
            RowDefinition ClientRow = new RowDefinition { Name = "ClientRow1", Height = new GridLength(40) };
            ClientList.ColumnDefinitions.Add(ClientNameColumn);
            ClientList.ColumnDefinitions.Add(ClientAddressColumn);
            ClientList.ColumnDefinitions.Add(ClientEditColumn);
            ClientList.ColumnDefinitions.Add(ClientDeleteColumn);
            ClientList.RowDefinitions.Add(ClientRow);

            int position = 0;
            foreach (Client client in currentRoute.Clients)
            {
                AppendListExistingClient(client, position++);
            }
            AddCreateNewClientButton();
            SerializeRoutes();
        }

        private void AddCreateNewClientButton()
        {
            if (ClientList.Children.Contains(createNewClientButton))
            {
                ClientList.Children.Remove(createNewClientButton);
            }
            createNewClientButton = new Button();
            createNewClientButton.Height = 40;
            createNewClientButton.Content = "Create New Client";
            Grid.SetColumn(createNewClientButton, 0);
            Grid.SetRow(createNewClientButton, currentRoute.Clients.Count);
            createNewClientButton.Click += NewClientButton_Click;
            ClientList.Children.Add(createNewClientButton);
        }

        private void RouteSelectionBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ClientList.ColumnDefinitions.Clear();
            ClientList.RowDefinitions.Clear();
            ClientList.Children.Clear();

            //reformat for route
            ColumnDefinition ClientNameColumn = new ColumnDefinition { Name = "ClientName", Width = new GridLength(360) };
            ColumnDefinition ClientAddressColumn = new ColumnDefinition { Name = "ClientAddress", Width = new GridLength(360) };
            ColumnDefinition ClientEditColumn = new ColumnDefinition { Name = "ClientEdit", Width = new GridLength(90) };
            ColumnDefinition ClientDeleteColumn = new ColumnDefinition { Name = "ClientDelete", Width = new GridLength(90) };
            RowDefinition ClientRow = new RowDefinition { Name = "ClientRow1", Height = new GridLength(40) };
            ClientList.ColumnDefinitions.Add(ClientNameColumn); 
            ClientList.ColumnDefinitions.Add(ClientAddressColumn);
            ClientList.ColumnDefinitions.Add(ClientEditColumn);
            ClientList.ColumnDefinitions.Add(ClientDeleteColumn);
            ClientList.RowDefinitions.Add(ClientRow);

            currentRoute = routeList.Find(x => x.Name == (string)((ComboBox)sender).SelectedItem);
            int position = 0;
            foreach(Client client in currentRoute.Clients)
            {
                AppendListExistingClient(client, position++);
            }
            AddCreateNewClientButton();
        }

        private void GoogleMapsButton_Click(object sender, RoutedEventArgs e)
        {
            LoadingText.Text = "Loading...";
            LoadingText.Visibility = Visibility.Visible;
            try
            {
                WriteOutputToExcel(currentRoute.GenerateGoogleMapsRoute());
                ClientList.ColumnDefinitions.Clear();
                ClientList.RowDefinitions.Clear();
                ClientList.Children.Clear();

                ColumnDefinition ClientNameColumn = new ColumnDefinition { Name = "ClientName", Width = new GridLength(360) };
                ColumnDefinition ClientAddressColumn = new ColumnDefinition { Name = "ClientAddress", Width = new GridLength(360) };
                ColumnDefinition ClientEditColumn = new ColumnDefinition { Name = "ClientEdit", Width = new GridLength(90) };
                ColumnDefinition ClientDeleteColumn = new ColumnDefinition { Name = "ClientDelete", Width = new GridLength(90) };
                RowDefinition ClientRow = new RowDefinition { Name = "ClientRow1", Height = new GridLength(40) };
                ClientList.ColumnDefinitions.Add(ClientNameColumn);
                ClientList.ColumnDefinitions.Add(ClientAddressColumn);
                ClientList.ColumnDefinitions.Add(ClientEditColumn);
                ClientList.ColumnDefinitions.Add(ClientDeleteColumn);
                ClientList.RowDefinitions.Add(ClientRow);

                int position = 0;
                foreach (Client client in currentRoute.Clients)
                {
                    AppendListExistingClient(client, position++);
                }
                AddCreateNewClientButton();
            }
            catch (Exception E)
            {
                LoadingText.Text = E.Message;
                LoadingText.Visibility = Visibility.Visible;
            }
        }

        private void WriteOutputToExcel(GoogleMapsJSON.Route writtenRoute)
        {
            int clientNameIndex = 1,
                clientAddressIndex = 2,
                coldMealsCountIndex = 3,
                hotMealsCountIndex = 4,
                specialInstructionsIndex = 5,
                phoneNumberIndex = 6,
                directionsIndex = 7;

            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
            app.Visible = false;
            Microsoft.Office.Interop.Excel._Workbook appWB = app.Workbooks.Add();
            Microsoft.Office.Interop.Excel._Worksheet appSheet = appWB.ActiveSheet;

            appSheet.Cells[1, clientNameIndex] = "Client Name";
            appSheet.Cells[1, clientAddressIndex] = "Address";
            appSheet.Cells[1, coldMealsCountIndex] = "# Cold Meals";
            appSheet.Cells[1, hotMealsCountIndex] = "# Hot Meals";
            appSheet.Cells[1, specialInstructionsIndex] = "Special Instructions";
            appSheet.Cells[1, phoneNumberIndex] = "Phone Number";
            appSheet.Cells[1, directionsIndex] = "Directions";

            appSheet.Columns[clientNameIndex].ColumnWidth = 20;
            appSheet.Columns[directionsIndex].ColumnWidth = 70;
            appSheet.Columns[clientAddressIndex].ColumnWidth = 30;
            appSheet.Columns[coldMealsCountIndex].ColumnWidth = 12;
            appSheet.Columns[hotMealsCountIndex].ColumnWidth = 12;
            appSheet.Columns[specialInstructionsIndex].ColumnWidth = 20;
            appSheet.Columns[phoneNumberIndex].ColumnWidth = 15;

            appSheet.UsedRange.Style.VerticalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            appSheet.UsedRange.Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            //appSheet.Columns[VerticalAlignment = VerticalAlignment.Top;

            int rowIndex = 2;
            foreach(Client client in currentRoute.Clients)
            {
                appSheet.Cells[rowIndex, clientNameIndex] = client.Name;
                appSheet.Cells[rowIndex, clientAddressIndex] = client.Address;
                appSheet.Cells[rowIndex, coldMealsCountIndex] = client.ColdMeals;
                appSheet.Cells[rowIndex, hotMealsCountIndex] = client.HotMeals;
                appSheet.Cells[rowIndex, specialInstructionsIndex] = client.SpecialInstructions;
                appSheet.Cells[rowIndex, phoneNumberIndex] = client.PhoneNumber;
                rowIndex++;
            }

            rowIndex = 2;
            foreach(Leg leg in writtenRoute.legs)
            {
                StringBuilder dirString = new StringBuilder();
                bool first = true;
                foreach(Step step in leg.steps)
                {
                    if (!first)
                    {
                        dirString.Append("\r\n");
                    }
                    first = false;
                    StringWriter writer = new StringWriter();
                    HttpUtility.HtmlDecode(step.html_instructions, writer);
                    string encodedString = writer.ToString();
                    string finalString = Regex.Replace(encodedString, @"<(.|\n)*?>", "");
                    dirString.Append(finalString + ".");
                }
                appSheet.Cells[rowIndex, directionsIndex] = dirString.ToString();
                rowIndex++;;
            }
            app.Visible = true;
            LoadingText.Visibility = Visibility.Hidden;
            appWB.SaveAs(currentRoute.Name + ".xls");
        }
    }
}
