using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Web;
using static Meals_On_Wheels.GoogleMapsJSON;

namespace Meals_On_Wheels
{
    [Serializable]
    class Route
    {
        public string Name { get; set; }
        public int Length { get; set; }
        public List<Client> Clients { get; set; }

        public static Route NewRoute()
        {
            NewRouteWindow win = new NewRouteWindow();
            win.ShowDialog();
            try
            {
                Route newRoute = new Route { Name = win.RouteNameTextBox.Text, Clients = new List<Client>() };
                return newRoute;
            }
            catch
            {
                System.Windows.MessageBox.Show("Error creating route. Please confirm you are entering valid data and try again");
                return null;
            }

        }

        public GoogleMapsJSON.Route GenerateGoogleMapsRoute()
        {
            //origin=2550+Honey+Creek+Circle+East+Troy&destination=1752+Gateway+Boulevard+Beloit&
            StringBuilder sb = new StringBuilder();
            sb.Append("https://maps.googleapis.com/maps/api/directions/json?key=" + System.Configuration.ConfigurationManager.AppSettings.Get("APIKey") +"&origin=424+College+St+Beloit+WI+53511&destination=424+College+St+Beloit+WI+53511");
            if (Clients.Count > 0)
            {
                sb.Append("&waypoints=optimize:true|");
            }
            foreach (Client client in Clients)
            {
                sb.Append(HttpUtility.UrlEncode(client.Address) + "|");
            }
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(sb.ToString());
            HttpWebResponse response = (HttpWebResponse)request.GetResponse();
            string jsonResponse = new StreamReader(response.GetResponseStream()).ReadToEnd();

            Rootobject googleResponse = JsonConvert.DeserializeObject<Rootobject>(jsonResponse);
            if(googleResponse.status != "OK")
            {
                throw new FormatException("Bad results from Google");
            }
            GoogleMapsJSON.Route primaryRoute = googleResponse.routes[0];

            //re-order clients in this route for optimal driving
            List<Client> orderedClientList = new List<Client>();
            foreach (int order in primaryRoute.waypoint_order)
            {
                orderedClientList.Add(Clients[order]);
            }
            Clients = orderedClientList;

            return primaryRoute;
        }
    }
}
