using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GymManagementSystem
{
    public class Client
    {
        public string FullName { get; set; }
        public string PhoneNumber { get; set; }
        public double Weight { get; set; }
        public string SubscriptionType { get; set; }
        public DateTime SubscriptionStart { get; set; }
        public DateTime SubscriptionEnd { get; set; }
        public int Sessions { get; set; }
        public bool IsFrozen { get; set; }
    }
}