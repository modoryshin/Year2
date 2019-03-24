using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Year2
{
    public class User
    {
        public string fullname;
        public string role;
        public string FullName
        {
            get { return fullname; }
            set { fullname = value; }
        }
        public string Role
        {
            get { return role; }
            set { role = value; }
        }
        public User(string f,string r)
        {
            FullName = f;
            Role = r;
        }
        public string Out()
        {
            return this.FullName + "|" + this.Role; 
        }
    }
}
