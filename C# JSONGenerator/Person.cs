using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JSONGenerator
{
    public class Person
    {
        public Person()
        {
            this.Email = new List<string>();
        }

        public string Name { get; set; }

        public string Surname { get; set; }

        public List<string> Email { get; set; }

        public int Age { get; set; }

        public bool Married { get; set; }
    }

}
