using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JSONGenerator
{
    public class PersonGenerator
    {
        private static Random random = new Random((int)DateTime.Now.Ticks);//thanks to McAden

        public List<Person> Persons { get; private set; }

        public PersonGenerator(int nrOfPersons)
        {

            this.Persons = new List<Person>();

            for (int i = 1; i <= nrOfPersons; i++)
            {
                var person = new Person();
                person.Name = this.RandomString(8);
                person.Surname = this.RandomString(16);
                person.Age = i;
                person.Married = (i % 2 != 0);

                for (int j = 1; j <= 3; j++)
                {
                    person.Email.Add(this.RandomString(3) + "@" + this.RandomString(8) + ".com");
                }

                this.Persons.Add(person);
            }
        }

        private string RandomString(int size)
        {
            StringBuilder builder = new StringBuilder();
            char ch;
            for (int i = 0; i < size; i++)
            {
                ch = Convert.ToChar(Convert.ToInt32(Math.Floor(26 * random.NextDouble() + 65)));
                builder.Append(ch);
            }

            return builder.ToString();
        }
    }
}
