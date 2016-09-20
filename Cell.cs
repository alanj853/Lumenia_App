using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApplication2
{
    class Cell
    {
        private Location location;
        private string value = "";

        public Cell(Location l, String value)
        {
            this.location = l;
            this.value = value;
        }
    }
}
