using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApplication2
{
    class Requirement
    {
        private Location location;
        //private List<InnerRequirement> InnerRequirements;
        private String value;
        private string p;
        private string title;

        public Requirement(String value, Location location)
        {
            this.location = location;
            this.value = value;
        }

        public Requirement(String value, Location loc, String title)
        {
            // TODO: Complete member initialization
            this.value = value;
            this.location = loc;
            this.title = title;
        }

        public void setUpdateTitle()
        {
            string[] arr = value.Split('.'); // split value up based on decimal points "."
            title = arr[arr.Length-1];// select last part

        }

        /* public void setInnerRequirements(List<InnerRequirement> reqs)
         {
             this.InnerRequirements = reqs;
         }

         public List<InnerRequirement> getRequirements()
         {
             return this.InnerRequirements;
         }
         */
        public Location getLocation()
        {
            return this.location;
        }

        public String getValue()
        {
            return this.value;
        }

        public String getTitle()
        {
            return this.title;
        }
    }
}
