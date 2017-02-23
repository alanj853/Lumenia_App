using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApplication2
{
    class JointHeading
    {
        private Boolean requirement = false;
        private Boolean subSubHeading = false;
        private Location location;
        private String value;
        private int numericValue = 0;
        private String title = "No title assigned";
        private SubSubHeading ssh;


        public JointHeading(String value, Location loc, String title,Boolean requirement, Boolean subSubHeading, SubSubHeading ssh)
        {
            // TODO: Complete member initialization
            this.value = value;
            this.location = loc;
            this.title = title;
            this.requirement = requirement;
            this.subSubHeading = subSubHeading;

            String temp = ""; // temp to store 'value'
            temp = value.Replace(".",String.Empty);
            this.numericValue = Int32.Parse(temp);

            if (subSubHeading)
                this.ssh = ssh;

        }

        public SubSubHeading getSubSubHeading() {
            return this.ssh;
        }

        public Location getLocation()
        {
            return this.location;
        }

        public int getNumbericValue() {
            return this.numericValue;
        }

        public String getValue()
        {
            return this.value;
        }

        public string getTitle()
        {
            return title;
        }

        public Boolean isRequirement() {
            return requirement;
        }

        public Boolean isSubSubHeading()
        {
            return subSubHeading;
        }
    }
}
