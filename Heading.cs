using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace ConsoleApplication2
{
    class Heading
    {
        private String value;
        private List<SubHeading> subHeadings = new List<SubHeading>();
        private Location location;
        private String title = "No title assigned";

        public Heading(String value, Location location)
        {
            this.value = value;
            this.location = location;
        }

        public Heading(String value, Location location, String title)
        {
            this.value = value;
            this.location = location;
            this.title = title;
        }



        public String getValue()
        {
            return value;
        }

        public List<SubHeading> getSubHeadings()
        {
            return subHeadings;
        }

        public void setSubHeadings(List<SubHeading> l)
        {
            this.subHeadings = l;
        }

        public void addSubHeadingToList(SubHeading s){
            this.subHeadings.Add(s);
        }

        public Location getLocation()
        {
            return location;
        }



        public String getTitle()
        {
            return title;
        }
    }
}
