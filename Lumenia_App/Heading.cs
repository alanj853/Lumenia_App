using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace Lumenia_App
{
    class Heading
    {
        private String value;
        private List<SubHeading> subHeadings = new List<SubHeading>();
        private List<Requirement> requirements = new List<Requirement>();
        private Location location;
        private String title = "No title assigned";

        private String averageOfRequirements = "";
        private Boolean averageOfRequirementsAssigned = false;

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

        public void addSubHeadingToList(SubHeading s)
        {
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

        public void setRequirements(List<Requirement> reqs)
        {
            this.requirements = reqs;
        }

        public List<Requirement> getRequirements()
        {
            return this.requirements;
        }

        public String assignAverageForRequirements(int systemNo)
        {
            averageOfRequirements = "";
            for (int i = 0; i < requirements.Count; i++)
            {
                int row = requirements[i].getLocation().getRow();
                int col = requirements[i].getLocation().getColumn() + 1 + systemNo;
                Location newLoc = new Location(row, col);
                if (i == (requirements.Count - 1))
                {
                    averageOfRequirements = "AVERAGE(" + averageOfRequirements + newLoc.getExcelAddress() + ") ";

                }
                else
                {
                    averageOfRequirements = averageOfRequirements + newLoc.getExcelAddress() + ", ";
                }

            }
            averageOfRequirements = "=IFERROR(" + averageOfRequirements + ",\"\")";
            averageOfRequirementsAssigned = true;
            return averageOfRequirements;
        }

        public String getAverageOfRequirements()
        {
            return averageOfRequirements;
        }

        public Boolean isAverageOfRequirementsAssigned()
        {
            return averageOfRequirementsAssigned;
        }

        public Boolean hasRequirements()
        {
            try
            {
                if (requirements[0] != null)
                    return true;
                return false;
            }
            catch (ArgumentOutOfRangeException ex)
            {
                string err = ex.StackTrace;
                return false;
            }
        }

        public void addRequirementToList(Requirement newRequirement)
        {
            this.requirements.Add(newRequirement);
        }
    }
}