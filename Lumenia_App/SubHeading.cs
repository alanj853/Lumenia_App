using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lumenia_App
{
    internal class SubHeading
    {
        private Location location;
        private List<SubSubHeading> subSubHeadings = new List<SubSubHeading>();
        private List<Requirement> requirements = new List<Requirement>();
        private String value;
        private String averageOfRequirements = "";
        private Boolean averageOfRequirementsAssigned = false;
        private String title = "No title assigned";

        public SubHeading(String value, Location location)
        {
            this.location = location;
            this.value = value;
        }

        public SubHeading(String value, Location location, String title)
        {
            this.value = value;
            this.location = location;
            this.title = title;
        }

        public void setSubSubHeadings(List<SubSubHeading> subsubHeadings)
        {
            this.subSubHeadings = subsubHeadings;
        }



        public List<SubSubHeading> getSubSubHeadings()
        {
            return this.subSubHeadings;
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

        public Location getLocation()
        {
            return this.location;
        }

        public String getValue()
        {
            return this.value;
        }

        public Boolean hasSubSubHeadings()
        {
            try
            {
                if (subSubHeadings[0] != null)
                    return true;
                return false;
            }
            catch (ArgumentOutOfRangeException ex)
            {
                string err = ex.StackTrace;
                return false;
            }
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

        public string getTitle()
        {
            return title;
        }

        public void addSubSubHeadingToList(SubSubHeading newSubSubHeading)
        {
            this.subSubHeadings.Add(newSubSubHeading);
        }

        public void addRequirementToList(Requirement newRequirement)
        {
            this.requirements.Add(newRequirement);
        }


    }
}