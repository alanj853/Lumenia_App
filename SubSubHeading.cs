using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ConsoleApplication2
{
    class SubSubHeading
    {
        private Location location;
        private List<Requirement> requirements = new List<Requirement>();
        private String value;
        private Boolean averageOfRequirementsAssigned = false;
        private String averageOfRequirements = "";
        private String title = "No title assigned";
        private Boolean isRequirement = false;

        public SubSubHeading(String value, Location location)
        {
            this.location = location;
            this.value = value;
        }

        public SubSubHeading(String value, Location location, String title)
        {
            this.value = value;
            this.location = location;
            this.title = title;
        }

        public void addRequirementToList(Requirement req)
        {
            this.requirements.Add(req);
        }

        public void setRequirements(List<Requirement> reqs)
        {
            this.requirements = reqs;
        }

        public List<Requirement> getRequirements()
        {
            return this.requirements;
        }

        public Location getLocation()
        {
            return this.location;
        }

        public String getValue()
        {
            return this.value;
        }

        public Boolean isAlsoRequirement() {
            return this.isRequirement;
        }

        public void setIsAlsoRequirement() {
            this.isRequirement = true;
        }

        public String assignAverageForRequirements(int systemNo)
        {
            averageOfRequirements = "";
            for (int i = 0; i < requirements.Count; i++)
            {
                Console.WriteLine("This is Req: " + requirements[i].getValue());
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

        internal string getTitle()
        {
            return title;
        }
    }
}
