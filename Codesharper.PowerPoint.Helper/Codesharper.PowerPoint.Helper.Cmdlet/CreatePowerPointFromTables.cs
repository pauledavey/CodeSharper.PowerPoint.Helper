using System.Management.Automation;

namespace Codesharper.PowerPoint.Helper.Cmdlet
{
    using Cmdlet = System.Management.Automation.Cmdlet;

    [Cmdlet(VerbsCommunications.Send, "Greeting")]
    public class CreatePowerPointFromTables : Cmdlet
    {
        [Parameter(Mandatory = true)]
        public string Name
        {
            get
            {
                return name;
            }

            set
            {
                name = value;
            }
        }

        private string name;

        protected override void ProcessRecord()
        {
            WriteObject("Hello " + name + "!");
        }

    }
        }