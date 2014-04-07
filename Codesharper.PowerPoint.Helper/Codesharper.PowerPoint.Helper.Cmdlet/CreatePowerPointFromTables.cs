using System.Management.Automation;
using System.Configuration.Install;

// How to debug a cmdlet
// http://www.sharepointjohn.com/powershell-debug-custom-csharp-powershell-cmdlet/


namespace Codesharper.PowerPoint.Helper.Cmdlet
{
    using System;
    using System.Collections;

    using Cmdlet = System.Management.Automation.Cmdlet;

    [Cmdlet(VerbsCommunications.Send, "Greeting")]
    public class CreatePowerPointFromTables : Cmdlet
    {
        [Parameter(Mandatory = true)]
        public PSObject[] objectIn
        {
            get;
            set;
        }

        protected override void ProcessRecord()
        {
            foreach (var entry in objectIn)

                
            {
                foreach (var innerEntry in entry.Properties)
                {
                    try
                    {
                        WriteObject(innerEntry.Name + " = " + innerEntry.Value);
                    }
                    catch (Exception)
                    {
                        
                    }
                   
                }

            }
        }
    }
}