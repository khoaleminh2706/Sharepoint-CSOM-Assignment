using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;

namespace CreateSPSite.Models
{
    public class ProjectContentType : AbstractContentType
    {
        public ProjectContentType(ClientContext clientContext) : base(clientContext)
        {
            Name = "Project";
            FieldsList = new List<AbstractField>
            {
                new NewColumnField(clientContext)
                {
                    InternalName = "ProjectName",
                    DisplayName = "Project Name",
                    XmlSchema = $"<Field ID='{Guid.NewGuid()}' Type='Text' Name='Project Name' StaticName='ProjectName' DisplayName='Project Name' />"
                },
                new NewColumnField(clientContext)
                {
                    InternalName = "ProjDescription",
                    DisplayName = "Project Description",
                    XmlSchema = $"<Field ID='{Guid.NewGuid()}' Type='Note' Name='ProjDescription' StaticName='ProjDescription' DisplayName='Description' NumLines='6' RichText='FALSE' Sortable='FALSE' />"
                },
                new NewColumnField(clientContext)
                {
                    InternalName = "State",
                    DisplayName = "State",
                    XmlSchema = $"<Field ID='{Guid.NewGuid()}' Type='Choice' Name='State' StaticName='State' DisplayName='State' Format='Dropdown'><CHOICES>" +
                    $"<CHOICE>Signed</CHOICE>" +
                    $"<CHOICE>Design</CHOICE>" +
                    $"<CHOICE>Development</CHOICE>" +
                    $"<CHOICE>Maintenance</CHOICE>" +
                    $"<CHOICE>Closed</CHOICE>" +
                    $"</CHOICES></Field>"
                },
                new SiteColumnField(clientContext)
                {
                    InternalName = "StartDate"
                },
                new SiteColumnField(clientContext) { InternalName = "_EndDate" }
            };
        }
    }
}
