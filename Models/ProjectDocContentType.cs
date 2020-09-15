using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;

namespace CreateSPSite.Models
{
    public class ProjectDocContentType : AbstractContentType
    {
        public ProjectDocContentType(ClientContext clientContext) : base(clientContext)
        {
            Name = "Project Document";
            ParentTypeTitle = "Document";
            FieldsList = new List<AbstractField>()
            {
                new NewColumnField(clientContext)
                {
                    InternalName = "DocDescription",
                    DisplayName = "Description",
                    XmlSchema = $"<Field ID='{Guid.NewGuid()}' Type='Note' Name='DocDescription' StaticName='DocDescription' DisplayName='Description' NumLines='6' RichText='FALSE' Sortable='FALSE' />"
                },
                new NewColumnField(clientContext)
                {
                    InternalName = "DocType",
                    DisplayName = "Document Type",
                    XmlSchema = $"<Field ID='{Guid.NewGuid()}' Type='Choice' Name='DocType' StaticName='DocType' DisplayName='Document Type' Format='Dropdown'><CHOICES>" +
                    $"<CHOICE>Business requirement</CHOICE>" +
                    $"<CHOICE>Technical document</CHOICE>" +
                    $"<CHOICE>User guide</CHOICE>" +
                    $"</CHOICES></Field>"
                }
            };
        }
    }
}
